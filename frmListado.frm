VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7890
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCambioNroFactura 
      Height          =   3480
      Left            =   0
      TabIndex        =   627
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   131
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   630
         Top             =   1950
         Width           =   1095
      End
      Begin VB.CommandButton CmdAcepCambioNro 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   4
         Left            =   3120
         TabIndex        =   631
         Top             =   2700
         Width           =   1335
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   15
         Left            =   4590
         TabIndex        =   633
         Top             =   2700
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   129
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   628
         Top             =   930
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   130
         Left            =   3210
         MaxLength       =   10
         TabIndex        =   629
         Top             =   1410
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Factura"
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
         Index           =   185
         Left            =   720
         TabIndex        =   637
         Top             =   1950
         Width           =   1260
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   23
         Left            =   2865
         Picture         =   "frmListado.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1950
         Width           =   240
      End
      Begin VB.Label Label25 
         Caption         =   "Recepci�n de N�mero de Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   660
         TabIndex        =   635
         Top             =   300
         Width           =   4815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo N�mero de Factura"
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
         Index           =   18
         Left            =   720
         TabIndex        =   634
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Confirmar N�mero de Factura"
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
         Left            =   720
         TabIndex        =   632
         Top             =   1410
         Width           =   2130
      End
   End
   Begin VB.Frame FrameGeneraClasifica 
      Height          =   3390
      Left            =   0
      TabIndex        =   400
      Top             =   -30
      Width           =   6615
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   79
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   406
         Tag             =   "Porcentaje Bonificaci�n|N|N|||rbonifentradas|porcbonif|#,##0||"
         Top             =   2100
         Width           =   1035
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   80
         Left            =   1935
         MaxLength       =   7
         TabIndex        =   405
         Tag             =   "Porcentaje Bonificaci�n|N|N|||rbonifentradas|porcbonif|#,##0||"
         Top             =   1680
         Width           =   1035
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5130
         TabIndex        =   408
         Top             =   2445
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepGene 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4020
         TabIndex        =   407
         Top             =   2445
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   83
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   404
         Top             =   1275
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   83
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   403
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.CommandButton Command24 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":0097
         Style           =   1  'Graphical
         TabIndex        =   402
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command23 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":03A1
         Style           =   1  'Graphical
         TabIndex        =   401
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% Destrio"
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
         Index           =   110
         Left            =   660
         TabIndex        =   412
         Top             =   2100
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campo"
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
         Index           =   111
         Left            =   660
         TabIndex        =   411
         Top             =   1695
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "Generaci�n Clasificaci�n"
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
         TabIndex        =   410
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
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
         Index           =   114
         Left            =   660
         TabIndex        =   409
         Top             =   1290
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   39
         Left            =   1620
         MouseIcon       =   "frmListado.frx":06AB
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
   End
   Begin VB.Frame FrameControlDestrio 
      Height          =   6690
      Left            =   0
      TabIndex        =   421
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check14 
         Caption         =   "Imprimir Resumen"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4110
         TabIndex        =   536
         Top             =   5070
         Width           =   2025
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   90
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   441
         Top             =   5070
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   91
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   442
         Top             =   5445
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   88
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   438
         Top             =   4140
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   89
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   440
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5220
         TabIndex        =   446
         Top             =   6150
         Width           =   975
      End
      Begin VB.CommandButton cmdAcepCtrolDestrio 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   444
         Top             =   6150
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   86
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   430
         Top             =   1275
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   87
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   431
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   86
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   429
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   87
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   428
         Text            =   "Text5"
         Top             =   1635
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   84
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   434
         Top             =   3210
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   85
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   436
         Top             =   3570
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   84
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   427
         Text            =   "Text5"
         Top             =   3210
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   85
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   426
         Text            =   "Text5"
         Top             =   3570
         Width           =   3375
      End
      Begin VB.CommandButton Command30 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":07FD
         Style           =   1  'Graphical
         TabIndex        =   425
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command27 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":0B07
         Style           =   1  'Graphical
         TabIndex        =   424
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   81
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   432
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   82
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   433
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   81
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   423
         Text            =   "Text5"
         Top             =   2220
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   82
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   422
         Text            =   "Text5"
         Top             =   2580
         Width           =   3375
      End
      Begin MSComctlLib.ProgressBar Pb4 
         Height          =   255
         Left            =   330
         TabIndex        =   458
         Top             =   5850
         Visible         =   0   'False
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         Height          =   195
         Index           =   117
         Left            =   360
         TabIndex        =   459
         Top             =   6150
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campo"
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
         Index           =   115
         Left            =   660
         TabIndex        =   457
         Top             =   4860
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   113
         Left            =   990
         TabIndex        =   456
         Top             =   5100
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   112
         Left            =   990
         TabIndex        =   455
         Top             =   5445
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   130
         Left            =   675
         TabIndex        =   454
         Top             =   3960
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   129
         Left            =   1005
         TabIndex        =   453
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   128
         Left            =   1005
         TabIndex        =   452
         Top             =   4545
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   127
         Left            =   960
         TabIndex        =   451
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   126
         Left            =   960
         TabIndex        =   450
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   125
         Left            =   675
         TabIndex        =   449
         Top             =   3015
         Width           =   630
      End
      Begin VB.Label Label18 
         Caption         =   "Informe de Control Destrio"
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
         TabIndex        =   448
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   124
         Left            =   1005
         TabIndex        =   447
         Top             =   3255
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   123
         Left            =   1005
         TabIndex        =   445
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
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
         Index           =   122
         Left            =   675
         TabIndex        =   443
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   53
         Left            =   1620
         MouseIcon       =   "frmListado.frx":0E11
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   54
         Left            =   1620
         MouseIcon       =   "frmListado.frx":0F63
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   51
         Left            =   1620
         MouseIcon       =   "frmListado.frx":10B5
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   52
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1207
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   15
         Left            =   1620
         Picture         =   "frmListado.frx":1359
         ToolTipText     =   "Buscar fecha"
         Top             =   4545
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   14
         Left            =   1620
         Picture         =   "frmListado.frx":13E4
         ToolTipText     =   "Buscar fecha"
         Top             =   4140
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   121
         Left            =   675
         TabIndex        =   439
         Top             =   2025
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   120
         Left            =   1005
         TabIndex        =   437
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   119
         Left            =   1005
         TabIndex        =   435
         Top             =   2655
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   55
         Left            =   1620
         MouseIcon       =   "frmListado.frx":146F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   56
         Left            =   1620
         MouseIcon       =   "frmListado.frx":15C1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2610
         Width           =   240
      End
   End
   Begin VB.Frame FrameKilosProducto 
      Height          =   6480
      Left            =   0
      TabIndex        =   165
      Top             =   -30
      Width           =   6615
      Begin VB.CheckBox Check22 
         Caption         =   "Detalle"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4140
         TabIndex        =   865
         Top             =   5400
         Width           =   2085
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Salta p�gina Producto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4140
         TabIndex        =   200
         Top             =   5070
         Width           =   2085
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   178
         Top             =   3210
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   179
         Top             =   3570
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   37
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   196
         Text            =   "Text5"
         Top             =   3210
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   38
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   195
         Text            =   "Text5"
         Top             =   3570
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   182
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   180
         Top             =   4050
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelInf 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5190
         TabIndex        =   188
         Top             =   5745
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepInf 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   186
         Top             =   5745
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   174
         Top             =   1275
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   175
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   33
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   173
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   34
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   172
         Text            =   "Text5"
         Top             =   1635
         Width           =   3375
      End
      Begin VB.CommandButton Command10 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1713
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command9 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1A1D
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   176
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   177
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   35
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   169
         Text            =   "Text5"
         Top             =   2220
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   36
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   168
         Text            =   "Text5"
         Top             =   2580
         Width           =   3375
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   585
         TabIndex        =   166
         Top             =   4860
         Width           =   3480
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   5
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   184
            Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
            Top             =   135
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Hect�reas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   38
            Left            =   90
            TabIndex        =   167
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
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
         Index           =   45
         Left            =   675
         TabIndex        =   199
         Top             =   3015
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   44
         Left            =   1005
         TabIndex        =   198
         Top             =   3255
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   43
         Left            =   1005
         TabIndex        =   197
         Top             =   3645
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   37
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1D27
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   38
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1E79
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   50
         Left            =   675
         TabIndex        =   194
         Top             =   3870
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   1005
         TabIndex        =   193
         Top             =   4110
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   48
         Left            =   1005
         TabIndex        =   192
         Top             =   4455
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   960
         TabIndex        =   191
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   960
         TabIndex        =   190
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Informe Kilos por Producto"
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
         Left            =   660
         TabIndex        =   189
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
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
         Index           =   42
         Left            =   675
         TabIndex        =   187
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   33
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1FCB
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   34
         Left            =   1620
         MouseIcon       =   "frmListado.frx":211D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1620
         Picture         =   "frmListado.frx":226F
         ToolTipText     =   "Buscar fecha"
         Top             =   4455
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1620
         Picture         =   "frmListado.frx":22FA
         ToolTipText     =   "Buscar fecha"
         Top             =   4050
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   41
         Left            =   675
         TabIndex        =   185
         Top             =   2025
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   40
         Left            =   1005
         TabIndex        =   183
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   39
         Left            =   1005
         TabIndex        =   181
         Top             =   2655
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   35
         Left            =   1620
         MouseIcon       =   "frmListado.frx":2385
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   36
         Left            =   1620
         MouseIcon       =   "frmListado.frx":24D7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2610
         Width           =   240
      End
   End
   Begin VB.Frame FrameGastosCampos 
      Height          =   6720
      Left            =   0
      TabIndex        =   573
      Top             =   -60
      Width           =   6765
      Begin VB.Frame Frame7 
         Caption         =   "Clasificado por"
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
         Height          =   720
         Left            =   570
         TabIndex        =   605
         Top             =   5010
         Width           =   3690
         Begin VB.OptionButton Opcion1 
            Caption         =   "Socio"
            Height          =   345
            Index           =   5
            Left            =   480
            TabIndex        =   591
            Top             =   225
            Width           =   1290
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Concepto"
            Height          =   255
            Index           =   6
            Left            =   2160
            TabIndex        =   606
            Top             =   270
            Width           =   1320
         End
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Resumen"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   4920
         TabIndex        =   604
         Top             =   5130
         Width           =   1695
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   126
         Left            =   2115
         MaxLength       =   10
         TabIndex        =   588
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   127
         Left            =   2115
         MaxLength       =   10
         TabIndex        =   590
         Top             =   4650
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   5220
         TabIndex        =   594
         Top             =   6180
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepGtosCampos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4140
         TabIndex        =   592
         Top             =   6180
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   120
         Left            =   2115
         MaxLength       =   6
         TabIndex        =   580
         Text            =   "000000"
         Top             =   1275
         Width           =   900
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   121
         Left            =   2115
         MaxLength       =   6
         TabIndex        =   581
         Top             =   1605
         Width           =   900
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   120
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   579
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   121
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   578
         Text            =   "Text5"
         Top             =   1605
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   124
         Left            =   2115
         MaxLength       =   6
         TabIndex        =   584
         Top             =   3360
         Width           =   825
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   125
         Left            =   2115
         MaxLength       =   6
         TabIndex        =   586
         Top             =   3690
         Width           =   825
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   124
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   577
         Text            =   "Text5"
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   125
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   576
         Text            =   "Text5"
         Top             =   3690
         Width           =   3375
      End
      Begin VB.CommandButton Command38 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":2629
         Style           =   1  'Graphical
         TabIndex        =   575
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command35 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":2933
         Style           =   1  'Graphical
         TabIndex        =   574
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   122
         Left            =   2115
         MaxLength       =   8
         TabIndex        =   582
         Tag             =   "C�digo Campo|N|N|1|99999999|rcampos|codcampo|00000000|S|"
         Text            =   "00000000"
         Top             =   2280
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   123
         Left            =   2115
         MaxLength       =   8
         TabIndex        =   583
         Tag             =   "C�digo Campo|N|N|1|99999999|rcampos|codcampo|00000000|S|"
         Top             =   2610
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb6 
         Height          =   255
         Left            =   570
         TabIndex        =   607
         Top             =   5850
         Visible         =   0   'False
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         Height          =   255
         Index           =   182
         Left            =   570
         TabIndex        =   608
         Top             =   6180
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   181
         Left            =   675
         TabIndex        =   603
         Top             =   4080
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Index           =   180
         Left            =   1005
         TabIndex        =   602
         Top             =   4350
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   225
         Index           =   179
         Left            =   1005
         TabIndex        =   601
         Top             =   4665
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   178
         Left            =   960
         TabIndex        =   600
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   177
         Left            =   960
         TabIndex        =   599
         Top             =   1650
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Concepto Gasto"
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
         Index           =   176
         Left            =   675
         TabIndex        =   598
         Top             =   3075
         Width           =   1155
      End
      Begin VB.Label Label23 
         Caption         =   "Informe de Gastos por Campo"
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
         TabIndex        =   597
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   175
         Left            =   1005
         TabIndex        =   596
         Top             =   3375
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   174
         Left            =   1005
         TabIndex        =   595
         Top             =   3735
         Width           =   630
      End
      Begin VB.Label Label2 
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
         Index           =   173
         Left            =   675
         TabIndex        =   593
         Top             =   1050
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   90
         Left            =   1800
         MouseIcon       =   "frmListado.frx":2C3D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   91
         Left            =   1800
         MouseIcon       =   "frmListado.frx":2D8F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   94
         Left            =   1800
         MouseIcon       =   "frmListado.frx":2EE1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   95
         Left            =   1800
         MouseIcon       =   "frmListado.frx":3033
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   22
         Left            =   1800
         Picture         =   "frmListado.frx":3185
         ToolTipText     =   "Buscar fecha"
         Top             =   4650
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   21
         Left            =   1800
         Picture         =   "frmListado.frx":3210
         ToolTipText     =   "Buscar fecha"
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campo"
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
         Index           =   157
         Left            =   675
         TabIndex        =   589
         Top             =   2025
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   156
         Left            =   975
         TabIndex        =   587
         Top             =   2325
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   155
         Left            =   975
         TabIndex        =   585
         Top             =   2655
         Width           =   780
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   92
         Left            =   1800
         MouseIcon       =   "frmListado.frx":329B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar campo"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   93
         Left            =   1800
         MouseIcon       =   "frmListado.frx":33ED
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar campo"
         Top             =   2640
         Width           =   240
      End
   End
   Begin VB.Frame FrameContabGastos 
      Height          =   5220
      Left            =   0
      TabIndex        =   609
      Top             =   -30
      Width           =   6615
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   108
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   626
         Text            =   "fecha gto"
         Top             =   1020
         Visible         =   0   'False
         Width           =   1230
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
         Index           =   119
         Left            =   1365
         MaxLength       =   30
         TabIndex        =   611
         Top             =   2400
         Width           =   4875
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
         Index           =   112
         Left            =   1365
         MaxLength       =   2
         TabIndex        =   610
         Tag             =   "C�digo Campo|N|N|1|99999999|rcampos|codcampo|00|S|"
         Top             =   1650
         Width           =   945
      End
      Begin VB.CommandButton Command41 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":353F
         Style           =   1  'Graphical
         TabIndex        =   618
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command40 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":3849
         Style           =   1  'Graphical
         TabIndex        =   617
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
         Index           =   128
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   616
         Text            =   "Text5"
         Top             =   3120
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
         Index           =   128
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   612
         Top             =   3120
         Width           =   1455
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
         Index           =   112
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   615
         Text            =   "Text5"
         Top             =   1650
         Width           =   3885
      End
      Begin VB.CommandButton CmdAcepContaGastos 
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
         Left            =   4170
         TabIndex        =   613
         Top             =   4500
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
         Index           =   14
         Left            =   5250
         TabIndex        =   614
         Top             =   4500
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   255
         Left            =   600
         TabIndex        =   619
         Top             =   3810
         Visible         =   0   'False
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Concepto Contable"
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
         Index           =   193
         Left            =   420
         TabIndex        =   624
         Top             =   1380
         Width           =   1890
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   81
         Left            =   1020
         MouseIcon       =   "frmListado.frx":3B53
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cuenta"
         Top             =   3150
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   80
         Left            =   1020
         MouseIcon       =   "frmListado.frx":3CA5
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label24 
         Caption         =   "Contabilizaci�n de Gastos al Diario"
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
         TabIndex        =   623
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ampliaci�n de Concepto"
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
         Index           =   189
         Left            =   420
         TabIndex        =   622
         Top             =   2130
         Width           =   2370
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cta Contrapartida"
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
         Index           =   184
         Left            =   420
         TabIndex        =   621
         Top             =   2880
         Width           =   1770
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   183
         Left            =   600
         TabIndex        =   620
         Top             =   4170
         Visible         =   0   'False
         Width           =   5535
      End
   End
   Begin VB.Frame FrameAsignacionGlobalgap 
      Height          =   3135
      Left            =   0
      TabIndex        =   867
      Top             =   0
      Width           =   6090
      Begin MSComctlLib.ProgressBar pb10 
         Height          =   225
         Left            =   540
         TabIndex        =   874
         Top             =   1740
         Visible         =   0   'False
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
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
         Index           =   26
         Left            =   4650
         TabIndex        =   872
         Top             =   2340
         Width           =   975
      End
      Begin VB.CommandButton cmdAcepAsigGlobalgap 
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
         Left            =   3600
         TabIndex        =   871
         Top             =   2340
         Width           =   975
      End
      Begin VB.CommandButton Command58 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":3DF7
         Style           =   1  'Graphical
         TabIndex        =   870
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command57 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":4101
         Style           =   1  'Graphical
         TabIndex        =   869
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CheckBox Check25 
         Caption         =   "Limpiar los c�digos que no existan"
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
         Left            =   570
         TabIndex        =   868
         Top             =   1170
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   262
         Left            =   570
         TabIndex        =   876
         Top             =   2490
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   238
         Left            =   570
         TabIndex        =   875
         Top             =   2250
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label35 
         Caption         =   "Asignaci�n de c�digos Globalgap"
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
         Left            =   540
         TabIndex        =   873
         Top             =   405
         Width           =   5025
      End
   End
   Begin VB.Frame FrameInformeFases 
      Height          =   3390
      Left            =   0
      TabIndex        =   413
      Top             =   0
      Width           =   6615
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
         Index           =   12
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   419
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   1440
         Width           =   2070
      End
      Begin VB.CommandButton Command29 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":440B
         Style           =   1  'Graphical
         TabIndex        =   417
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command28 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":4715
         Style           =   1  'Graphical
         TabIndex        =   416
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton CmdAcepInfFases 
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
         Left            =   4020
         TabIndex        =   415
         Top             =   2445
         Width           =   1005
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
         Left            =   5130
         TabIndex        =   414
         Top             =   2445
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Fase"
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
         Index           =   116
         Left            =   630
         TabIndex        =   420
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label Label17 
         Caption         =   "Socios agrupados por Fases"
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
         TabIndex        =   418
         Top             =   420
         Width           =   5025
      End
   End
   Begin VB.Frame FrameBajaSocios 
      Height          =   4050
      Left            =   0
      TabIndex        =   220
      Top             =   0
      Width           =   7785
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   675
         Left            =   540
         TabIndex        =   883
         Top             =   2520
         Width           =   7065
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
            Index           =   181
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   884
            Text            =   "Text5"
            Top             =   210
            Width           =   4065
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
            Index           =   181
            Left            =   2100
            MaxLength       =   3
            TabIndex        =   227
            Top             =   210
            Width           =   630
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   181
            Left            =   1800
            MouseIcon       =   "frmListado.frx":4A1F
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar situaci�n"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Situaci�n campo"
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
            Index           =   266
            Left            =   120
            TabIndex        =   885
            Top             =   255
            Width           =   1635
         End
      End
      Begin VB.CheckBox chkBaja 
         Caption         =   "Dar de baja los campos asignados"
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
         Height          =   195
         Left            =   660
         TabIndex        =   226
         Top             =   2310
         Width           =   5415
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
         Index           =   47
         Left            =   2610
         MaxLength       =   10
         TabIndex        =   225
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelBajaSocio 
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
         Left            =   6330
         TabIndex        =   229
         Top             =   3285
         Width           =   1005
      End
      Begin VB.CommandButton cmdAcepBajaSocio 
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
         Left            =   5220
         TabIndex        =   228
         Top             =   3285
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
         Index           =   46
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   224
         Top             =   1260
         Width           =   630
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
         Index           =   46
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   223
         Text            =   "Text5"
         Top             =   1260
         Width           =   4065
      End
      Begin VB.CommandButton Command12 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":4B71
         Style           =   1  'Graphical
         TabIndex        =   222
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command11 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":4E7B
         Style           =   1  'Graphical
         TabIndex        =   221
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
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
         Index           =   66
         Left            =   660
         TabIndex        =   232
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Baja de Socio"
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
         TabIndex        =   231
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Situaci�n socio"
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
         Index           =   58
         Left            =   660
         TabIndex        =   230
         Top             =   1230
         Width           =   1500
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   46
         Left            =   2340
         MouseIcon       =   "frmListado.frx":5185
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar situaci�n"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   2310
         Picture         =   "frmListado.frx":52D7
         ToolTipText     =   "Buscar fecha"
         Top             =   1830
         Width           =   240
      End
   End
   Begin VB.Frame FrameDiferenciaKilos 
      Height          =   5670
      Left            =   0
      TabIndex        =   887
      Top             =   0
      Width           =   6615
      Begin MSComctlLib.ProgressBar pb11 
         Height          =   225
         Left            =   630
         TabIndex        =   915
         Top             =   4830
         Visible         =   0   'False
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame Frame11 
         Caption         =   "Clasificado por"
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
         Height          =   750
         Left            =   630
         TabIndex        =   912
         Top             =   3990
         Width           =   2820
         Begin VB.OptionButton Opcion1 
            Caption         =   "Variedad"
            Height          =   255
            Index           =   15
            Left            =   1500
            TabIndex        =   914
            Top             =   330
            Width           =   1200
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Socio"
            Height          =   345
            Index           =   14
            Left            =   300
            TabIndex        =   913
            Top             =   270
            Width           =   1290
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   189
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   901
         Text            =   "Text5"
         Top             =   2580
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   188
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   900
         Text            =   "Text5"
         Top             =   2220
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   189
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   893
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   188
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   892
         Top             =   2220
         Width           =   735
      End
      Begin VB.CommandButton Command64 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":5362
         Style           =   1  'Graphical
         TabIndex        =   899
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command63 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":566C
         Style           =   1  'Graphical
         TabIndex        =   898
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   187
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   896
         Text            =   "Text5"
         Top             =   1635
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   186
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   894
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   187
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   891
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   186
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   890
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepDifKilos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   889
         Top             =   5130
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   27
         Left            =   5250
         TabIndex        =   888
         Top             =   5130
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   185
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   897
         Top             =   3630
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   184
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   895
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         Height          =   195
         Index           =   267
         Left            =   630
         TabIndex        =   916
         Top             =   5100
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   128
         Left            =   1620
         MouseIcon       =   "frmListado.frx":5976
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   127
         Left            =   1620
         MouseIcon       =   "frmListado.frx":5AC8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   279
         Left            =   1005
         TabIndex        =   911
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   278
         Left            =   1005
         TabIndex        =   910
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
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
         Index           =   277
         Left            =   675
         TabIndex        =   909
         Top             =   2025
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   36
         Left            =   1650
         Picture         =   "frmListado.frx":5C1A
         ToolTipText     =   "Buscar fecha"
         Top             =   3630
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   35
         Left            =   1650
         Picture         =   "frmListado.frx":5CA5
         ToolTipText     =   "Buscar fecha"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   126
         Left            =   1620
         MouseIcon       =   "frmListado.frx":5D30
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   125
         Left            =   1620
         MouseIcon       =   "frmListado.frx":5E82
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   276
         Left            =   675
         TabIndex        =   908
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label36 
         Caption         =   "Informe de Diferencia de Kilos"
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
         Left            =   660
         TabIndex        =   907
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   275
         Left            =   960
         TabIndex        =   906
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   274
         Left            =   960
         TabIndex        =   905
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   273
         Left            =   1035
         TabIndex        =   904
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   272
         Left            =   1035
         TabIndex        =   903
         Top             =   3300
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   271
         Left            =   705
         TabIndex        =   902
         Top             =   3060
         Width           =   435
      End
   End
   Begin VB.Frame FrameCambioSocio 
      Height          =   4890
      Left            =   0
      TabIndex        =   519
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   106
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   525
         Top             =   2610
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   5100
         TabIndex        =   529
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdAcepCambsoc 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   5
         Left            =   4020
         TabIndex        =   527
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   111
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   524
         Top             =   2145
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   111
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   523
         Text            =   "Text5"
         Top             =   2145
         Width           =   3375
      End
      Begin VB.CommandButton Command34 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":5FD4
         Style           =   1  'Graphical
         TabIndex        =   522
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command31 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":62DE
         Style           =   1  'Graphical
         TabIndex        =   521
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   107
         Left            =   1905
         MaxLength       =   4
         TabIndex        =   526
         Top             =   3090
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   107
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   520
         Text            =   "Text5"
         Top             =   3090
         Width           =   3375
      End
      Begin VB.Label Label21 
         Caption         =   $"frmListado.frx":65E8
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   540
         TabIndex        =   533
         Top             =   960
         Width           =   5475
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Cambio"
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
         Index           =   158
         Left            =   540
         TabIndex        =   532
         Top             =   2610
         Width           =   1005
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   20
         Left            =   1590
         Picture         =   "frmListado.frx":6681
         ToolTipText     =   "Buscar fecha"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Label Label20 
         Caption         =   "Cambio de Socio "
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
         Left            =   570
         TabIndex        =   531
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Socio"
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
         Index           =   161
         Left            =   540
         TabIndex        =   530
         Top             =   2130
         Width           =   885
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   71
         Left            =   1590
         MouseIcon       =   "frmListado.frx":670C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2145
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Incidencia"
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
         Index           =   160
         Left            =   540
         TabIndex        =   528
         Top             =   3090
         Width           =   720
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   72
         Left            =   1590
         MouseIcon       =   "frmListado.frx":685E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar incidencia"
         Top             =   3090
         Width           =   240
      End
   End
   Begin VB.Frame FrameEntradasCampo 
      Height          =   6690
      Left            =   30
      TabIndex        =   92
      Top             =   -30
      Width           =   6615
      Begin VB.Frame FrameRecolectado 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   570
         TabIndex        =   300
         Top             =   4890
         Width           =   3300
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   14
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   768
            Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
            Top             =   1020
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   9
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   302
            Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
            Top             =   570
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   8
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   301
            Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
            Top             =   135
            Width           =   1575
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   2
            Left            =   2940
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   1050
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Socio"
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
            Index           =   229
            Left            =   90
            TabIndex        =   769
            Top             =   1065
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Entrada"
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
            Index           =   89
            Left            =   90
            TabIndex        =   373
            Top             =   615
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Recolectado"
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
            Index           =   81
            Left            =   90
            TabIndex        =   303
            Top             =   180
            Width           =   885
         End
      End
      Begin VB.Frame FrameTipoAlbaran 
         BorderStyle     =   0  'None
         Height          =   1545
         Left            =   570
         TabIndex        =   375
         Top             =   4890
         Width           =   3480
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   10
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   376
            Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
            Top             =   135
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo albaran"
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
            Index           =   92
            Left            =   90
            TabIndex        =   377
            Top             =   180
            Width           =   885
         End
      End
      Begin VB.Frame FrameTipo 
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   540
         TabIndex        =   141
         Top             =   4950
         Width           =   3480
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
            Top             =   585
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
            Top             =   135
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Entradas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   29
            Left            =   90
            TabIndex        =   143
            Top             =   585
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Informe"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   36
            Left            =   90
            TabIndex        =   142
            Top             =   180
            Width           =   1110
         End
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Omitir Gastos"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   374
         Top             =   5550
         Width           =   2205
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Detallar Notas"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   299
         Top             =   4490
         Width           =   1815
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Salta p�gina por socio"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   263
         Top             =   5190
         Width           =   2205
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Clasificado por Socio"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   262
         Top             =   4840
         Width           =   1815
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   137
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
         TabIndex        =   136
         Text            =   "Text5"
         Top             =   2220
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   96
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   95
         Top             =   2220
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":69B0
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
         Text            =   "Text5"
         Top             =   3210
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   98
         Top             =   3570
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   97
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
         TabIndex        =   121
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
         TabIndex        =   106
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   94
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   93
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   4110
         TabIndex        =   103
         Top             =   5955
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5190
         TabIndex        =   104
         Top             =   5955
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   100
         Top             =   4545
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   99
         Top             =   4140
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Imprimir Resumen"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   105
         Top             =   4140
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1620
         MouseIcon       =   "frmListado.frx":6FC4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1620
         MouseIcon       =   "frmListado.frx":7116
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
         TabIndex        =   140
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   1005
         TabIndex        =   139
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
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
         Index           =   11
         Left            =   675
         TabIndex        =   138
         Top             =   2025
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1620
         Picture         =   "frmListado.frx":7268
         ToolTipText     =   "Buscar fecha"
         Top             =   4140
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1620
         Picture         =   "frmListado.frx":72F3
         ToolTipText     =   "Buscar fecha"
         Top             =   4545
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1620
         MouseIcon       =   "frmListado.frx":737E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1620
         MouseIcon       =   "frmListado.frx":74D0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1620
         MouseIcon       =   "frmListado.frx":7622
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1620
         MouseIcon       =   "frmListado.frx":7774
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   27
         Left            =   675
         TabIndex        =   135
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   1005
         TabIndex        =   134
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   1005
         TabIndex        =   133
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
         TabIndex        =   132
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Left            =   675
         TabIndex        =   131
         Top             =   3015
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   960
         TabIndex        =   130
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   960
         TabIndex        =   129
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   1005
         TabIndex        =   128
         Top             =   4545
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   1005
         TabIndex        =   127
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   19
         Left            =   675
         TabIndex        =   126
         Top             =   3960
         Width           =   435
      End
   End
   Begin VB.Frame FrameCampos 
      Height          =   9795
      Left            =   0
      TabIndex        =   48
      Top             =   -60
      Width           =   6615
      Begin VB.CheckBox Check27 
         Caption         =   "Ordenar por partida"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3150
         TabIndex        =   886
         Top             =   8700
         Width           =   2115
      End
      Begin VB.CheckBox Check26 
         Caption         =   "Salta p�gina por Socio"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3150
         TabIndex        =   882
         Top             =   8460
         Width           =   2115
      End
      Begin VB.CheckBox Check23 
         Caption         =   "Informe Conselleria"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3150
         TabIndex        =   866
         Top             =   8370
         Width           =   2115
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   134
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   644
         Text            =   "Text5"
         Top             =   5880
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   134
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   69
         Top             =   5880
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   133
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   641
         Text            =   "Text5"
         Top             =   5490
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   133
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   68
         Top             =   5490
         Width           =   735
      End
      Begin VB.CheckBox Check16 
         Caption         =   "Hect�reas"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   572
         Top             =   7500
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Imprimir Datos Recintos"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3150
         TabIndex        =   518
         Top             =   8070
         Width           =   2115
      End
      Begin MSComctlLib.ProgressBar pb5 
         Height          =   225
         Left            =   480
         TabIndex        =   471
         Top             =   9000
         Visible         =   0   'False
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   95
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   67
         Top             =   5025
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   94
         Left            =   1935
         MaxLength       =   4
         TabIndex        =   66
         Top             =   4620
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   95
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   467
         Text            =   "Text5"
         Top             =   5025
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   94
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   466
         Text            =   "Text5"
         Top             =   4620
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   92
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   462
         Text            =   "Text5"
         Top             =   3750
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   93
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   461
         Text            =   "Text5"
         Top             =   4155
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   92
         Left            =   1950
         MaxLength       =   4
         TabIndex        =   64
         Top             =   3750
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   93
         Left            =   1950
         MaxLength       =   4
         TabIndex        =   65
         Top             =   4140
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   11
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   7050
         Width           =   1440
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Imprimir Cabecera Cooperativa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3150
         TabIndex        =   233
         Top             =   7770
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   6690
         Width           =   1440
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3150
         TabIndex        =   90
         Top             =   7470
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   6330
         Width           =   1440
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   63
         Top             =   3285
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1935
         MaxLength       =   2
         TabIndex        =   62
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "Text5"
         Top             =   3285
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "Text5"
         Top             =   2880
         Width           =   3375
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5160
         TabIndex        =   75
         Top             =   9240
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   73
         Top             =   9240
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   59
         Top             =   1560
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   58
         Top             =   1155
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text5"
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Text5"
         Top             =   1155
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   61
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   60
         Top             =   1995
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text5"
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "Text5"
         Top             =   1995
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":78C6
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command1 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":7BD0
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Frame Frame3 
         Caption         =   "Clasificado por"
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
         Height          =   2250
         Left            =   480
         TabIndex        =   49
         Top             =   6240
         Width           =   2460
         Begin VB.OptionButton Opcion1 
            Caption         =   "Variedad/Zona"
            Height          =   255
            Index           =   7
            Left            =   300
            TabIndex        =   700
            Top             =   1920
            Width           =   2040
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Variedad/Resp./Partida"
            Height          =   255
            Index           =   4
            Left            =   300
            TabIndex        =   460
            Top             =   1590
            Width           =   2040
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Zona"
            Height          =   255
            Index           =   3
            Left            =   300
            TabIndex        =   83
            Top             =   1260
            Width           =   1470
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Termino Municipal"
            Height          =   255
            Index           =   2
            Left            =   300
            TabIndex        =   82
            Top             =   945
            Width           =   1605
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Socio"
            Height          =   345
            Index           =   0
            Left            =   300
            TabIndex        =   51
            Top             =   225
            Width           =   1290
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Clase/Variedad"
            Height          =   255
            Index           =   1
            Left            =   300
            TabIndex        =   50
            Top             =   585
            Width           =   1470
         End
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   3
         Left            =   5340
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   8340
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   83
         Left            =   1560
         MouseIcon       =   "frmListado.frx":7EDA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar zona"
         Top             =   5880
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   191
         Left            =   960
         TabIndex        =   645
         Top             =   5910
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   82
         Left            =   1560
         MouseIcon       =   "frmListado.frx":802C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar zona"
         Top             =   5520
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   190
         Left            =   960
         TabIndex        =   643
         Top             =   5535
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
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
         Index           =   188
         Left            =   630
         TabIndex        =   642
         Top             =   5340
         Width           =   360
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   5340
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   8040
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         Height          =   195
         Index           =   136
         Left            =   540
         TabIndex        =   472
         Top             =   9300
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Partida"
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
         Index           =   135
         Left            =   630
         TabIndex        =   470
         Top             =   4470
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   134
         Left            =   960
         TabIndex        =   469
         Top             =   4665
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   133
         Left            =   960
         TabIndex        =   468
         Top             =   5055
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   58
         Left            =   1560
         MouseIcon       =   "frmListado.frx":817E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar partida"
         Top             =   5025
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   57
         Left            =   1575
         MouseIcon       =   "frmListado.frx":82D0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar partida"
         Top             =   4650
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   59
         Left            =   1575
         MouseIcon       =   "frmListado.frx":8422
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   60
         Left            =   1560
         MouseIcon       =   "frmListado.frx":8574
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   4155
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   132
         Left            =   960
         TabIndex        =   465
         Top             =   4185
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   131
         Left            =   960
         TabIndex        =   464
         Top             =   3795
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
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
         Index           =   118
         Left            =   630
         TabIndex        =   463
         Top             =   3600
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Campo"
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
         Index           =   109
         Left            =   3150
         TabIndex        =   399
         Top             =   7080
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Producci�n"
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
         Index           =   10
         Left            =   3150
         TabIndex        =   91
         Top             =   6720
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Hect�reas"
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
         Index           =   9
         Left            =   3150
         TabIndex        =   89
         Top             =   6360
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Situaci�n"
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
         Index           =   8
         Left            =   630
         TabIndex        =   88
         Top             =   2730
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   87
         Top             =   2925
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   86
         Top             =   3315
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1560
         MouseIcon       =   "frmListado.frx":86C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar situaci�n"
         Top             =   3285
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1575
         MouseIcon       =   "frmListado.frx":8818
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar situaci�n"
         Top             =   2910
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   81
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   80
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label2 
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
         Index           =   3
         Left            =   630
         TabIndex        =   79
         Top             =   1845
         Width           =   390
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Campos/Huertos"
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
         TabIndex        =   78
         Top             =   405
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   77
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   76
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   630
         TabIndex        =   74
         Top             =   990
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1560
         MouseIcon       =   "frmListado.frx":896A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1560
         MouseIcon       =   "frmListado.frx":8ABC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1560
         MouseIcon       =   "frmListado.frx":8C0E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1575
         MouseIcon       =   "frmListado.frx":8D60
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2025
         Width           =   240
      End
   End
   Begin VB.Frame FrameTraspasoCalibrador 
      Height          =   4665
      Left            =   0
      TabIndex        =   201
      Top             =   -30
      Width           =   6555
      Begin VB.Frame FrameNota 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Enabled         =   0   'False
         Height          =   795
         Left            =   3630
         TabIndex        =   878
         Top             =   1440
         Visible         =   0   'False
         Width           =   2685
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   179
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   208
            Tag             =   "N� Albar�n|N|S|||rhisfruta|numalbar|0000000|S|"
            Top             =   510
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   170
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   207
            Tag             =   "N� Albar�n|N|S|||rhisfruta|numalbar|0000000|S|"
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta Nota "
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
            Index           =   264
            Left            =   270
            TabIndex        =   880
            Top             =   510
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Desde Nota "
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
            Index           =   263
            Left            =   270
            TabIndex        =   879
            Top             =   180
            Width           =   885
         End
      End
      Begin VB.Frame FrameFecha 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   795
         Left            =   3570
         TabIndex        =   340
         Top             =   1380
         Width           =   2685
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   63
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   206
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            Index           =   90
            Left            =   270
            TabIndex        =   341
            Top             =   270
            Width           =   435
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   9
            Left            =   900
            Picture         =   "frmListado.frx":8EB2
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   285
         Left            =   240
         TabIndex        =   212
         Top             =   2370
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   6
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   205
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   1620
         Width           =   2295
      End
      Begin VB.CommandButton cmdAcepTras 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   209
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelTras 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   210
         Top             =   3780
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   570
         Top             =   3390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Calibrador"
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
         Index           =   51
         Left            =   390
         TabIndex        =   211
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Proceso que realiza el Traspaso desde el Calibrador seleccionado de la clasificaci�n de entradas."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   525
         Index           =   37
         Left            =   300
         TabIndex        =   204
         Top             =   630
         Width           =   5820
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProgres 
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   203
         Top             =   3480
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   202
         Top             =   3120
         Width           =   6195
      End
   End
   Begin VB.Frame FrameTrazabilidad 
      Height          =   4665
      Left            =   30
      TabIndex        =   213
      Top             =   -60
      Width           =   6555
      Begin VB.CommandButton CmdCancelTraza 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   216
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepTraza 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   215
         Top             =   3780
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   285
         Left            =   240
         TabIndex        =   214
         Top             =   2130
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   570
         Top             =   3390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.Label lblProgres 
         Caption         =   "aa"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   219
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Caption         =   "aa"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   218
         Top             =   3480
         Width           =   6195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Proceso que realiza el Traspaso de TRAZABILIDAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   525
         Index           =   53
         Left            =   240
         TabIndex        =   217
         Top             =   870
         Width           =   5820
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame FrameSociosSeccion 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7020
      Begin VB.CheckBox Check24 
         Caption         =   "S�lo de baja"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2820
         TabIndex        =   877
         Top             =   4290
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   15
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   831
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   4500
         Width           =   2100
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Socios O.P. Control democr�tico"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   540
         TabIndex        =   830
         Top             =   4710
         Width           =   2655
      End
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Agrupado por"
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
         Height          =   975
         Left            =   495
         TabIndex        =   21
         Top             =   3195
         Width           =   2190
         Begin VB.OptionButton Opcion 
            Caption         =   "Socio"
            Height          =   255
            Index           =   1
            Left            =   495
            TabIndex        =   23
            Top             =   585
            Width           =   975
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Secci�n"
            Height          =   345
            Index           =   0
            Left            =   495
            TabIndex        =   22
            Top             =   225
            Width           =   1290
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ordenado por"
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
         Height          =   975
         Left            =   3990
         TabIndex        =   396
         Top             =   3180
         Width           =   2190
         Begin VB.OptionButton Opcion 
            Caption         =   "C�digo"
            Height          =   345
            Index           =   5
            Left            =   495
            TabIndex        =   398
            Top             =   225
            Width           =   1290
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Alfab�tico"
            Height          =   255
            Index           =   4
            Left            =   495
            TabIndex        =   397
            Top             =   585
            Width           =   1305
         End
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Imprimir Socios de baja"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   540
         TabIndex        =   298
         Top             =   4290
         Width           =   2355
      End
      Begin VB.CommandButton Command6 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":8F3D
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command5 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":9247
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   2760
         Width           =   3315
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   2400
         Width           =   3315
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   8
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   7
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text5"
         Top             =   1635
         Width           =   3345
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   1275
         Width           =   3345
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   4170
         TabIndex        =   2
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5220
         TabIndex        =   1
         Top             =   5040
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   3
         Left            =   6120
         TabIndex        =   13
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Socio"
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
         Index           =   233
         Left            =   4110
         TabIndex        =   832
         Top             =   4230
         Width           =   945
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1575
         MouseIcon       =   "frmListado.frx":9551
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1560
         MouseIcon       =   "frmListado.frx":96A3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1560
         MouseIcon       =   "frmListado.frx":97F5
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar secci�n"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1560
         MouseIcon       =   "frmListado.frx":9947
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar secci�n"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   17
         Left            =   600
         TabIndex        =   20
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   16
         Left            =   960
         TabIndex        =   19
         Top             =   2790
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   18
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label5 
         Caption         =   "Informe de Socios por Secci�n"
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
         Left            =   540
         TabIndex        =   17
         Top             =   405
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Secci�n"
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
         Index           =   14
         Left            =   600
         TabIndex        =   16
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   15
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   14
         Top             =   1320
         Width           =   465
      End
   End
   Begin VB.Frame FrameTraspasoAlbRetirada 
      Height          =   4665
      Left            =   0
      TabIndex        =   855
      Top             =   90
      Width           =   6555
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   169
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   864
         ToolTipText     =   " "
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   169
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   863
         Text            =   "Text5"
         Top             =   1620
         Width           =   3375
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   25
         Left            =   4905
         TabIndex        =   858
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepTrasRetirada 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   857
         Top             =   3780
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pb9 
         Height          =   285
         Left            =   240
         TabIndex        =   856
         Top             =   2370
         Visible         =   0   'False
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   122
         Left            =   1500
         MouseIcon       =   "frmListado.frx":9A99
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cooperativa"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   5
         Left            =   180
         TabIndex        =   862
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   861
         Top             =   2700
         Width           =   6195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Proceso que realiza el Traspaso de albaranes de retirada de las cooperativas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   525
         Index           =   261
         Left            =   300
         TabIndex        =   860
         Top             =   630
         Width           =   5820
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cooperativa"
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
         Index           =   254
         Left            =   390
         TabIndex        =   859
         Top             =   1680
         Width           =   885
      End
   End
   Begin VB.Frame FrameTraspDatosATrazabilidad 
      Height          =   4320
      Left            =   0
      TabIndex        =   833
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   178
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   838
         Top             =   2610
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   177
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   837
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   177
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   846
         Text            =   "Text5"
         Top             =   2265
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   178
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   845
         Text            =   "Text5"
         Top             =   2625
         Width           =   3375
      End
      Begin VB.CommandButton Command60 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":9BEB
         Style           =   1  'Graphical
         TabIndex        =   844
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command59 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":9EF5
         Style           =   1  'Graphical
         TabIndex        =   843
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   171
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   841
         Text            =   "Text5"
         Top             =   1305
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   172
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   839
         Text            =   "Text5"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   172
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   836
         Top             =   1680
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   171
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   835
         ToolTipText     =   " "
         Top             =   1320
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepDatosTraza 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   840
         Top             =   3615
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   24
         Left            =   5160
         TabIndex        =   842
         Top             =   3615
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar ProgressBar5 
         Height          =   225
         Left            =   540
         TabIndex        =   834
         Top             =   4080
         Visible         =   0   'False
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   253
         Left            =   570
         TabIndex        =   854
         Top             =   2070
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   252
         Left            =   900
         TabIndex        =   853
         Top             =   2310
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   251
         Left            =   900
         TabIndex        =   852
         Top             =   2700
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   117
         Left            =   1530
         MouseIcon       =   "frmListado.frx":A1FF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   116
         Left            =   1515
         MouseIcon       =   "frmListado.frx":A351
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   115
         Left            =   1530
         MouseIcon       =   "frmListado.frx":A4A3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   114
         Left            =   1530
         MouseIcon       =   "frmListado.frx":A5F5
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   250
         Left            =   585
         TabIndex        =   851
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label34 
         Caption         =   "Traspaso datos a Trazabilidad"
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
         TabIndex        =   850
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   249
         Left            =   870
         TabIndex        =   849
         Top             =   1740
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   248
         Left            =   870
         TabIndex        =   848
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         Height          =   195
         Index           =   237
         Left            =   540
         TabIndex        =   847
         Top             =   3360
         Visible         =   0   'False
         Width           =   3345
      End
   End
   Begin VB.Frame FrameTraspasoROPAS 
      Height          =   4890
      Left            =   0
      TabIndex        =   304
      Top             =   30
      Width           =   6615
      Begin MSComctlLib.ProgressBar pb7 
         Height          =   225
         Left            =   540
         TabIndex        =   639
         Top             =   4080
         Visible         =   0   'False
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   132
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   316
         Top             =   3660
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   62
         Left            =   1830
         MaxLength       =   4
         TabIndex        =   315
         Top             =   3105
         Width           =   735
      End
      Begin VB.CommandButton CmdCancelROPAS 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5160
         TabIndex        =   318
         Top             =   4395
         Width           =   975
      End
      Begin VB.CommandButton cmdAcepROPAS 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   317
         Top             =   4395
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   311
         ToolTipText     =   " "
         Top             =   1335
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   59
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   312
         Top             =   1695
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   58
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   310
         Text            =   "Text5"
         Top             =   1335
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   59
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   309
         Text            =   "Text5"
         Top             =   1695
         Width           =   3375
      End
      Begin VB.CommandButton Command16 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":A747
         Style           =   1  'Graphical
         TabIndex        =   308
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command19 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":AA51
         Style           =   1  'Graphical
         TabIndex        =   307
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   61
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   306
         Text            =   "Text5"
         Top             =   2625
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   60
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   305
         Text            =   "Text5"
         Top             =   2265
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   1830
         MaxLength       =   3
         TabIndex        =   314
         Top             =   2625
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   1830
         MaxLength       =   3
         TabIndex        =   313
         Top             =   2265
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         Height          =   195
         Index           =   187
         Left            =   570
         TabIndex        =   640
         Top             =   4320
         Visible         =   0   'False
         Width           =   3345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Envio"
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
         Index           =   186
         Left            =   570
         TabIndex        =   638
         Top             =   3660
         Width           =   870
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   24
         Left            =   1515
         Picture         =   "frmListado.frx":AD5B
         ToolTipText     =   "Buscar fecha"
         Top             =   3660
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   83
         Left            =   870
         TabIndex        =   326
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   84
         Left            =   870
         TabIndex        =   325
         Top             =   1740
         Width           =   420
      End
      Begin VB.Label Label12 
         Caption         =   "Traspaso ROPAS"
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
         TabIndex        =   324
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
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
         Index           =   85
         Left            =   585
         TabIndex        =   323
         Top             =   1140
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   45
         Left            =   1530
         MouseIcon       =   "frmListado.frx":ADE6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   42
         Left            =   1530
         MouseIcon       =   "frmListado.frx":AF38
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   44
         Left            =   1515
         MouseIcon       =   "frmListado.frx":B08A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   43
         Left            =   1530
         MouseIcon       =   "frmListado.frx":B1DC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1740
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   52
         Left            =   900
         TabIndex        =   322
         Top             =   2700
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   86
         Left            =   900
         TabIndex        =   321
         Top             =   2310
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
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
         Index           =   87
         Left            =   570
         TabIndex        =   320
         Top             =   2070
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ejercicio"
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
         Index           =   79
         Left            =   570
         TabIndex        =   319
         Top             =   3150
         Width           =   585
      End
   End
   Begin VB.Frame FrameGastosporConcepto 
      DragMode        =   1  'Automatic
      Height          =   7680
      Left            =   0
      TabIndex        =   473
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check11 
         Caption         =   "Saltar p�gina"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4530
         TabIndex        =   490
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Frame Frame6 
         Caption         =   "Clasificado por"
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
         Height          =   720
         Left            =   540
         TabIndex        =   517
         Top             =   5880
         Width           =   3690
         Begin VB.OptionButton Opcion1 
            Caption         =   "Variedad"
            Height          =   255
            Index           =   9
            Left            =   2160
            TabIndex        =   489
            Top             =   270
            Width           =   1320
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Socio"
            Height          =   345
            Index           =   8
            Left            =   480
            TabIndex        =   488
            Top             =   225
            Width           =   1290
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   97
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   516
         Text            =   "Text5"
         Top             =   5430
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   96
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   515
         Text            =   "Text5"
         Top             =   5070
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   105
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   496
         Text            =   "Text5"
         Top             =   2580
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   104
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   495
         Text            =   "Text5"
         Top             =   2220
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   105
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   477
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   104
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   476
         Top             =   2220
         Width           =   735
      End
      Begin VB.CommandButton Command33 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":B32E
         Style           =   1  'Graphical
         TabIndex        =   493
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command32 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":B638
         Style           =   1  'Graphical
         TabIndex        =   491
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   103
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   486
         Text            =   "Text5"
         Top             =   3570
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   102
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   484
         Text            =   "Text5"
         Top             =   3210
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   103
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   479
         Top             =   3570
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   102
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   478
         Top             =   3210
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   101
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   482
         Text            =   "Text5"
         Top             =   1635
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   100
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   480
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   101
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   475
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   100
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   474
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepGtosConcep 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   492
         Top             =   7020
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5220
         TabIndex        =   494
         Top             =   7020
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   99
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   483
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   98
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   481
         Top             =   4140
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   97
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   487
         Top             =   5445
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   96
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   485
         Top             =   5070
         Width           =   735
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   330
         TabIndex        =   497
         Top             =   6720
         Visible         =   0   'False
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   70
         Left            =   1620
         MouseIcon       =   "frmListado.frx":B942
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   5460
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   65
         Left            =   1620
         MouseIcon       =   "frmListado.frx":BA94
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   5100
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   64
         Left            =   1620
         MouseIcon       =   "frmListado.frx":BBE6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   63
         Left            =   1620
         MouseIcon       =   "frmListado.frx":BD38
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   152
         Left            =   1005
         TabIndex        =   514
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   151
         Left            =   1005
         TabIndex        =   513
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
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
         Index           =   150
         Left            =   675
         TabIndex        =   512
         Top             =   2025
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   17
         Left            =   1620
         Picture         =   "frmListado.frx":BE8A
         ToolTipText     =   "Buscar fecha"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   16
         Left            =   1620
         Picture         =   "frmListado.frx":BF15
         ToolTipText     =   "Buscar fecha"
         Top             =   4140
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   62
         Left            =   1620
         MouseIcon       =   "frmListado.frx":BFA0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   61
         Left            =   1620
         MouseIcon       =   "frmListado.frx":C0F2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   50
         Left            =   1620
         MouseIcon       =   "frmListado.frx":C244
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   47
         Left            =   1620
         MouseIcon       =   "frmListado.frx":C396
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   149
         Left            =   675
         TabIndex        =   511
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   148
         Left            =   1005
         TabIndex        =   510
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   147
         Left            =   1005
         TabIndex        =   509
         Top             =   3255
         Width           =   465
      End
      Begin VB.Label Label19 
         Caption         =   "Informe de Gastos por Concepto"
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
         TabIndex        =   508
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   146
         Left            =   675
         TabIndex        =   507
         Top             =   3015
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   145
         Left            =   960
         TabIndex        =   506
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   144
         Left            =   960
         TabIndex        =   505
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   143
         Left            =   1005
         TabIndex        =   504
         Top             =   4545
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   142
         Left            =   1005
         TabIndex        =   503
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   141
         Left            =   675
         TabIndex        =   502
         Top             =   3960
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   140
         Left            =   990
         TabIndex        =   501
         Top             =   5445
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   139
         Left            =   990
         TabIndex        =   500
         Top             =   5100
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
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
         Index           =   138
         Left            =   660
         TabIndex        =   499
         Top             =   4860
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         Height          =   195
         Index           =   137
         Left            =   360
         TabIndex        =   498
         Top             =   7020
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Frame FrameOrdenRecoleccion 
      Height          =   4575
      Left            =   0
      TabIndex        =   659
      Top             =   0
      Width           =   6615
      Begin VB.Frame FrameNroOrden 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   3630
         TabIndex        =   698
         Top             =   2910
         Visible         =   0   'False
         Width           =   2715
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   141
            Left            =   1470
            MaxLength       =   10
            TabIndex        =   667
            Top             =   210
            Width           =   1005
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   89
            Left            =   1170
            MouseIcon       =   "frmListado.frx":C4E8
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar Nro.Orden"
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nro.Orden"
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
            Index           =   204
            Left            =   150
            TabIndex        =   699
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.CheckBox Check19 
         Caption         =   "Reimpresi�n"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   666
         Top             =   3150
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   138
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   665
         Top             =   2610
         Width           =   1005
      End
      Begin VB.CommandButton Command44 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":C63A
         Style           =   1  'Graphical
         TabIndex        =   673
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command39 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":C944
         Style           =   1  'Graphical
         TabIndex        =   672
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   149
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   671
         Text            =   "Text5"
         Top             =   1635
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   149
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   663
         Top             =   1635
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   147
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   670
         Text            =   "Text5"
         Top             =   1155
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   147
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   662
         Top             =   1155
         Width           =   750
      End
      Begin VB.CommandButton cmdAcepOrdenRec 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3990
         TabIndex        =   668
         Top             =   3990
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   5070
         TabIndex        =   669
         Top             =   3990
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   142
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   664
         Top             =   2130
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   142
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   661
         Text            =   "Text5"
         Top             =   2130
         Width           =   3375
      End
      Begin MSComctlLib.ProgressBar ProgressBar4 
         Height          =   225
         Left            =   570
         TabIndex        =   660
         Top             =   3720
         Visible         =   0   'False
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgAyuda 
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   2130
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   3120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   28
         Left            =   1620
         Picture         =   "frmListado.frx":CC4E
         ToolTipText     =   "Buscar fecha"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   195
         Left            =   630
         TabIndex        =   679
         Top             =   2640
         Width           =   435
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   85
         Left            =   1620
         MouseIcon       =   "frmListado.frx":CCD9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   84
         Left            =   1620
         MouseIcon       =   "frmListado.frx":CE2B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Partida"
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
         Index           =   219
         Left            =   630
         TabIndex        =   678
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label Label27 
         Caption         =   "Ordenes de Recolecci�n"
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
         TabIndex        =   677
         Top             =   405
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   216
         Left            =   630
         TabIndex        =   676
         Top             =   1635
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
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
         Index           =   207
         Left            =   630
         TabIndex        =   675
         Top             =   1170
         Width           =   915
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   86
         Left            =   1620
         MouseIcon       =   "frmListado.frx":CF7D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar partida"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         Height          =   195
         Index           =   201
         Left            =   570
         TabIndex        =   674
         Top             =   3990
         Visible         =   0   'False
         Width           =   3525
      End
   End
   Begin VB.Frame FrameRevisionCampos 
      Height          =   5340
      Left            =   30
      TabIndex        =   770
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   166
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   782
         Top             =   3900
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   165
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   781
         Top             =   3510
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   22
         Left            =   5190
         TabIndex        =   786
         Top             =   4545
         Width           =   975
      End
      Begin VB.CommandButton cmdAcepRevisionCampos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   784
         Top             =   4545
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   164
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   778
         Top             =   1620
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   163
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   777
         Top             =   1260
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   163
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   776
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   164
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   775
         Text            =   "Text5"
         Top             =   1635
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   162
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   780
         Top             =   2790
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   161
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   779
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   161
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   774
         Text            =   "Text5"
         Top             =   2430
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   162
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   773
         Text            =   "Text5"
         Top             =   2790
         Width           =   3375
      End
      Begin VB.CommandButton Command54 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":D0CF
         Style           =   1  'Graphical
         TabIndex        =   772
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command53 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":D3D9
         Style           =   1  'Graphical
         TabIndex        =   771
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   247
         Left            =   675
         TabIndex        =   794
         Top             =   3270
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   246
         Left            =   1005
         TabIndex        =   793
         Top             =   3510
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   245
         Left            =   1005
         TabIndex        =   792
         Top             =   3855
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   244
         Left            =   960
         TabIndex        =   791
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   243
         Left            =   960
         TabIndex        =   790
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   242
         Left            =   675
         TabIndex        =   789
         Top             =   2070
         Width           =   630
      End
      Begin VB.Label Label32 
         Caption         =   "Registro Diario de Visitas a Parcelas"
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
         TabIndex        =   788
         Top             =   420
         Width           =   5475
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   241
         Left            =   975
         TabIndex        =   787
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   240
         Left            =   975
         TabIndex        =   785
         Top             =   2805
         Width           =   420
      End
      Begin VB.Label Label2 
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
         Index           =   239
         Left            =   675
         TabIndex        =   783
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   111
         Left            =   1650
         MouseIcon       =   "frmListado.frx":D6E3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   110
         Left            =   1650
         MouseIcon       =   "frmListado.frx":D835
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   109
         Left            =   1650
         MouseIcon       =   "frmListado.frx":D987
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   108
         Left            =   1650
         MouseIcon       =   "frmListado.frx":DAD9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   34
         Left            =   1650
         Picture         =   "frmListado.frx":DC2B
         ToolTipText     =   "Buscar fecha"
         Top             =   3915
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   33
         Left            =   1650
         Picture         =   "frmListado.frx":DCB6
         ToolTipText     =   "Buscar fecha"
         Top             =   3510
         Width           =   240
      End
   End
   Begin VB.Frame FramePrecios 
      Height          =   4455
      Left            =   30
      TabIndex        =   748
      Top             =   30
      Width           =   6690
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   158
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   752
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   157
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   751
         Top             =   2220
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   13
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   753
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   3180
         Width           =   1575
      End
      Begin VB.CommandButton Command52 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":DD41
         Style           =   1  'Graphical
         TabIndex        =   759
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command51 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":E04B
         Style           =   1  'Graphical
         TabIndex        =   758
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   156
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   756
         Text            =   "Text5"
         Top             =   1635
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   155
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   754
         Text            =   "Text5"
         Top             =   1275
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   156
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   750
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   155
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   749
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton cmdAcepPrecios 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   5
         Left            =   4170
         TabIndex        =   755
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   21
         Left            =   5250
         TabIndex        =   757
         Top             =   3735
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
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
         Index           =   228
         Left            =   570
         TabIndex        =   767
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   227
         Left            =   900
         TabIndex        =   766
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   226
         Left            =   900
         TabIndex        =   765
         Top             =   2625
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   32
         Left            =   1545
         Picture         =   "frmListado.frx":E355
         ToolTipText     =   "Buscar fecha"
         Top             =   2625
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   31
         Left            =   1545
         Picture         =   "frmListado.frx":E3E0
         ToolTipText     =   "Buscar fecha"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo precio"
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
         Index           =   225
         Left            =   570
         TabIndex        =   764
         Top             =   3225
         Width           =   780
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   103
         Left            =   1545
         MouseIcon       =   "frmListado.frx":E46B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   102
         Left            =   1560
         MouseIcon       =   "frmListado.frx":E5BD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label Label31 
         Caption         =   "Informe de Precios"
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
         Left            =   540
         TabIndex        =   763
         Top             =   405
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   224
         Left            =   600
         TabIndex        =   762
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   223
         Left            =   900
         TabIndex        =   761
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   222
         Left            =   900
         TabIndex        =   760
         Top             =   1320
         Width           =   465
      End
   End
   Begin VB.Frame FrameInformeSocios 
      Height          =   4995
      Left            =   0
      TabIndex        =   701
      Top             =   0
      Width           =   6630
      Begin VB.Frame Frame8 
         Caption         =   "Tipo"
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
         Height          =   1095
         Left            =   480
         TabIndex        =   718
         Top             =   2220
         Width           =   2430
         Begin VB.OptionButton Opcion 
            Caption         =   "Tel�fonos"
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   719
            Top             =   300
            Width           =   1305
         End
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   19
         Left            =   4740
         TabIndex        =   709
         Top             =   4050
         Width           =   975
      End
      Begin VB.CommandButton cmdAcepInfSocios 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   708
         Top             =   4050
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   145
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   703
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   146
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   704
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   145
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   713
         Text            =   "Text5"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   146
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   712
         Text            =   "Text5"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CommandButton Command48 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":E70F
         Style           =   1  'Graphical
         TabIndex        =   711
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command47 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":EA19
         Style           =   1  'Graphical
         TabIndex        =   710
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Imprimir Socios de baja"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   630
         TabIndex        =   707
         Top             =   3720
         Width           =   2355
      End
      Begin VB.Frame Frame10 
         Caption         =   "Ordenado por"
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
         Height          =   1095
         Left            =   3330
         TabIndex        =   702
         Top             =   2220
         Width           =   2400
         Begin VB.OptionButton Opcion 
            Caption         =   "Alfab�tico"
            Height          =   255
            Index           =   9
            Left            =   480
            TabIndex        =   706
            Top             =   630
            Width           =   1305
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "C�digo"
            Height          =   255
            Index           =   8
            Left            =   480
            TabIndex        =   705
            Top             =   330
            Width           =   1305
         End
      End
      Begin VB.Label Label29 
         Caption         =   "Informe Datos Socios"
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
         Left            =   540
         TabIndex        =   717
         Top             =   405
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   210
         Left            =   900
         TabIndex        =   716
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   209
         Left            =   900
         TabIndex        =   715
         Top             =   1830
         Width           =   420
      End
      Begin VB.Label Label2 
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
         Index           =   208
         Left            =   540
         TabIndex        =   714
         Top             =   1200
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   96
         Left            =   1500
         MouseIcon       =   "frmListado.frx":ED23
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   97
         Left            =   1515
         MouseIcon       =   "frmListado.frx":EE75
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1830
         Width           =   240
      End
   End
   Begin VB.Frame FrameKilosRecolect 
      Height          =   6840
      Left            =   30
      TabIndex        =   264
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   57
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   281
         Text            =   "Text5"
         Top             =   2580
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   56
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   279
         Text            =   "Text5"
         Top             =   2220
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   57
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   270
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   56
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   269
         Top             =   2220
         Width           =   735
      End
      Begin VB.CommandButton Command18 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":EFC7
         Style           =   1  'Graphical
         TabIndex        =   277
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command17 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":F2D1
         Style           =   1  'Graphical
         TabIndex        =   275
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   55
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   273
         Text            =   "Text5"
         Top             =   1635
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   54
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   271
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   55
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   268
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   54
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   267
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepKilosSoc 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   282
         Top             =   6255
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelKilosSoc 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5160
         TabIndex        =   284
         Top             =   6255
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   53
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   278
         Top             =   4470
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   52
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   276
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   51
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   266
         Text            =   "Text5"
         Top             =   3570
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   50
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   265
         Text            =   "Text5"
         Top             =   3210
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   51
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   274
         Top             =   3570
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   50
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   272
         Top             =   3210
         Width           =   735
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Incluir pendiente de clasificar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   690
         TabIndex        =   280
         Top             =   4920
         Width           =   2565
      End
      Begin MSComctlLib.ProgressBar Pb3 
         Height          =   255
         Left            =   450
         TabIndex        =   297
         Top             =   5880
         Visible         =   0   'False
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1665
         Index           =   0
         Left            =   3660
         TabIndex        =   535
         Top             =   4170
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2937
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
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
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   5580
         Picture         =   "frmListado.frx":F5DB
         ToolTipText     =   "Desmarcar todos"
         Top             =   3900
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   5850
         Picture         =   "frmListado.frx":FFDD
         ToolTipText     =   "Marcar todos"
         Top             =   3900
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Entrada"
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
         Index           =   153
         Left            =   3690
         TabIndex        =   534
         Top             =   3930
         Width           =   1140
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   41
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1682F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   32
         Left            =   1620
         MouseIcon       =   "frmListado.frx":16981
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   78
         Left            =   1005
         TabIndex        =   296
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   77
         Left            =   1005
         TabIndex        =   295
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
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
         Index           =   76
         Left            =   675
         TabIndex        =   294
         Top             =   2025
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1620
         Picture         =   "frmListado.frx":16AD3
         ToolTipText     =   "Buscar fecha"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1620
         Picture         =   "frmListado.frx":16B5E
         ToolTipText     =   "Buscar fecha"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   31
         Left            =   1620
         MouseIcon       =   "frmListado.frx":16BE9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1620
         MouseIcon       =   "frmListado.frx":16D3B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   75
         Left            =   675
         TabIndex        =   293
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Kilos Recolectados Socio/Cooperativa"
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
         Left            =   660
         TabIndex        =   292
         Top             =   420
         Width           =   5595
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   74
         Left            =   960
         TabIndex        =   291
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   73
         Left            =   960
         TabIndex        =   290
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   72
         Left            =   1005
         TabIndex        =   289
         Top             =   4455
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   71
         Left            =   1005
         TabIndex        =   288
         Top             =   4110
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   70
         Left            =   675
         TabIndex        =   287
         Top             =   3870
         Width           =   435
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1620
         MouseIcon       =   "frmListado.frx":16E8D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   40
         Left            =   1620
         MouseIcon       =   "frmListado.frx":16FDF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   69
         Left            =   1005
         TabIndex        =   286
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   62
         Left            =   1005
         TabIndex        =   285
         Top             =   3255
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
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
         Index           =   61
         Left            =   675
         TabIndex        =   283
         Top             =   3015
         Width           =   645
      End
   End
   Begin VB.Frame FrameInfATRIA 
      DragMode        =   1  'Automatic
      Height          =   5400
      Left            =   0
      TabIndex        =   720
      Top             =   0
      Width           =   6615
      Begin MSComctlLib.ProgressBar Pb8 
         Height          =   225
         Left            =   630
         TabIndex        =   747
         Top             =   4380
         Visible         =   0   'False
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   151
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   743
         Text            =   "Text5"
         Top             =   3570
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   152
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   742
         Text            =   "Text5"
         Top             =   3960
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   152
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   730
         Top             =   3930
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   151
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   729
         Top             =   3570
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   148
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   738
         Text            =   "Text5"
         Top             =   2430
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   150
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   737
         Text            =   "Text5"
         Top             =   2820
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   150
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   728
         Top             =   2790
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   148
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   727
         Top             =   2430
         Width           =   750
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   20
         Left            =   5100
         TabIndex        =   732
         Top             =   4650
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepInfATRIA 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   731
         Top             =   4650
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   154
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   726
         Top             =   1650
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   153
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   725
         Top             =   1260
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   154
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   724
         Text            =   "Text5"
         Top             =   1650
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   153
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   723
         Text            =   "Text5"
         Top             =   1260
         Width           =   3375
      End
      Begin VB.CommandButton Command50 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":17131
         Style           =   1  'Graphical
         TabIndex        =   722
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command49 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1743B
         Style           =   1  'Graphical
         TabIndex        =   721
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   101
         Left            =   1620
         MouseIcon       =   "frmListado.frx":17745
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3930
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   100
         Left            =   1620
         MouseIcon       =   "frmListado.frx":17897
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3570
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   218
         Left            =   660
         TabIndex        =   746
         Top             =   3390
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   217
         Left            =   945
         TabIndex        =   745
         Top             =   3990
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   215
         Left            =   945
         TabIndex        =   744
         Top             =   3630
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   99
         Left            =   1650
         MouseIcon       =   "frmListado.frx":179E9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   98
         Left            =   1650
         MouseIcon       =   "frmListado.frx":17B3B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
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
         Index           =   213
         Left            =   690
         TabIndex        =   741
         Top             =   2250
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   212
         Left            =   975
         TabIndex        =   740
         Top             =   2850
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   211
         Left            =   975
         TabIndex        =   739
         Top             =   2490
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   221
         Left            =   960
         TabIndex        =   736
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   220
         Left            =   960
         TabIndex        =   735
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label30 
         Caption         =   "Informe de Miembros ATRIA"
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
         TabIndex        =   734
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
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
         Index           =   214
         Left            =   675
         TabIndex        =   733
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   105
         Left            =   1650
         MouseIcon       =   "frmListado.frx":17C8D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1650
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   104
         Left            =   1650
         MouseIcon       =   "frmListado.frx":17DDF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1260
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6750
      Top             =   5220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameGeneracionEntradasSIN 
      Height          =   3690
      Left            =   0
      TabIndex        =   646
      Top             =   -60
      Width           =   6615
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   135
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   649
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command43 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":17F31
         Style           =   1  'Graphical
         TabIndex        =   653
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command42 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1823B
         Style           =   1  'Graphical
         TabIndex        =   652
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton cmdAcepGenEntradas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   650
         Top             =   2865
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   5280
         TabIndex        =   651
         Top             =   2865
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   137
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   648
         Top             =   1965
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   136
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   647
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   27
         Left            =   1605
         Picture         =   "frmListado.frx":18545
         ToolTipText     =   "Buscar fecha"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F. Albar�n"
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
         Index           =   194
         Left            =   630
         TabIndex        =   658
         Top             =   2520
         Width           =   750
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   26
         Left            =   1620
         Picture         =   "frmListado.frx":185D0
         ToolTipText     =   "Buscar fecha"
         Top             =   1980
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   25
         Left            =   1620
         Picture         =   "frmListado.frx":1865B
         ToolTipText     =   "Buscar fecha"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label26 
         Caption         =   "Generaci�n Entradas Facturas"
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
         TabIndex        =   657
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   198
         Left            =   975
         TabIndex        =   656
         Top             =   1965
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   197
         Left            =   975
         TabIndex        =   655
         Top             =   1620
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   196
         Left            =   645
         TabIndex        =   654
         Top             =   1380
         Width           =   435
      End
   End
   Begin VB.Frame FrameTraspasoFactCoop 
      Height          =   5490
      Left            =   0
      TabIndex        =   234
      Top             =   -60
      Width           =   6615
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   45
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   251
         Text            =   "Text5"
         Top             =   1095
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   45
         Left            =   1890
         MaxLength       =   2
         TabIndex        =   250
         Top             =   1095
         Width           =   750
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   7
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   259
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   4380
         Width           =   2115
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   255
         Top             =   2580
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   256
         Top             =   2985
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelTrasCoop 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5280
         TabIndex        =   261
         Top             =   4695
         Width           =   975
      End
      Begin VB.CommandButton cmdAcepTrasCoop 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   4
         Left            =   4200
         TabIndex        =   260
         Top             =   4695
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   48
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   252
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   254
         Top             =   2025
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   48
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   238
         Text            =   "Text5"
         Top             =   1665
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   49
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   237
         Text            =   "Text5"
         Top             =   2025
         Width           =   3375
      End
      Begin VB.CommandButton Command14 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":186E6
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command13 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":189F0
         Style           =   1  'Graphical
         TabIndex        =   235
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   258
         Top             =   3930
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   257
         Top             =   3540
         Width           =   1065
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1590
         MouseIcon       =   "frmListado.frx":18CFA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cooperativa"
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
         Index           =   60
         Left            =   630
         TabIndex        =   253
         Top             =   990
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Factura"
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
         Index           =   54
         Left            =   630
         TabIndex        =   249
         Top             =   4425
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   68
         Left            =   645
         TabIndex        =   248
         Top             =   2400
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   67
         Left            =   975
         TabIndex        =   247
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   65
         Left            =   975
         TabIndex        =   246
         Top             =   2985
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   64
         Left            =   930
         TabIndex        =   245
         Top             =   1710
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   63
         Left            =   930
         TabIndex        =   244
         Top             =   2070
         Width           =   420
      End
      Begin VB.Label Label10 
         Caption         =   "Traspaso Facturas Cooperativa"
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
         TabIndex        =   243
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
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
         Index           =   59
         Left            =   645
         TabIndex        =   242
         Top             =   1470
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   30
         Left            =   1590
         MouseIcon       =   "frmListado.frx":18E4C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   29
         Left            =   1590
         MouseIcon       =   "frmListado.frx":18F9E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1590
         Picture         =   "frmListado.frx":190F0
         ToolTipText     =   "Buscar fecha"
         Top             =   2985
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1590
         Picture         =   "frmListado.frx":1917B
         ToolTipText     =   "Buscar fecha"
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
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
         Index           =   57
         Left            =   675
         TabIndex        =   241
         Top             =   3375
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   56
         Left            =   1005
         TabIndex        =   240
         Top             =   3615
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   55
         Left            =   1005
         TabIndex        =   239
         Top             =   4005
         Width           =   420
      End
   End
   Begin VB.Frame FrameBonificaciones 
      Height          =   4800
      Left            =   0
      TabIndex        =   378
      Top             =   -30
      Width           =   6615
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2685
         Left            =   180
         TabIndex        =   391
         Top             =   1680
         Width           =   3615
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   74
            Left            =   1755
            MaxLength       =   10
            TabIndex        =   380
            Top             =   225
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   76
            Left            =   1755
            MaxLength       =   5
            TabIndex        =   381
            Tag             =   "Porcentaje Bonificaci�n|N|N|||rbonifentradas|porcbonif|#,##0||"
            Top             =   795
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   77
            Left            =   1755
            MaxLength       =   10
            TabIndex        =   382
            Top             =   1350
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   78
            Left            =   1755
            MaxLength       =   10
            TabIndex        =   383
            Top             =   1890
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio"
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
            Index           =   91
            Left            =   480
            TabIndex        =   395
            Top             =   240
            Width           =   855
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   13
            Left            =   1440
            Picture         =   "frmListado.frx":19206
            ToolTipText     =   "Buscar fecha"
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nro.D�as"
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
            Index           =   106
            Left            =   480
            TabIndex        =   394
            Top             =   810
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje inicio"
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
            Index           =   107
            Left            =   480
            TabIndex        =   393
            Top             =   1350
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Indice Correcci�n"
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
            Index           =   108
            Left            =   480
            TabIndex        =   392
            Top             =   1920
            Width           =   1245
         End
      End
      Begin VB.CommandButton Command26 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":19291
         Style           =   1  'Graphical
         TabIndex        =   388
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command25 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1959B
         Style           =   1  'Graphical
         TabIndex        =   387
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   75
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   386
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   75
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   379
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepBonifica 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4020
         TabIndex        =   384
         Top             =   3945
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5130
         TabIndex        =   385
         Top             =   3945
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1620
         MouseIcon       =   "frmListado.frx":198A5
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   105
         Left            =   660
         TabIndex        =   390
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label Label15 
         Caption         =   "Alta Masiva Bonificaciones"
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
         TabIndex        =   389
         Top             =   420
         Width           =   5025
      End
   End
   Begin VB.Frame FrameGrabacionAgriweb 
      Height          =   6735
      Left            =   0
      TabIndex        =   144
      Top             =   0
      Width           =   6675
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   32
         Left            =   2610
         MaxLength       =   5
         TabIndex        =   118
         Tag             =   "Campol|N|S|0|99.99|clientes|codposta|00.00||"
         Top             =   5400
         Width           =   1200
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   109
         Top             =   1830
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   110
         Top             =   2205
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   160
         Text            =   "Text5"
         Top             =   1830
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   26
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   159
         Text            =   "Text5"
         Top             =   2205
         Width           =   3675
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   4
         Left            =   2610
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   28
         Left            =   2610
         MaxLength       =   9
         TabIndex        =   114
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3870
         Width           =   1200
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   29
         Left            =   2610
         MaxLength       =   13
         TabIndex        =   115
         Tag             =   "Campol|N|S|||clientes|codposta|#,###,###,###||"
         Top             =   4260
         Width           =   1200
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   111
         Top             =   2580
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   22
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   154
         Text            =   "Text5"
         Top             =   2595
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   30
         Left            =   2610
         MaxLength       =   10
         TabIndex        =   116
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   4650
         Width           =   1200
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   31
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   117
         Tag             =   "Campol|N|S|||clientes|codposta|#,##0.00||"
         Top             =   5025
         Width           =   1200
      End
      Begin VB.CommandButton CmdCancelAgri 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5250
         TabIndex        =   120
         Top             =   6060
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepAgri 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3990
         TabIndex        =   119
         Top             =   6060
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   107
         Top             =   975
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   108
         Top             =   1350
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   146
         Text            =   "Text5"
         Top             =   975
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   145
         Text            =   "Text5"
         Top             =   1350
         Width           =   3675
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   27
         Left            =   2610
         MaxLength       =   4
         TabIndex        =   112
         Tag             =   "Campol|N|S|||clientes|codposta|0000||"
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "Precio Estipulado Compra"
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
         Height          =   285
         Index           =   39
         Left            =   390
         TabIndex        =   164
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   780
         TabIndex        =   163
         Top             =   1860
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   780
         TabIndex        =   162
         Top             =   2235
         Width           =   420
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
         Index           =   2
         Left            =   390
         TabIndex        =   161
         Top             =   1620
         Width           =   390
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1425
         MouseIcon       =   "frmListado.frx":199F7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1830
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   1425
         MouseIcon       =   "frmListado.frx":19B49
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2205
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Superficie Total Contrato"
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
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   158
         Top             =   5055
         Width           =   2025
      End
      Begin VB.Label Label4 
         Caption         =   "CIF Industria transformadora"
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
         Height          =   285
         Index           =   29
         Left            =   390
         TabIndex        =   157
         Top             =   3870
         Width           =   2595
      End
      Begin VB.Label Label4 
         Caption         =   "Kgs. Contratados"
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
         Height          =   285
         Index           =   36
         Left            =   390
         TabIndex        =   156
         Top             =   4260
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
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
         Left            =   390
         TabIndex        =   155
         Top             =   2610
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   22
         Left            =   1410
         MouseIcon       =   "frmListado.frx":19C9B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   2595
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Grabaci�n Fichero Agriweb"
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
         Left            =   330
         TabIndex        =   153
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Formalizaci�n"
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
         Height          =   255
         Index           =   30
         Left            =   390
         TabIndex        =   152
         Top             =   4680
         Width           =   1485
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   11
         Left            =   2250
         Picture         =   "frmListado.frx":19DED
         ToolTipText     =   "Buscar fecha"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   33
         Left            =   795
         TabIndex        =   151
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   34
         Left            =   795
         TabIndex        =   150
         Top             =   1380
         Width           =   420
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
         Index           =   35
         Left            =   390
         TabIndex        =   149
         Top             =   765
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   1440
         MouseIcon       =   "frmListado.frx":19E78
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   975
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1440
         MouseIcon       =   "frmListado.frx":19FCA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Ejercicio"
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
         Height          =   285
         Index           =   27
         Left            =   390
         TabIndex        =   148
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Producto seg�n tabla"
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
         Height          =   315
         Index           =   28
         Left            =   390
         TabIndex        =   147
         Top             =   3480
         Width           =   1665
      End
   End
   Begin VB.Frame FrameVentaFruta 
      Height          =   6690
      Left            =   0
      TabIndex        =   537
      Top             =   -30
      Width           =   6615
      Begin VB.CheckBox Check18 
         Caption         =   "Salida a P�gina Excel"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   636
         Top             =   4650
         Width           =   2025
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   118
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   556
         Text            =   "Text5"
         Top             =   2580
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   117
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   555
         Text            =   "Text5"
         Top             =   2220
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   118
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   542
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   117
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   541
         Top             =   2220
         Width           =   735
      End
      Begin VB.CommandButton Command37 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1A11C
         Style           =   1  'Graphical
         TabIndex        =   554
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command36 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1A426
         Style           =   1  'Graphical
         TabIndex        =   553
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   116
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   552
         Text            =   "Text5"
         Top             =   3570
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   115
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   551
         Text            =   "Text5"
         Top             =   3210
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   116
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   544
         Top             =   3570
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   115
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   543
         Top             =   3210
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   114
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   550
         Text            =   "Text5"
         Top             =   1635
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   113
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   549
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   114
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   540
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   113
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   539
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepVtaFruta 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4140
         TabIndex        =   547
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   5220
         TabIndex        =   548
         Top             =   5880
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   110
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   546
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   109
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   545
         Top             =   4140
         Width           =   1095
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Detallar Albaranes"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   538
         Top             =   4260
         Width           =   2025
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   360
         TabIndex        =   557
         Top             =   5220
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   78
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1A730
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   77
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1A882
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   172
         Left            =   1005
         TabIndex        =   571
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   171
         Left            =   1005
         TabIndex        =   570
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   170
         Left            =   675
         TabIndex        =   569
         Top             =   2025
         Width           =   495
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   19
         Left            =   1620
         Picture         =   "frmListado.frx":1A9D4
         ToolTipText     =   "Buscar fecha"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   18
         Left            =   1620
         Picture         =   "frmListado.frx":1AA5F
         ToolTipText     =   "Buscar fecha"
         Top             =   4140
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   76
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1AAEA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   75
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1AC3C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   74
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1AD8E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   73
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1AEE0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   169
         Left            =   675
         TabIndex        =   568
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   168
         Left            =   1005
         TabIndex        =   567
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   167
         Left            =   1005
         TabIndex        =   566
         Top             =   3255
         Width           =   465
      End
      Begin VB.Label Label22 
         Caption         =   "Listado Comprobaci�n Venta Fruta"
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
         TabIndex        =   565
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   166
         Left            =   675
         TabIndex        =   564
         Top             =   3015
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   165
         Left            =   960
         TabIndex        =   563
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   164
         Left            =   960
         TabIndex        =   562
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   163
         Left            =   1005
         TabIndex        =   561
         Top             =   4545
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   162
         Left            =   1005
         TabIndex        =   560
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   159
         Left            =   675
         TabIndex        =   559
         Top             =   3960
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         Height          =   195
         Index           =   154
         Left            =   360
         TabIndex        =   558
         Top             =   5700
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Frame FrameEntradasPesada 
      Height          =   5715
      Left            =   0
      TabIndex        =   342
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check13 
         Caption         =   "Imprimir Resumen"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   353
         Top             =   4140
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   73
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   352
         Top             =   4545
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   72
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   351
         Top             =   4140
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5235
         TabIndex        =   355
         Top             =   4815
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   4
         Left            =   4155
         TabIndex        =   354
         Top             =   4815
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   71
         Left            =   1935
         MaxLength       =   7
         TabIndex        =   346
         Tag             =   "Pesadal|N|S|||clientes|nropesada|0000000||"
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   70
         Left            =   1935
         MaxLength       =   7
         TabIndex        =   345
         Tag             =   "Pesadal|N|S|||clientes|nropesada|0000000||"
         Top             =   1260
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   69
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   350
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   68
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   349
         Top             =   3195
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   68
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   359
         Text            =   "Text5"
         Top             =   3210
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   69
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   358
         Text            =   "Text5"
         Top             =   3600
         Width           =   3375
      End
      Begin VB.CommandButton Command20 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1B032
         Style           =   1  'Graphical
         TabIndex        =   357
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command15 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1B33C
         Style           =   1  'Graphical
         TabIndex        =   356
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   67
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   348
         Top             =   2610
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   347
         Top             =   2205
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   66
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   344
         Text            =   "Text5"
         Top             =   2205
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   67
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   343
         Text            =   "Text5"
         Top             =   2610
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   104
         Left            =   675
         TabIndex        =   372
         Top             =   3960
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   103
         Left            =   1005
         TabIndex        =   371
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   102
         Left            =   1005
         TabIndex        =   370
         Top             =   4545
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   101
         Left            =   960
         TabIndex        =   369
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   100
         Left            =   960
         TabIndex        =   368
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   99
         Left            =   675
         TabIndex        =   367
         Top             =   3015
         Width           =   630
      End
      Begin VB.Label Label14 
         Caption         =   "Informe de Entradas de Pesadas"
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
         TabIndex        =   366
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   98
         Left            =   1005
         TabIndex        =   365
         Top             =   3255
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   97
         Left            =   1005
         TabIndex        =   364
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pesada"
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
         Index           =   96
         Left            =   675
         TabIndex        =   363
         Top             =   1080
         Width           =   525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   69
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1B646
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   68
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1B798
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3195
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   12
         Left            =   1620
         Picture         =   "frmListado.frx":1B8EA
         ToolTipText     =   "Buscar fecha"
         Top             =   4545
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   10
         Left            =   1620
         Picture         =   "frmListado.frx":1B975
         ToolTipText     =   "Buscar fecha"
         Top             =   4140
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   95
         Left            =   675
         TabIndex        =   362
         Top             =   2025
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   94
         Left            =   1005
         TabIndex        =   361
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   93
         Left            =   1005
         TabIndex        =   360
         Top             =   2655
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   67
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1BA00
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   66
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1BB52
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2205
         Width           =   240
      End
   End
   Begin VB.Frame FrameTraspasoAAlmazara 
      Height          =   3450
      Left            =   0
      TabIndex        =   327
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Command22 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1BCA4
         Style           =   1  'Graphical
         TabIndex        =   335
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command21 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1BFAE
         Style           =   1  'Graphical
         TabIndex        =   334
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   65
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   332
         Text            =   "Text5"
         Top             =   1695
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   64
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   330
         Text            =   "Text5"
         Top             =   1335
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   329
         Top             =   1695
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   64
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   328
         ToolTipText     =   " "
         Top             =   1335
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepTrasDatosAlmz 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   331
         Top             =   2535
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   5130
         TabIndex        =   333
         Top             =   2535
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   49
         Left            =   1530
         MouseIcon       =   "frmListado.frx":1C2B8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1740
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   48
         Left            =   1530
         MouseIcon       =   "frmListado.frx":1C40A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   88
         Left            =   585
         TabIndex        =   339
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Traspaso Datos a Almazara"
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
         TabIndex        =   338
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   82
         Left            =   870
         TabIndex        =   337
         Top             =   1740
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   80
         Left            =   870
         TabIndex        =   336
         Top             =   1380
         Width           =   465
      End
   End
   Begin VB.Frame FrameListOrdenesEmitidas 
      Height          =   4575
      Left            =   0
      TabIndex        =   680
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   144
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   687
         Top             =   2730
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   144
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   696
         Text            =   "Text5"
         Top             =   2730
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   143
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   686
         Top             =   2355
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   143
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   693
         Text            =   "Text5"
         Top             =   2355
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   140
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   685
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   139
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   684
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   5070
         TabIndex        =   691
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepOrdEmitidas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3990
         TabIndex        =   689
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton Command46 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1C55C
         Style           =   1  'Graphical
         TabIndex        =   682
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command45 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1C866
         Style           =   1  'Graphical
         TabIndex        =   681
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   206
         Left            =   990
         TabIndex        =   697
         Top             =   2805
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   88
         Left            =   1605
         MouseIcon       =   "frmListado.frx":1CB70
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   205
         Left            =   660
         TabIndex        =   695
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   203
         Left            =   990
         TabIndex        =   694
         Top             =   2400
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   87
         Left            =   1605
         MouseIcon       =   "frmListado.frx":1CCC2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2355
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   202
         Left            =   660
         TabIndex        =   692
         Top             =   1230
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   200
         Left            =   990
         TabIndex        =   690
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   199
         Left            =   990
         TabIndex        =   688
         Top             =   1815
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   30
         Left            =   1575
         Picture         =   "frmListado.frx":1CE14
         ToolTipText     =   "Buscar fecha"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   29
         Left            =   1575
         Picture         =   "frmListado.frx":1CE9F
         ToolTipText     =   "Buscar fecha"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label28 
         Caption         =   "Informe de Ordenes Emitidas"
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
         TabIndex        =   683
         Top             =   405
         Width           =   5025
      End
   End
   Begin VB.Frame FrameCalidades 
      Height          =   4455
      Left            =   0
      TabIndex        =   24
      Top             =   30
      Width           =   7020
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   42
         Top             =   3735
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   40
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   34
         Top             =   1275
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   35
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "Text5"
         Top             =   1275
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   1635
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   36
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   38
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text5"
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text5"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CommandButton Command8 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1CF2A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command7 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1D234
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ordenar por"
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
         Height          =   975
         Left            =   495
         TabIndex        =   25
         Top             =   3195
         Width           =   2190
         Begin VB.OptionButton Opcion 
            Caption         =   "Variedad"
            Height          =   345
            Index           =   2
            Left            =   495
            TabIndex        =   27
            Top             =   225
            Width           =   1290
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Calidad"
            Height          =   255
            Index           =   3
            Left            =   495
            TabIndex        =   26
            Top             =   585
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   2
         Left            =   5265
         TabIndex        =   37
         Top             =   405
         Visible         =   0   'False
         Width           =   1050
         _ExtentX        =   1852
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
         Height          =   195
         Index           =   35
         Left            =   960
         TabIndex        =   47
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   34
         Left            =   960
         TabIndex        =   46
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Left            =   600
         TabIndex        =   45
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label8 
         Caption         =   "Informe de Calidades"
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
         Left            =   540
         TabIndex        =   44
         Top             =   405
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   32
         Left            =   960
         TabIndex        =   43
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   31
         Left            =   960
         TabIndex        =   41
         Top             =   2790
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Calidad"
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
         Index           =   30
         Left            =   600
         TabIndex        =   39
         Top             =   2160
         Width           =   525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1575
         MouseIcon       =   "frmListado.frx":1D53E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1575
         MouseIcon       =   "frmListado.frx":1D690
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1575
         MouseIcon       =   "frmListado.frx":1D7E2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar calidad"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1575
         MouseIcon       =   "frmListado.frx":1D934
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar calidad"
         Top             =   2430
         Width           =   240
      End
   End
   Begin VB.Frame FrameRegFitosanitario 
      Height          =   5745
      Left            =   0
      TabIndex        =   795
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   180
         Left            =   1980
         MaxLength       =   20
         TabIndex        =   808
         Text            =   "12345678901234566789"
         Top             =   4530
         Width           =   2055
      End
      Begin VB.CommandButton Command56 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1DA86
         Style           =   1  'Graphical
         TabIndex        =   816
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command55 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1DD90
         Style           =   1  'Graphical
         TabIndex        =   815
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   175
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   814
         Text            =   "Text5"
         Top             =   1995
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   176
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   813
         Text            =   "Text5"
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   176
         Left            =   1980
         MaxLength       =   3
         TabIndex        =   803
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   175
         Left            =   1980
         MaxLength       =   3
         TabIndex        =   802
         Top             =   2010
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   173
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   812
         Text            =   "Text5"
         Top             =   1155
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   174
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   811
         Text            =   "Text5"
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   174
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   801
         Top             =   1560
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   173
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   800
         Top             =   1170
         Width           =   750
      End
      Begin VB.CommandButton cmdAcepInfFito 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   809
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   23
         Left            =   5160
         TabIndex        =   810
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   167
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   799
         Text            =   "Text5"
         Top             =   2850
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   168
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   798
         Text            =   "Text5"
         Top             =   3255
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   168
         Left            =   1965
         MaxLength       =   4
         TabIndex        =   805
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   167
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   804
         Top             =   2835
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   160
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   807
         Top             =   4170
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   159
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   797
         Text            =   "Text5"
         Top             =   3780
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   159
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   806
         Top             =   3780
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   160
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   796
         Text            =   "Text5"
         Top             =   4170
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campa�a"
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
         Index           =   265
         Left            =   630
         TabIndex        =   881
         Top             =   4560
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   121
         Left            =   1590
         MouseIcon       =   "frmListado.frx":1E09A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   120
         Left            =   1590
         MouseIcon       =   "frmListado.frx":1E1EC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   2010
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   119
         Left            =   1590
         MouseIcon       =   "frmListado.frx":1E33E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1590
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   118
         Left            =   1590
         MouseIcon       =   "frmListado.frx":1E490
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   260
         Left            =   630
         TabIndex        =   829
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   259
         Left            =   960
         TabIndex        =   828
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   258
         Left            =   960
         TabIndex        =   827
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label33 
         Caption         =   "Registro Aplicaci�n Fitosanitarios"
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
         TabIndex        =   826
         Top             =   405
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
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
         Index           =   257
         Left            =   630
         TabIndex        =   825
         Top             =   1845
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   256
         Left            =   960
         TabIndex        =   824
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   255
         Left            =   960
         TabIndex        =   823
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   113
         Left            =   1590
         MouseIcon       =   "frmListado.frx":1E5E2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar partida"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   112
         Left            =   1590
         MouseIcon       =   "frmListado.frx":1E734
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar partida"
         Top             =   2835
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   236
         Left            =   960
         TabIndex        =   822
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   235
         Left            =   960
         TabIndex        =   821
         Top             =   2895
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Partida"
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
         Index           =   234
         Left            =   630
         TabIndex        =   820
         Top             =   2700
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "T�rmino Municipal"
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
         Index           =   232
         Left            =   630
         TabIndex        =   819
         Top             =   3600
         Width           =   1260
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   231
         Left            =   960
         TabIndex        =   818
         Top             =   3825
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   107
         Left            =   1590
         MouseIcon       =   "frmListado.frx":1E886
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   230
         Left            =   960
         TabIndex        =   817
         Top             =   4200
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   106
         Left            =   1590
         MouseIcon       =   "frmListado.frx":1E9D8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   3810
         Width           =   240
      End
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   79
      Left            =   960
      MouseIcon       =   "frmListado.frx":1EB2A
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar seccion"
      Top             =   135
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Secci�n"
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   192
      Left            =   0
      TabIndex        =   625
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "frmListado"
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
    ' 10 .- Reimpresion de Entradas de Bascula
    ' 11 .- Listado de Entradas de Pesadas
    ' 12 .- Listado de Calidades
    ' 13 .- Listado de Socios por Secci�n
    ' 14 .- Listado de Entradas en Bascula
    ' 15 .- Listado de Campos
    ' 16 .- Listado de Entradas clasificacion
    ' 17 .- Reimpresion de albaranes de Clasificacion
    ' 18 .- Informe de Kilos/Gastos (rhisfruta)
    ' 19 .- Grabaci�n de Fichero Agriweb
    ' 20 .- Informe de Kilos Por Producto
    ' 21 .- Traspaso desde el calibrador
    ' 22 .- Traspaso TRAZABILIDAD
    
    
    ' 23 .- Baja de Socios (dentro del mantenimiento socios)
    
    ' 24 .- Traspaso de Facturas Cooperativa ( traspaso liquidacion )
    ' 25 .- Listado de Kilos recolectados socio / cooperativa
    ' 26 .- Traspaso de ROPAS solo para Catadau
    ' 27 .- Traspaso de datos a Almazara solo para Mogente
    
    ' 28 .- Alta Masiva de bonificaciones de entradas
    ' 29 .- Baja Masiva de bonificaciones de entradas
    
    ' 30 .- Generacion de clasificaci�n (solo para Picassent frmManClasAutoPic)
    
    ' 31 .- Impresion (Informe Fases) de informe de socios
            'seleccionando unicamente la fases (para Castelduc)
    
    
    ' 32 .- Impresion de control de destrio
    
    ' 33 .- Impresion de Gastos de Albaran por concepto
    '
    ' 34 .- Cambio de socio de un campo
    
    ' 35 .- Listado de comprobacion de venta fruta
    ' 36 .- Listado de Gastos por campo
    ' 37 .- Contabilizacion de gastos de campo
    ' 38 .- Cambio de nro de factura de socio
    
    ' 39 .- Carga de entradas de albaranes a partir de facturas de SIN
    
    ' 40 .- Impresion de ordenes de recoleccion. ALZIRA
    ' 41 .- Listado de Ordenes de recoleccion emitidas. ALZIRA
    ' 42 .- Informe de socios (telefonos)
    
    ' 43 .- Informe oficial de miembros ATRIA
    
    ' 44 .- Listado de mantenimiento de precios
    
    ' 45 .- Listado de revisiones de campos
    ' 46 .- Listado de registros de fitosanitarios
    
    ' 47 .- Traspaso datos a trazabilidad (solo Castelduc)
    ' 48 .- Traspaso de albaranes de retirada de cooperativas(bolbaite,navarres..) a ABN
    
    ' 49 .- Asignacion de codigos de globalgap a campos, segun producto y partida (Catadau)
    ' 50 .- Informe de diferencias de kilos
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmSec As frmManSeccion 'Secciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'Calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmProd As frmComercial 'Ayuda Productos de comercial
Attribute frmProd.VB_VarHelpID = -1
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
Private WithEvents frmMens1 As frmMensajes 'Mensajes
Attribute frmMens1.VB_VarHelpID = -1
Private WithEvents frmMens2 As frmMensajes 'Mensajes
Attribute frmMens2.VB_VarHelpID = -1
Private WithEvents frmMens3 As frmMensajes 'Mensajes
Attribute frmMens3.VB_VarHelpID = -1
Private WithEvents frmMens4 As frmMensajes 'Mensajes
Attribute frmMens4.VB_VarHelpID = -1
Private WithEvents frmMens5 As frmMensajes 'Mensajes
Attribute frmMens5.VB_VarHelpID = -1
Private WithEvents frmMens6 As frmMensajes 'Mensajes
Attribute frmMens6.VB_VarHelpID = -1
Private WithEvents frmMens7 As frmMensajes 'Mensajes
Attribute frmMens7.VB_VarHelpID = -1
Private WithEvents frmMens8 As frmMensajes 'Mensajes
Attribute frmMens8.VB_VarHelpID = -1
Private WithEvents frmSitu As frmManSituacion 'Situacion de socio
Attribute frmSitu.VB_VarHelpID = -1
Private WithEvents frmCoop As frmManCoope 'Cooperativa
Attribute frmCoop.VB_VarHelpID = -1
Private WithEvents frmCapa As frmManCapataz 'capataces
Attribute frmCapa.VB_VarHelpID = -1
Private WithEvents frmPar As frmManPartidas 'partidas
Attribute frmPar.VB_VarHelpID = -1
Private WithEvents frmPue As frmManPueblos 'pueblos
Attribute frmPue.VB_VarHelpID = -1
Private WithEvents frmCon As frmManConcepGasto 'conceptos de gastos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden 'mantenimiento de incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico 'mantenimiento de clientes de comercial
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmCampos As frmManCampos 'mantenimiento de campos para el listado de gastos por campo
Attribute frmCampos.VB_VarHelpID = -1
Private WithEvents frmZon As frmManZonas 'zonas
Attribute frmZon.VB_VarHelpID = -1

Private WithEvents frmCConta As frmConceConta 'conceptos de contabilidad
Attribute frmCConta.VB_VarHelpID = -1
Private WithEvents frmDConta As frmDiaConta 'conceptos de contabilidad
Attribute frmDConta.VB_VarHelpID = -1
Private WithEvents frmCtaConta As frmCtasConta 'cuentas contables
Attribute frmCtaConta.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Tabla1 As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim Indice As Integer

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean
Dim vSql2 As String
Dim vSeccion As CSeccion

Dim FecFacInicial As String

Dim ConPropietario As Boolean

Dim vTipoMov As CTiposMov
Dim CodTipoMov As String

Dim EsReimpresion As Boolean

'[Monica]11/11/2013: indicamos si han entrado o no por campos
Dim HayRegistros As Boolean
Dim PriFact As Long

Dim CifEmpre As String

Dim AlbaranAnterior As Long
Dim Albaran2 As Long
Dim Contratos As String

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub Check19_Click()
    Me.FrameNroOrden.visible = (Check19.Value = 1)
    EsReimpresion = (Check19.Value = 1)
    If Not FrameNroOrden.visible Then
        txtCodigo(141).Text = ""
    End If
End Sub

Private Sub Check19_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'[Monica] 01/10/2009 a�adido el poder detallar las notas
Private Sub Check2_Click()
    If OpcionListado = 18 Then
        Check9.Enabled = (Check2.Value = 0)
        If Not Check9.Enabled Then Check9.Value = 0
    End If
End Sub

Private Sub Check5_Click()
    Check6.Enabled = (Check5.Value = 1)
    Check10.Enabled = (Check5.Value = 1)
End Sub

Private Sub Check8_Click()
    If Check8.Value Then
        Check24.Enabled = True
    Else
        Check24.Enabled = False
        Check24.Value = False
    End If
End Sub

Private Sub chkBaja_Click()
    Frame9.Enabled = (chkBaja.Value = 1)
    If Not Frame9.Enabled Then
        txtCodigo(181).Text = ""
        txtNombre(181).Text = ""
    End If
End Sub

Private Sub CmdAcepAgri_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String

    If Not DatosOK Then Exit Sub


    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
     'D/H Socio
     cDesde = Trim(txtCodigo(23).Text)
     cHasta = Trim(txtCodigo(24).Text)
     nDesde = txtNombre(23).Text
     nHasta = txtNombre(24).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
     End If

     'D/H Clase
     cDesde = Trim(txtCodigo(25).Text)
     cHasta = Trim(txtCodigo(26).Text)
     nDesde = txtNombre(25).Text
     nHasta = txtNombre(26).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codclase}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
     End If
     
    ' PRODUCTO
    If txtCodigo(22).Text <> "" Then
        If Not AnyadirAFormula(cadSelect, "{variedades.codprodu} = " & DBSet(txtCodigo(22).Text, "N")) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{variedades.codprodu}  = " & DBSet(txtCodigo(22).Text, "N")) Then Exit Sub
    End If
     
     If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
     If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub

     tabla = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
     tabla = "(" & tabla & ") INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
     
     vSQL = ""
     If txtCodigo(25).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(25).Text, "N")
     If txtCodigo(26).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(26).Text, "N")
     If txtCodigo(22).Text <> "" Then vSQL = vSQL & " and variedades.codprodu = " & DBSet(txtCodigo(22).Text, "N")
     Set frmMens = New frmMensajes
     
     frmMens.OpcionMensaje = 16
     frmMens.cadWHERE = vSQL
     frmMens.Show vbModal
     
     Set frmMens = Nothing
     
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If HayRegParaInforme(tabla, cadSelect) Then
        B = GeneraFicheroAgriweb(tabla, cadSelect)
        If B Then
            If CopiarFichero Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                CmdCancelAgri_Click
            End If
        End If
     End If

End Sub

Private Sub cmdAcepAsigGlobalgap_Click()
Dim sql As String
Dim RS As ADODB.Recordset
Dim RsPar As ADODB.Recordset
Dim RsPro As ADODB.Recordset
Dim GGPar As String
Dim GGPro As String

    On Error GoTo eAsignacion

    conn.BeginTrans


    sql = "select count(*) from rcampos where fecbajas is null  "
    
    CargarProgres pb10, TotalRegistros(sql)
    Me.pb10.visible = True
    
    Label2(238).Caption = "Calculando registros GlobalGap"
    Label2(262).Caption = ""
    Label2(238).visible = True
    Label2(262).visible = True
    
    
    
    
    sql = "select * from rcampos where fecbajas is null order by codcampo "
    
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Label2(262).Caption = "Campo : " & RS!codcampo
        IncrementarProgres pb10, 1
        Me.Refresh
    
        GGPar = ""
        GGPro = ""
        sql = "select globalgap from rpartida where codparti = " & DBSet(RS!codparti, "N")
        
        Set RsPar = New ADODB.Recordset
        RsPar.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not RsPar.EOF Then
            GGPar = DBLet(RsPar!globalgap, "T")
        End If
        Set RsPar = Nothing
        
        
        sql = "select globalgap from productos, variedades where variedades.codvarie = " & DBSet(RS!codvarie, "N")
        sql = sql & " and productos.codprodu = variedades.codprodu "
        
        Set RsPro = New ADODB.Recordset
        RsPro.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not RsPro.EOF Then
            GGPro = DBLet(RsPro!globalgap, "T")
        End If
        Set RsPro = Nothing
    
        If GGPar = "" Or GGPro = "" Then
            If Check25.Value = 1 Then
                sql = "update rcampos set codigoggap = " & ValorNulo & " where codcampo = " & DBSet(RS!codcampo, "N")
                conn.Execute sql
            End If
        Else
            sql = "update rcampos set codigoggap = " & DBSet(GGPro & GGPar, "T") & " where codcampo = " & DBSet(RS!codcampo, "N")
            conn.Execute sql
        End If
    
    
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    conn.CommitTrans
    MsgBox "Proceso realizado correctamente.", vbExclamation
    Unload Me
    Exit Sub
    
eAsignacion:
    MuestraError Err.Number, "Asignaci�n Globalgap", Err.Description
    conn.RollbackTrans
End Sub

Private Sub cmdAcepBajaSocio_Click()
Dim sql As String

    On Error GoTo eErrores

    If txtCodigo(47).Text = "" Then
        MsgBox "Debe introducir la fecha de baja.", vbExclamation
        PonerFoco txtCodigo(47)
        Exit Sub
    End If
    If txtCodigo(46).Text = "" Then
        MsgBox "Debe introducir la nueva situaci�n del socio.", vbExclamation
        PonerFoco txtCodigo(46)
        Exit Sub
    End If
    
    If Me.chkBaja.Value = 1 Then
        If txtCodigo(181).Text = "" Then
            MsgBox "Debe introducir la nueva situaci�n de los campos del socio.", vbExclamation
            PonerFoco txtCodigo(46)
            Exit Sub
        End If
    End If
    
    sql = "update rsocios_seccion set fecbaja = " & DBSet(txtCodigo(47), "F")
    sql = sql & " where codsocio = " & DBSet(NumCod, "N")
    sql = sql & " and (fecbaja is null or fecbaja = '')"
    conn.Execute sql
    
    sql = "update rsocios set codsitua = " & DBSet(txtCodigo(46).Text, "N")
    sql = sql & ", fechabaja = " & DBSet(txtCodigo(47), "F")
    sql = sql & " where codsocio = " & DBSet(NumCod, "N")
    conn.Execute sql
    
    If Me.chkBaja.Value = 1 Then
        sql = "update rcampos set fecbajas = " & DBSet(txtCodigo(47), "F")
        sql = sql & ", codsitua = " & DBSet(txtCodigo(181), "N")
        sql = sql & " where codsocio = " & DBSet(NumCod, "N")
        sql = sql & " and (fecbajas is null or fecbajas = '')"
        
        conn.Execute sql
    End If
    
    
    MsgBox "Proceso realizado correctamente.", vbExclamation
    cmdCancelBajaSocio_Click
    Exit Sub

eErrores:
    MuestraError Err.Number, "Baja de Socio", Err.Description
End Sub

Private Sub CmdAcepBonifica_Click()
Dim sql As String
Dim Sql2 As String
Dim Porcentaje As Currency
Dim I As Long

    If DatosOK Then
        Select Case OpcionListado
            Case 28
                If InsertarBonificaciones Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (6)
                End If
            Case 29
                If EliminarBonificaciones Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (6)
                End If
            
        End Select
    End If

End Sub

Private Sub cmdAcepCambsoc_Click(Index As Integer)
Dim SocioCampo As String

    ' Cambio de Socio de un campo
    
    If txtCodigo(111).Text = "" Then
         MsgBox "Debe introducir un nuevo socio para el campo.", vbExclamation
         PonerFoco txtCodigo(111)
         Exit Sub
    Else
        If CLng(DevuelveValor("select codsocio from rcampos where codcampo = " & DBSet(NumCod, "N"))) = CLng(txtCodigo(111).Text) Then
            MsgBox "El c�digo de socio coincide con el actual del campo. Revise.", vbExclamation
            PonerFoco txtCodigo(111)
            Exit Sub
        Else
            If TotalRegistros("select count(*) from rsocios_seccion where codsocio = " & DBSet(txtCodigo(111).Text, "N") & " and fecbaja is null and codsecci = " & vParamAplic.Seccionhorto) = 0 Then
                MsgBox "El nuevo Socio no existe en la Seccion de Horto o est� dado de baja.", vbExclamation
                PonerFoco txtCodigo(111)
                Exit Sub
            End If
        End If
    End If
    
    If txtCodigo(106).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una fecha de cambio para el campo. Revise.", vbExclamation
        PonerFoco txtCodigo(106)
        Exit Sub
    Else
        If CDate(txtCodigo(106).Text) < CDate(DevuelveValor("select fecalta from rcampos where codcampo = " & DBSet(NumCod, "N"))) Then
            MsgBox "La fecha de cambio ha de ser superior a la fecha de alta del campo. Revise.", vbExclamation
            PonerFoco txtCodigo(106)
            Exit Sub
        End If
    End If
        
    If txtCodigo(107).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una incidencia. Revise.", vbExclamation
        PonerFoco txtCodigo(107)
        Exit Sub
    Else
        txtNombre(107).Text = PonerNombreDeCod(txtCodigo(107), "rincidencia", "nomincid", "codincid", "N")
        If txtNombre(107).Text = "" Then
            MsgBox "El c�digo de incidencia no existe. Revise.", vbExclamation
            PonerFoco txtCodigo(107)
            Exit Sub
        End If
    End If

    SocioCampo = DevuelveValor("select codsocio from rcampos where codcampo = " & DBSet(NumCod, "N"))

    ConPropietario = False
    If MsgBox("� Desea cambiar tambi�n el propietario ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then ConPropietario = True

    ' comprobamos que no hayan facturas de liquidacion del socio
    If HayFacturasdelSocio(SocioCampo) Then
        If MsgBox("� Desea continuar con el proceso ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            If CambiarSocio(SocioCampo) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                Unload Me
            End If
        End If
    Else
        If CambiarSocio(SocioCampo) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            Unload Me
        End If
    End If

End Sub

Private Function CambiarSocio(SocioAnt As String)
Dim sql As String
Dim FecAltas As String
Dim NumF As Long

    On Error GoTo eCambiarSocio

    CambiarSocio = False

    conn.BeginTrans
    
    '[Monica]21/09/2012: metemos una accion en el log de que ha habido un cambio de socio
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 9, vUsu, "Campo:" & NumCod & " Socio Ant." & SocioAnt & " - Nuevo : " & txtCodigo(111).Text
    Set LOG = Nothing
    '-----------------------------------------------------------------------------

    sql = "update rentradas set codsocio = " & DBSet(txtCodigo(111).Text, "N")
    sql = sql & " where codcampo = " & DBSet(NumCod, "N")
    
    conn.Execute sql
    
    sql = "update rclasifica set codsocio = " & DBSet(txtCodigo(111).Text, "N")
    sql = sql & " where codcampo = " & DBSet(NumCod, "N")
    
    conn.Execute sql
    
    sql = "update rhisfruta set codsocio = " & DBSet(txtCodigo(111).Text, "N")
    sql = sql & " where codcampo = " & DBSet(NumCod, "N")
    
    conn.Execute sql
    
    NumF = SugerirCodigoSiguienteStr("rcampos_hco", "numlinea", "codcampo = " & DBSet(NumCod, "N"))
    FecAltas = DevuelveValor("select fecaltas from rcampos where codcampo = " & DBSet(NumCod, "N"))
    
    '[Monica]31/10/2012: en el caso de alzira no tenemos que cambiar la fecha de alta del campo, luego la fecha desde del socio
    '                    debe ser la maxima + 1 del ultimo socio (si lo hay) o la fecha de alta del campo si no lo hay
    If vParamAplic.Cooperativa = 4 Then
        sql = "select date_add(max(fechabaja), interval 1 day) from rcampos_hco where codcampo = " & DBSet(NumCod, "N")
        If DevuelveValor(sql) <> 0 Then
            FecAltas = DevuelveValor(sql)
        End If
    End If
    
    sql = "insert into rcampos_hco (codcampo, numlinea, codsocio, fechaalta, fechabaja, codincid) values ("
    sql = sql & DBSet(NumCod, "N") & "," & DBSet(NumF, "N") & "," & DBSet(SocioAnt, "N") & "," & DBSet(FecAltas, "F") & ","
    sql = sql & DBSet(txtCodigo(106).Text, "F") & "," & DBSet(txtCodigo(107).Text, "N") & ")"
    
    conn.Execute sql
    
    '[Monica]31/10/2012: no tocaremos la fecha de alta del campo
    If vParamAplic.Cooperativa = 4 Then ' la fecha de alta del campo no la toco
        sql = "update rcampos set codsocio = " & DBSet(txtCodigo(111).Text, "N")
    Else
        sql = "update rcampos set codsocio = " & DBSet(txtCodigo(111).Text, "N") & ", fecaltas = " & DBSet(txtCodigo(106).Text, "F")
    End If
    
    '[Monica]21/09/2012: hemos preguntado si quieren cambiar tambien el propietario
    If ConPropietario Then sql = sql & ", codpropiet = " & DBSet(txtCodigo(111).Text, "N")
         
    sql = sql & " where codcampo = " & DBSet(NumCod, "N")
    
    conn.Execute sql
    
    ' actualizamos la tabla de coopropietarios
    sql = "update rcampos_cooprop set codsocio = " & DBSet(txtCodigo(111).Text, "N") & " where codcampo = " & DBSet(NumCod, "N")
    sql = sql & " and codsocio = " & DBSet(SocioAnt, "N")
    
    conn.Execute sql
    
    '[Monica]08/06/2012: si la cooperativa es Escalona y cambiamos el socio del campo, me actualiza la tabla de contadores
    If vParamAplic.Cooperativa = 10 Then
        sql = "update rpozos set codsocio = " & DBSet(txtCodigo(111).Text, "N") & " where codcampo = " & DBSet(NumCod, "N")
        
        '[Monica]21/09/2012: le traemos de que hidrantes queremos los cambios
        If CadTag <> "" Then sql = sql & " and " & CadTag
        
        conn.Execute sql
    End If
    
    CambiarSocio = True
    conn.CommitTrans
    Exit Function
    
eCambiarSocio:
    MuestraError Err.Number, "Cambiar Socio", Err.Description
    conn.RollbackTrans
End Function

Private Function HayFacturasdelSocio(Socio As String) As Boolean
Dim sql As String

    HayFacturasdelSocio = False
    
    sql = "select * from rfactsoc where codsocio = " & DBSet(Socio, "N")
    
    If TotalRegistrosConsulta(sql) <> 0 Then
    
        Set frmMens = New frmMensajes
        
        frmMens.cadWHERE = "codsocio = " & DBSet(Socio, "N")
        frmMens.OpcionMensaje = 33
        frmMens.Show vbModal
        Set frmMens = Nothing
        
        HayFacturasdelSocio = True
    
    End If


End Function

Private Sub CmdAcepContaGastos_Click()
Dim cadWHERE As String
Dim sql As String
Dim B As Boolean

    cadWHERE = NumCod

    txtCodigo(108).Text = Format(DevuelveValor("select fecha from rcampos_gastos where " & cadWHERE), "dd/mm/yyyy")

    If Not DatosOkGastos(cadWHERE) Then Exit Sub

    sql = "CONGAS" 'contabilizar recibos de pozos

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (sql)
    If Not BloqueoManual(sql, "1") Then
        MsgBox "No se pueden Contabilizar Asiento de Gastos Campos. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los par�metros
    If Not ComprobarFechasConta(108) Then Exit Sub
    txtCodigo(108).Text = Format(DevuelveValor("select fecha from rcampos_gastos where " & cadWHERE), "dd/mm/yyyy")

    
    '===========================================================================
    'CONTABILIZAR ASIENTO DE GASTOS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Asiento de Gastos: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Asiento en Diario..."


    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar Asiento Gastos: " & vbCrLf & "rcampos_gastos" & vbCrLf & cadWHERE
    Set LOG = Nothing
    '-----------------------------------------------------------------------------

    '---- Pasar Asiento a la Contabilidad
    B = PasarAsientoGastoCampo(cadWHERE, Orden2, txtCodigo(108).Text, txtCodigo(128).Text, txtCodigo(112).Text, txtCodigo(119).Text)

    '---- Mostrar ListView de posibles errores (si hay)
    If B Then
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
        cmdCancel_Click (12)
    End If


End Sub

Private Sub cmdAcepCtrolDestrio_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vcad As String

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H SOCIO
    cDesde = Trim(txtCodigo(86).Text)
    cHasta = Trim(txtCodigo(87).Text)
    nDesde = txtNombre(86).Text
    nHasta = txtNombre(87).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    'D/H CLASE
    cDesde = Trim(txtCodigo(81).Text)
    cHasta = Trim(txtCodigo(82).Text)
    nDesde = txtNombre(81).Text
    nHasta = txtNombre(82).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
    End If
    
    'D/H VARIEDAD
    cDesde = Trim(txtCodigo(84).Text)
    cHasta = Trim(txtCodigo(85).Text)
    nDesde = txtNombre(84).Text
    nHasta = txtNombre(85).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
    End If

    'D/H fecha
    cDesde = Trim(txtCodigo(88).Text)
    cHasta = Trim(txtCodigo(89).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        Codigo = "{" & tabla & ".fechacla}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
            
    'D/H CAMPO
    cDesde = Trim(txtCodigo(90).Text)
    cHasta = Trim(txtCodigo(91).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codcampo}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCampo=""") Then Exit Sub
    End If
            
    nTabla = "(rcontrol INNER JOIN variedades ON rcontrol.codvarie = variedades.codvarie) "
    
    If HayRegParaInforme(nTabla, cadSelect) Then
        If CargarTemporalDatosDestrio(nTabla, cadSelect) Then
            If HayRegParaInforme("tmpexcel", "codusu = " & vUsu.Codigo) Then
                If Check14.Value = 0 Then
                    cadNombreRPT = "rControlDestrio.rpt"
                Else
                    cadNombreRPT = "rControlDestrioRes.rpt"
                End If
                cadTitulo = "Resumen Control Destrio"
                
                cadFormula = "{tmpexcel.codusu} = " & vUsu.Codigo
                
                LlamarImprimir
            End If
        End If
    End If
End Sub

Private Sub CmdAcepDatosTraza_Click()
Dim B As Boolean

    B = GeneraFichero()
    
    If B Then
        If CopiarFicheroTraza Then
            MsgBox "Proceso realizado correctamente", vbExclamation
            Unload Me
        End If
    Else
    
        MsgBox "No se ha realizado el proceso.", vbExclamation
    
    End If

End Sub

Public Function CopiarFicheroTraza() As Boolean
Dim nomFich As String
Dim cadena As String
On Error GoTo ecopiarfichero

    CopiarFicheroTraza = False
    ' abrimos el commondialog para indicar donde guardarlo
    Me.CommonDialog1.InitDir = App.Path
    Me.CommonDialog1.DefaultExt = "txt"
    'cadena = Format(CDate(txtCodigo(2).Text), FormatoFecha)
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    CommonDialog1.FileName = "trazabilidad.txt"
    
    Me.CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        FileCopy App.Path & "\trazabilidad.txt", CommonDialog1.FileName
        CopiarFicheroTraza = True
    End If

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear

End Function



Private Function GeneraFichero() As Boolean
Dim NFich As Integer
Dim RS As ADODB.Recordset
Dim cad As String
Dim sql As String
Dim I As Integer
Dim vSocio As cSocio
Dim v_total As Currency
Dim v_total1 As Currency
Dim v_total2 As Currency
Dim v_lineas As Currency
Dim v_socios As Currency
Dim v_socios1 As Currency
Dim v_socios2 As Currency
Dim v_dombanco As String
Dim v_pobbanco As String
Dim AntCoope As Integer
Dim ActCoope As Integer
Dim Banco As Currency

    On Error GoTo EGen
    GeneraFichero = False

    NFich = FreeFile
    Open App.Path & "\trazabilidad.txt" For Output As #NFich

    Set RS = New ADODB.Recordset
    
    sql = "select cc.codsocio, ss.nomsocio, ss.nifsocio, cc.nrocampo, cc.poligono,cc.parcela, pp.codpobla, pp.nomparti, ppp.despobla, cc.codparti, cc.recintos, cc.codvarie, cc.supcoope, cc.anoplant "
    sql = sql & " from rcampos cc, rsocios ss, rpartida pp, rpueblos ppp    "
    sql = sql & " where cc.codsocio = ss.codsocio and cc.codparti = pp.codparti and pp.codpobla = ppp.codpobla "
    sql = sql & " and cc.fecbajas is null "
    
    'rpueblos", "despobla", "codpobla", CodPobla,
    
    If txtCodigo(171).Text <> "" Then sql = sql & " and cc.codsocio >= " & DBSet(txtCodigo(171).Text, "N")
    If txtCodigo(172).Text <> "" Then sql = sql & " and cc.codsocio <= " & DBSet(txtCodigo(172).Text, "N")
    If txtCodigo(177).Text <> "" Then sql = sql & " and cc.codvarie >= " & DBSet(txtCodigo(177).Text, "N")
    If txtCodigo(178).Text <> "" Then sql = sql & " and cc.codvarie <= " & DBSet(txtCodigo(178).Text, "N")
    
    sql = sql & " order by cc.codsocio "
    
    RS.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    v_lineas = 0
    
    
    While Not RS.EOF
    
        Label2(237).Caption = "Socio: " & Format(DBLet(RS!Codsocio), "000000") & " Nro.Campo: " & Format(DBLet(RS!NroCampo), "000000")
        DoEvents
    
        cad = Format(DBLet(RS!Codsocio), "000000") & ";"
        cad = cad & DBLet(RS!nomsocio) & ";"
        cad = cad & DBLet(RS!nifSocio) & ";"
        cad = cad & Format(DBLet(RS!NroCampo), "000000") & ";"
        cad = cad & Format(DBLet(RS!Poligono), "000") & ";"
        cad = cad & Format(DBLet(RS!Parcela), "000000") & ";"
        cad = cad & DBLet(RS!desPobla) & ";"
        cad = cad & DBLet(RS!nomparti) & ";"
        cad = cad & Format(DBLet(RS!recintos), "000") & ";"
        cad = cad & Format(DBLet(RS!codvarie), "000000") & ";"
        cad = cad & Format(DBLet(RS!supcoope), "###0.0000") & ";"
        cad = cad & Format(DBLet(RS!anoplant), "0000") & ";"
                    
        v_lineas = v_lineas + 1
            
        Print #NFich, cad
        
        RS.MoveNext
    Wend
       
    Close (NFich)
    If v_lineas > 0 Then GeneraFichero = True
    Exit Function
EGen:
    Set RS = Nothing
    Close (NFich)
    MuestraError Err.Number, Err.Description

End Function



Private Sub CmdAcepDifKilos_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String
Dim Tabla2 As String

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
     'D/H Socio
     cDesde = Trim(txtCodigo(186).Text)
     cHasta = Trim(txtCodigo(187).Text)
     nDesde = txtNombre(186).Text
     nHasta = txtNombre(187).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
     End If

     'D/H Clase
     cDesde = Trim(txtCodigo(188).Text)
     cHasta = Trim(txtCodigo(189).Text)
     nDesde = txtNombre(188).Text
     nHasta = txtNombre(189).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codclase}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
     End If
        
     
     
     'D/H fecha
     cDesde = Trim(txtCodigo(184).Text)
     cHasta = Trim(txtCodigo(185).Text)
     nDesde = ""
     nHasta = ""
     
     devuelve = CadenaDesdeHasta(cDesde, cHasta, "fecalbar", "F")
     
     CadParam = CadParam & AnyadirParametroDH("pDHFecha=""", cDesde, cHasta, "", "")
     numParam = numParam + 1

     tabla = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
     tabla = "(" & tabla & ") INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
     
     
     '[Monica]13/11/2013: faltaria el tema de los coopropietarios
     Tabla2 = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) inner join rcampos_cooprop on rcampos.codcampo = rcampos_cooprop.codcampo"
     Tabla2 = "(" & Tabla2 & ") INNER JOIN rsocios ON rcampos_cooprop.codsocio = rsocios.codsocio "
     
     
     
     vSQL = ""
     If txtCodigo(188).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(188).Text, "N")
     If txtCodigo(189).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(189).Text, "N")
     
     Set frmMens = New frmMensajes
     
     frmMens.OpcionMensaje = 16
     frmMens.cadWHERE = vSQL
     frmMens.Show vbModal
     
     Set frmMens = Nothing
            
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If HayRegParaInforme(tabla, cadSelect) Then
        If CargarTemporal7(tabla, cadSelect, Tabla2) Then
           If HayRegParaInforme("tmpinfkilos", "codusu = " & vUsu.Codigo) Then
           
                indRPT = 112 'informe de diferencias de kilos
            
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub '   cadNombreRPT = "rInfKilosSocio.rpt"
                cadTitulo = "Informe de Diferencias de Kilos"
                cadNombreRPT = nomDocu '"rInfDiferencias.rpt"
                
                If Opcion1(15).Value Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "Var.rpt")
                cadFormula = "{tmpinfkilos.codusu} = " & vUsu.Codigo & " and {tmpinfkilos.kilosnet} <> 0 "
                LlamarImprimir
           End If
        End If
     End If
End Sub

Private Sub CmdAcepGene_Click()
Dim sql As String
Dim Ordinal As Long

    On Error GoTo eError

'    CadTag = "codvarie|fechacla|codsocio|codcampo|kilosnet|observac|situacio|"

    If Not DatosOK Then Exit Sub

    sql = "select max(ordinal) from rclasifauto    "
    sql = sql & " where codsocio = " & DBSet(txtCodigo(83).Text, "N")
    sql = sql & " and codcampo = " & DBSet(txtCodigo(80).Text, "N")
    sql = sql & " and codvarie = " & DBSet(RecuperaValor(CadTag, 1), "N")
    
    Ordinal = DevuelveValor(sql) + 1
    
    conn.BeginTrans
    
    
    ' cabecera
    sql = "insert into rclasifauto (numnotac,codsocio,codcampo,codvarie,fechacla,kilosnet,kilospeq,observac,situacion,porcdest,ordinal) values ("
    sql = sql & DBSet(1, "N") & "," ' nro clasificacion
    sql = sql & DBSet(txtCodigo(83).Text, "N") & ","
    sql = sql & DBSet(txtCodigo(80).Text, "N") & ","
    sql = sql & DBSet(RecuperaValor(CadTag, 1), "N") & "," ' variedad
    sql = sql & DBSet(RecuperaValor(CadTag, 2), "F") & "," ' fecha de clasificacion
    sql = sql & DBSet(RecuperaValor(CadTag, 5), "N") & "," ' kilosnet
    sql = sql & DBSet(0, "N") & "," 'kilos manuales
    sql = sql & DBSet(RecuperaValor(CadTag, 6), "T") & "," ' observac
    sql = sql & DBSet(RecuperaValor(CadTag, 7), "N") & "," ' situacion
    sql = sql & DBSet(txtCodigo(79).Text, "N") & "," ' porcentaje de destrio
    sql = sql & DBSet(Ordinal, "N") & ")"
    
    conn.Execute sql
    
    ' lineas
    sql = "insert into rclasifauto_clasif (numnotac,codvarie,codcalid,kiloscal,codcampo,codsocio,fechacla,ordinal) "
    sql = sql & " select 1, codvarie, codcalid, kiloscal," & DBSet(txtCodigo(80).Text, "N") & "," & DBSet(txtCodigo(83).Text, "N") & ", fechacla, " & DBSet(Ordinal, "N")
    sql = sql & " from rclasifauto_clasif "
    sql = sql & " where codvarie = " & DBSet(RecuperaValor(CadTag, 1), "N")  ' variedad
    sql = sql & " and codcampo = 9999"
    sql = sql & " and codsocio = 999"
    sql = sql & " and fechacla = " & DBSet(RecuperaValor(CadTag, 2), "F")  ' fecha de clasificacion
    
    conn.Execute sql
    
eError:
    If Err.Number <> 0 Then
        conn.RollbackTrans
        MuestraError Err.Number, "Generacion de clasificaci�n Autom�tica", Err.Description
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (0)
    End If
End Sub

Private Sub cmdAcepGenEntradas_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String

    If txtCodigo(135).Text = "" Then
        MsgBox "Debe de introducir la fecha de albar�n. Reintroduzca.", vbExclamation
        Exit Sub
    End If
    
    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================

    'D/H Fecha de Factura
    cDesde = Trim(txtCodigo(136).Text)
    cHasta = Trim(txtCodigo(137).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= """) Then Exit Sub
    End If

    ' Tipo de Factura
    If Not AnyadirAFormula(cadSelect, "{rfactsoc.codtipom} = ""SIN""") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "{rfactsoc.codtipom}  = ""SIN""") Then Exit Sub
     
    tabla = "rfactsoc INNER JOIN rsocios ON rfactsoc.codsocio = rsocios.codsocio"
     
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadSelect) Then
        B = GenerarEntradasSIN(tabla, Replace(cadSelect, "rfactsoc", "aaa"))
        If B Then
            MsgBox "Proceso realizado correctamente", vbExclamation
            cmdCancel_Click (16)
        End If
     End If
End Sub

Private Sub CmdAcepGtosCampos_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
     'D/H Socio
     cDesde = Trim(txtCodigo(120).Text)
     cHasta = Trim(txtCodigo(121).Text)
     nDesde = txtNombre(120).Text
     nHasta = txtNombre(121).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
     End If

     'D/H Campo
     cDesde = Trim(txtCodigo(122).Text)
     cHasta = Trim(txtCodigo(123).Text)
     nDesde = ""
     nHasta = ""
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos_gastos.codcampo}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCampo= """) Then Exit Sub
     End If
        
     
    ' D/H Concepto de gastos
     cDesde = Trim(txtCodigo(124).Text)
     cHasta = Trim(txtCodigo(125).Text)
     nDesde = txtNombre(124).Text
     nHasta = txtNombre(125).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos_gastos.codgasto}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHConcepto= """) Then Exit Sub
     End If
        
     
     'D/H fecha
     cDesde = Trim(txtCodigo(126).Text)
     cHasta = Trim(txtCodigo(127).Text)
     nDesde = ""
     nHasta = ""
     'devuelve = CadenaDesdeHasta(cDesde, cHasta, "fecalbar", "F")
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos_gastos.fecha}"
         TipCod = "F"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= """) Then Exit Sub
     End If


     tabla = "(rcampos_gastos INNER JOIN rcampos ON rcampos_gastos.codcampo = rcampos.codcampo) "
     tabla = "(" & tabla & ") INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
     tabla = "(" & tabla & ") INNER JOIN rconcepgasto ON rcampos_gastos.codgasto = rconcepgasto.codgasto "
     
     ' agrupado por socio
     If Opcion1(5).Value Then
        CadParam = CadParam & "pGroup1={rcampos.codsocio}" & "|"
        CadParam = CadParam & "pGroup1Name= ""SOCIO: "" & " & " totext({rcampos.codsocio},""000000"") & " & """  """ & " & {rsocios.nomsocio}" & "|"
        CadParam = CadParam & "pGroup2={rcampos_gastos.codgasto}" & "|"
        CadParam = CadParam & "pGroup2Name= totext({rcampos_gastos.codgasto},""00"") & " & """  """ & " & {rconcepgasto.nomgasto}" & "|"
        
        CadParam = CadParam & "pTitulo1=""Gasto""|"
        
        numParam = numParam + 5
     End If
    
     'agrupado por concepto de gasto
     If Opcion1(6).Value Then
        CadParam = CadParam & "pGroup1={rcampos_gastos.codgasto}" & "|"
        CadParam = CadParam & "pGroup1Name= ""CONCEPTO DE GASTO: "" & " & " totext({rcampos_gastos.codgasto},""00"") & " & """  """ & " & {rconcepgasto.nomgasto}" & "|"
        CadParam = CadParam & "pGroup2={rcampos.codsocio}" & "|"
        CadParam = CadParam & "pGroup2Name= totext({rcampos.codsocio},""000000"") & " & """  """ & " & {rsocios.nomsocio}" & "|"
        
        CadParam = CadParam & "pTitulo1=""Socio""|"
        
        numParam = numParam + 5
     End If
     
     ' si hay resumen lo marcamos
     CadParam = CadParam & "pResumen=" & Check17.Value & "|"
     numParam = numParam + 1
     
     
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If HayRegParaInforme(tabla, cadSelect) Then
'         indRPT = 69 'informe de gastos por concepto
'
'         If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub '   cadNombreRPT = "rInfKilosSocio.rpt"
         If Opcion1(5).Value Then
            cadTitulo = "Informe de Gastos por Socio"
         Else
            cadTitulo = "Informe de Gastos por Concepto"
         End If
         cadNombreRPT = "rInfGtosCampos.rpt"
                
         ConSubInforme = False
         
         LlamarImprimir
     End If

End Sub

Private Sub CmdAcepGtosConcep_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
     'D/H Socio
     cDesde = Trim(txtCodigo(100).Text)
     cHasta = Trim(txtCodigo(101).Text)
     nDesde = txtNombre(100).Text
     nHasta = txtNombre(101).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rhisfruta.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
     End If

     'D/H Clase
     cDesde = Trim(txtCodigo(104).Text)
     cHasta = Trim(txtCodigo(105).Text)
     nDesde = txtNombre(104).Text
     nHasta = txtNombre(105).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codclase}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
     End If
        
     
    ' VARIEDAD
     cDesde = Trim(txtCodigo(102).Text)
     cHasta = Trim(txtCodigo(103).Text)
     nDesde = txtNombre(102).Text
     nHasta = txtNombre(103).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rhisfruta.codvarie}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
     End If
        
     
     'D/H fecha
     cDesde = Trim(txtCodigo(98).Text)
     cHasta = Trim(txtCodigo(99).Text)
     nDesde = ""
     nHasta = ""
     'devuelve = CadenaDesdeHasta(cDesde, cHasta, "fecalbar", "F")
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rhisfruta.fecalbar}"
         TipCod = "F"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= """) Then Exit Sub
     End If


    ' CONCEPTO DE GASTOS
     cDesde = Trim(txtCodigo(96).Text)
     cHasta = Trim(txtCodigo(97).Text)
     nDesde = txtNombre(96).Text
     nHasta = txtNombre(97).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rhisfruta_gastos.codgasto}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHConcepto= """) Then Exit Sub
     End If
    

     tabla = "(rhisfruta INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
     tabla = "(" & tabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
     tabla = "(" & tabla & ") INNER JOIN rhisfruta_gastos ON rhisfruta.numalbar = rhisfruta_gastos.numalbar "
     
     vSQL = ""
     If txtCodigo(104).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(104).Text, "N")
     If txtCodigo(105).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(105).Text, "N")
     
     Set frmMens = New frmMensajes
     
     frmMens.OpcionMensaje = 16
     frmMens.cadWHERE = vSQL
     frmMens.Show vbModal
     
     Set frmMens = Nothing
            
     ' salto de pagina o no por producto
     CadParam = CadParam & "pSalto=" & Check11.Value & "|"
     numParam = numParam + 1
     
     ' agrupado por socio
     If Opcion1(8).Value Then
        CadParam = CadParam & "pGroup1={rhisfruta.codsocio}" & "|"
        CadParam = CadParam & "pGroup1Name= ""SOCIO: "" & " & " totext({rhisfruta.codsocio},""000000"") & " & """  """ & " & {rsocios.nomsocio}" & "|"
        
        CadParam = CadParam & "pTitulo1=""Variedad""|"
        
        numParam = numParam + 3
     End If
    
     'agrupado por variedad
     If Opcion1(9).Value Then
        CadParam = CadParam & "pGroup1={rhisfruta.codvarie}" & "|"
        CadParam = CadParam & "pGroup1Name= ""VARIEDAD: "" & " & " totext({rhisfruta.codvarie},""000000"") & " & """  """ & " & {variedades.nomvarie}" & "|"
        CadParam = CadParam & "pTitulo1=""Socio""|"
        
        numParam = numParam + 3
     End If
     
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If HayRegParaInforme(tabla, cadSelect) Then
         indRPT = 69 'informe de gastos por concepto
     
         If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub '   cadNombreRPT = "rInfKilosSocio.rpt"
         cadTitulo = "Informe de Gastos por Concepto"
                            
         cadNombreRPT = nomDocu '"rInfGtosConcep.rpt"
                
         ConSubInforme = False
         
         LlamarImprimir
     End If
End Sub


Private Sub CmdAcepInf_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String
Dim Tabla2 As String

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
     'D/H Socio
     cDesde = Trim(txtCodigo(33).Text)
     cHasta = Trim(txtCodigo(34).Text)
     nDesde = txtNombre(33).Text
     nHasta = txtNombre(34).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
     End If

     'D/H Clase
     cDesde = Trim(txtCodigo(35).Text)
     cHasta = Trim(txtCodigo(36).Text)
     nDesde = txtNombre(35).Text
     nHasta = txtNombre(36).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codclase}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
     End If
        
     
    ' PRODUCTO
     cDesde = Trim(txtCodigo(37).Text)
     cHasta = Trim(txtCodigo(38).Text)
     nDesde = txtNombre(37).Text
     nHasta = txtNombre(38).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codprodu}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
     End If
        
     
     'D/H fecha
     cDesde = Trim(txtCodigo(39).Text)
     cHasta = Trim(txtCodigo(40).Text)
     nDesde = ""
     nHasta = ""
     devuelve = CadenaDesdeHasta(cDesde, cHasta, "fecalbar", "F")

     tabla = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
     tabla = "(" & tabla & ") INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
     
     
     '[Monica]13/11/2013: faltaria el tema de los coopropietarios
     Tabla2 = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) inner join rcampos_cooprop on rcampos.codcampo = rcampos_cooprop.codcampo"
     Tabla2 = "(" & Tabla2 & ") INNER JOIN rsocios ON rcampos_cooprop.codsocio = rsocios.codsocio "
     
     
     
     vSQL = ""
     If txtCodigo(35).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(35).Text, "N")
     If txtCodigo(36).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(36).Text, "N")
     
     If txtCodigo(37).Text <> "" Then vSQL = vSQL & " and variedades.codprodu >= " & DBSet(txtCodigo(37).Text, "N")
     If txtCodigo(38).Text <> "" Then vSQL = vSQL & " and variedades.codprodu <= " & DBSet(txtCodigo(38).Text, "N")
     
     Set frmMens = New frmMensajes
     
     frmMens.OpcionMensaje = 16
     frmMens.cadWHERE = vSQL
     frmMens.Show vbModal
     
     Set frmMens = Nothing
            
     'combo1(5): tipo de has
     CadParam = CadParam & "pTipoHas=" & Combo1(5).ListIndex & "|"
     numParam = numParam + 1
     
     ' salto de pagina o no por producto
     CadParam = CadParam & "pSaltoProd=" & Check3.Value & "|"
     numParam = numParam + 1
     
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If HayRegParaInforme(tabla, cadSelect) Then
        If CargarTemporal5(tabla, cadSelect, Tabla2) Then
           If HayRegParaInforme("tmpinfkilos", "codusu = " & vUsu.Codigo) Then
               
                '[Monica]25/01/2017: Personalizacion del informe de kilos por producto
                indRPT = 113 'Informe de kilos por producto
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                  
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                ConSubInforme = True
               
               
               '[Monica]20/07/2015: nuevo informe para Mogente
               If Check22.Value = 1 Then
                    cadNombreRPT = Replace(cadNombreRPT, "Prod.rpt", "ProdDet.rpt")
                    cadTitulo = "Informe de Kilos por Producto"
                    cadFormula = "{tmpinfkilos.codusu} = " & vUsu.Codigo & " and {tmpinfkilos.kilosnet} <> 0 "
               Else
                
'                    cadNombreRPT = "rInfKilosProd.rpt"
                    cadTitulo = "Informe de Kilos por Producto"
                    cadFormula = "{tmpinfkilos.codusu} = " & vUsu.Codigo
               End If
               
               
               LlamarImprimir
           End If
        End If
     End If
End Sub

Private Sub CmdAcepInfATRIA_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String
Dim I As Integer

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
     
     '======== FORMULA  ====================================
     'D/H Socio
     cDesde = Trim(txtCodigo(153).Text)
     cHasta = Trim(txtCodigo(154).Text)
     nDesde = txtNombre(153).Text
     nHasta = txtNombre(154).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
     End If
     
     'D/H Producto
     cDesde = Trim(txtCodigo(148).Text)
     cHasta = Trim(txtCodigo(150).Text)
     nDesde = txtNombre(148).Text
     nHasta = txtNombre(150).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codprodu}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
     End If
     
     'D/H Variedad
     cDesde = Trim(txtCodigo(151).Text)
     cHasta = Trim(txtCodigo(152).Text)
     nDesde = txtNombre(151).Text
     nHasta = txtNombre(152).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos.codvarie}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
     End If
     
     
    ' campo no debe de estar dado de baja
    If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub
     
    ' el socio no debe de estar dado de baja
    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "isnull({rsocios.fechabaja})") Then Exit Sub

    
    tabla = "rcampos inner join rsocios on rsocios.codsocio = rcampos.codsocio"
    tabla = "(" & tabla & ") inner join variedades on variedades.codvarie = rcampos.codvarie"
    tabla = "(" & tabla & ") inner join productos on productos.codprodu = variedades.codprodu"
    tabla = "(" & tabla & ") inner join grupopro on productos.codgrupo = grupopro.codgrupo"
    
    
            
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadSelect) Then
        If CargarTemporalAtria(tabla, cadSelect) Then
            cadNombreRPT = "rInfATRIA.rpt"
            cadTitulo = "Informe de Miembros ATRIA"
            
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            LlamarImprimir
        End If
    End If

End Sub

Private Function CargarTemporalAtria(nTabla As String, nSelect As String) As Boolean
Dim sql As String
Dim Sql2 As String
Dim Rs2 As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim Has As Currency
Dim Nregs As Integer

    On Error GoTo eCargarTemporalAtria
    
    CargarTemporalAtria = False
    
    sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute sql
    
    
    sql = "insert into tmpinformes (codusu, codigo1, nombre1, precio1) select distinct " & vUsu.Codigo & ", rcampos.codsocio, grupopro.codatria, 0 from " & nTabla & " where " & nSelect
    conn.Execute sql
    
    sql = "select rcampos.*, codatria from " & nTabla & " where " & nSelect
    
    Pb8.Max = TotalRegistrosConsulta(sql)
    Pb8.visible = True
    Pb8.Value = 0
    
        
    
    
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        IncrementarProgresNew Pb8, 1
        DoEvents

        If TieneCopropietarios(RS!codcampo, RS!Codsocio) Then
            Sql2 = "select * from rcampos_cooprop where codcampo = " & DBSet(RS!codcampo, "N")
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                Has = Round2(RS!supcoope * DBLet(RS!Porcentaje, "N") / 100, 4)
                
                sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo & " and codigo1 = " & DBSet(Rs2!Codsocio, "N") & " and nombre1 = " & DBSet(RS!codatria, "T")
                If TotalRegistros(sql) = 0 Then
                    sql = "insert into tmpinformes (codusu, codigo1, nombre1, precio1) values (" & vUsu.Codigo & "," & DBSet(Rs2!Codsocio, "N") & "," & DBSet(RS!codatria, "T") & "," & DBSet(Has, "N") & ")"
                Else
                    sql = "update tmpinformes set precio1 = precio1 + " & DBSet(Has, "N") & " where codusu = " & vUsu.Codigo & " and codigo1 = " & DBSet(Rs2!Codsocio, "N") & " and nombre1 = " & DBSet(RS!codatria, "T")
                End If
                conn.Execute sql
                
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
        Else
            sql = "update tmpinformes set precio1 = precio1 + " & DBSet(RS!supcoope, "N") & " where codusu = " & vUsu.Codigo & " and codigo1 = " & DBSet(RS!Codsocio, "N") & " and nombre1 = " & DBSet(RS!codatria, "T")
            conn.Execute sql
        End If
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    Pb8.visible = False
    CargarTemporalAtria = True
    Exit Function
    
eCargarTemporalAtria:
    Pb8.visible = False
    MuestraError Err.Number, "Cargar Temporal Atria", Err.Description
End Function

Private Function TieneCopropietarios(campo As String, Propietario As String) As Boolean
Dim NroCampo As String
Dim sql As String

    sql = "select count(*) from rcampos_cooprop where codcampo = " & DBSet(campo, "N") & " and codsocio <> " & DBSet(Propietario, "N")
    
    TieneCopropietarios = TotalRegistros(sql) > 0

End Function




Private Sub CmdAcepInfFases_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String
Dim I As Integer

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    CadParam = CadParam & "pFase=" & Combo1(12).ListIndex & "|"
    numParam = numParam + 1
    
    If Combo1(12).ListIndex <> 0 Then
        If Not AnyadirAFormula(cadFormula, "{rsocios_pozos.numfases} = " & Combo1(12).Text) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{rsocios_pozos.numfases} = " & Combo1(12).Text) Then Exit Sub
    End If
    
    tabla = "rsocios inner join rsocios_pozos on rsocios.codsocio = rsocios_pozos.codsocio"
            
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadSelect) Then
        cadNombreRPT = "rManSociosporFases.rpt"
        cadTitulo = "Informe de Socios por Fases"
        
        LlamarImprimir
    End If


End Sub

Private Sub cmdAcepInfFito_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vcad As String

    InicializarVbles
    
    ConSubInforme = False
    
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtCodigo(173).Text)
    cHasta = Trim(txtCodigo(174).Text)
    nDesde = txtNombre(173).Text
    nHasta = txtNombre(174).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcampos.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
       
    'D/H producto
    cDesde = Trim(txtCodigo(175).Text)
    cHasta = Trim(txtCodigo(176).Text)
    nDesde = txtNombre(175).Text
    nHasta = txtNombre(176).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codprodu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
    End If
    
    'D/H Partida
    cDesde = Trim(txtCodigo(167).Text)
    cHasta = Trim(txtCodigo(168).Text)
    nDesde = txtNombre(167).Text
    nHasta = txtNombre(168).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcampos.codparti}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPartida= """) Then Exit Sub
    End If
            
    'D/H Poblacion ( Termino Municipal )
    cDesde = Trim(txtCodigo(159).Text)
    cHasta = Trim(txtCodigo(160).Text)
    nDesde = txtNombre(159).Text
    nHasta = txtNombre(160).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpueblos.codpobla}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPueblos= """) Then Exit Sub
    End If
    
    If txtCodigo(180).Text <> "" Then
        CadParam = CadParam & "pCampanya=""" & txtCodigo(180).Text & """|"
        numParam = numParam + 1
    End If
    
    tabla = "((rcampos INNER JOIN rpartida ON rcampos.codparti = rpartida.codparti) "
    tabla = tabla & " INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
    tabla = tabla & " INNER JOIN rpueblos ON rpartida.codpobla = rpueblos.codpobla "
            
            
    If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub
            
            
     'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadSelect) Then
        
        cadTitulo = "Informe de Registro Aplicaci�n de Fitosanitarios"
        
        cadNombreRPT = "rInfRegFitosanitarios.rpt"
        indRPT = 106 'Informe de registro de Aplicacion de fitosanitarios
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
          
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
        ConSubInforme = True

'        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo

        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Titulo = cadTitulo
            .NombreRPT = cadNombreRPT
            .ConSubInforme = True
            .Opcion = 0
            .Show vbModal
        End With
    End If
End Sub

Private Sub cmdAcepInfSocios_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vcad As String

    InicializarVbles
    
    ConSubInforme = False
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtCodigo(145).Text)
    cHasta = Trim(txtCodigo(146).Text)
    nDesde = txtNombre(145).Text
    nHasta = txtNombre(146).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rsocios.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
'[Monica] 16/09/2009 incluir los socios dados de baja
    If Check20.Value = 0 Then
        vcad = "isnull({rsocios.fechabaja})"
        If AnyadirAFormula(cadFormula, vcad) = False Then Exit Sub
        vcad = "rsocios.fechabaja is null"
        If AnyadirAFormula(cadSelect, vcad) = False Then Exit Sub
    End If
    
    '[Monica]19/01/2012: insertamos las situaciones de socios que vamos a incluir
    
    Set frmMens4 = New frmMensajes
    
    frmMens4.OpcionMensaje = 36
    frmMens4.Show vbModal
    
    Set frmMens4 = Nothing
    
    
    'Obtener el parametro con el ORDEN del Informe
    '---------------------------------------------
'    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
'    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente

    tabla = "rsocios"

    If Opcion(7).Value Then
        CadParam = CadParam & "pTitulo1=""Listado de Tel�fonos de Socios""|"
        numParam = numParam + 1
    End If

    cadNombreRPT = "rManSocSeccion.rpt"
    
    '[Monica]18/05/2012: personalizacion de informe de socios/seccion
    indRPT = 99
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    
    cadTitulo = "Listado de Datos de Socios"
    
    ' por codigo
    If Opcion(8).Value Then
        CadParam = CadParam & "pOrden={rsocios.codsocio}|"
    End If
    ' alfabetico
    If Opcion(9).Value Then
        CadParam = CadParam & "pOrden={rsocios.nomsocio}|"
    End If
    numParam = numParam + 1
        
        
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadSelect) Then
        LlamarImprimir
    End If

End Sub

Private Sub CmdAcepKilosSoc_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String
Dim I As Integer
Dim J As Integer
Dim cadena As String


    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
     'D/H Socio
     cDesde = Trim(txtCodigo(54).Text)
     cHasta = Trim(txtCodigo(55).Text)
     nDesde = txtNombre(54).Text
     nHasta = txtNombre(55).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rhisfruta.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
     End If

     'D/H Clase
     cDesde = Trim(txtCodigo(56).Text)
     cHasta = Trim(txtCodigo(57).Text)
     nDesde = txtNombre(56).Text
     nHasta = txtNombre(57).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codclase}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
     End If
        
    vSQL = ""
    If txtCodigo(56).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(56).Text, "N")
    If txtCodigo(57).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(57).Text, "N")
     
    ' PRODUCTO
     cDesde = Trim(txtCodigo(50).Text)
     cHasta = Trim(txtCodigo(51).Text)
     nDesde = txtNombre(50).Text
     nHasta = txtNombre(51).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codprodu}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
     End If
        
    If txtCodigo(50).Text <> "" Then vSQL = vSQL & " and variedades.codprodu >= " & DBSet(txtCodigo(50).Text, "N")
    If txtCodigo(51).Text <> "" Then vSQL = vSQL & " and variedades.codprodu <= " & DBSet(txtCodigo(51).Text, "N")
     
     'D/H fecha
     cDesde = Trim(txtCodigo(52).Text)
     cHasta = Trim(txtCodigo(53).Text)
     nDesde = ""
     nHasta = ""
     devuelve = CadenaDesdeHasta(cDesde, cHasta, "fecalbar", "F")

     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rhisfruta.fecalbar}"
         TipCod = "F"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
     End If
     
     cadSelect = ""
     
     Set frmMens = New frmMensajes
     
     frmMens.OpcionMensaje = 16
     frmMens.cadWHERE = vSQL
     frmMens.Show vbModal
     
     Set frmMens = Nothing
            
     tabla = "rcampos"
            
     '[Monica]29/05/2012: Cargamos todos los tipos de entrada de tipos de entrada en el parametro
     cadena = ""
     J = 0
     For I = 1 To ListView1(0).ListItems.Count
        If ListView1(0).ListItems(I).Checked Then
            J = J + 1
            cadena = cadena & ListView1(0).ListItems(I).Text & ", "
        End If
     Next I
     If J = ListView1(0).ListItems.Count Then
        CadParam = CadParam & "pTipos=""Todas""|"
     Else
        CadParam = CadParam & "pTipos=""" & Mid(cadena, 1, Len(cadena) - 2) & """|"
     End If
     numParam = numParam + 1
            
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If CargarTemporal6(tabla, cadSelect) Then
         If HayRegParaInforme("tmpclasifica", "codusu = " & vUsu.Codigo) Then
             indRPT = 62 'informe de Kilos Recolectados Socio/Cooperativa
     
             If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub '   cadNombreRPT = "rInfKilosSocio.rpt"
             cadTitulo = "Informe de Kilos Socio/Cooperativa"
                            
             cadFormula = "{tmpclasifica.codusu} = " & vUsu.Codigo
    
             cadNombreRPT = nomDocu
                
             ConSubInforme = False
             
             LlamarImprimir
         End If
     End If

End Sub

Private Sub CmdAcepOrdEmitidas_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vcad As String
Dim sql As String

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H Fecha de Impresion Orden de Recoleccion
    cDesde = Trim(txtCodigo(139).Text)
    cHasta = Trim(txtCodigo(140).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcampos_ordrec.fecimpre}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= """) Then Exit Sub
    End If
    
    'D/H VARIEDAD
    cDesde = Trim(txtCodigo(143).Text)
    cHasta = Trim(txtCodigo(144).Text)
    nDesde = txtNombre(143).Text
    nHasta = txtNombre(144).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcampos.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
    End If
    
    tabla = "rcampos INNER JOIN rcampos_ordrec ON rcampos.codcampo = rcampos_ordrec.codcampo "
        
    If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub
        
    cadNombreRPT = "rInfOrdenRecol.rpt"
    
    indRPT = 97 ' Informe de Ordenes de recoleccion emitidas
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    ConSubInforme = False
    cadTitulo = "Informe de Ordenes de Recolecci�n"
   
     'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadSelect) Then
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Titulo = cadTitulo
            .NombreRPT = cadNombreRPT
            .ConSubInforme = True
            .Opcion = 0
            .NroCopias = 1
            .Show vbModal
        End With
    End If

End Sub

Private Sub cmdAcepOrdenRec_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vcad As String
Dim sql As String

    InicializarVbles
    
    If Not EsReimpresion Then
        'Bloqueo el proceso pq puede tengo q coger el contador ORR
        sql = "IMPORD" 'IMPresion ORDenes recoleccion
        
        'Bloquear para que nadie mas pueda contabilizar
        DesBloqueoManual (sql)
        If Not BloqueoManual(sql, "1") Then
            MsgBox "No se pueden Imprimir Ordenes de Recolecci�n. Hay otro usuario realiz�ndolo.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    
    If Not DatosOK Then Exit Sub
    
    sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute sql
 
    
    vSQL = ""
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H Capataz(Responsable)
    cDesde = Trim(txtCodigo(147).Text)
    cHasta = Trim(txtCodigo(147).Text)
    nDesde = txtNombre(147).Text
    nHasta = txtNombre(147).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcampos.codcapat}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHResponsable= """) Then Exit Sub
    End If
    
    vSQL = vSQL & " and rcampos.codcapat = " & DBSet(txtCodigo(147).Text, "N")
    
    'D/H VARIEDAD
    cDesde = Trim(txtCodigo(149).Text)
    cHasta = Trim(txtCodigo(149).Text)
    nDesde = txtNombre(149).Text
    nHasta = txtNombre(149).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcampos.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
    End If
    
    vSQL = vSQL & " and rcampos.codvarie = " & DBSet(txtCodigo(149).Text, "N")

    'D/H Partida
    cDesde = Trim(txtCodigo(142).Text)
    cHasta = Trim(txtCodigo(142).Text)
    nDesde = txtNombre(142).Text
    nHasta = txtNombre(142).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcampos.codparti}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPartida= """) Then Exit Sub
    End If
    
    '[Monica]30/09/2013: dejo que la partida sea opcional
    If txtCodigo(142).Text <> "" Then
        vSQL = vSQL & " and rcampos.codparti = " & DBSet(txtCodigo(142).Text, "N")
    End If

    tabla = "rcampos"
        
    If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub
        
        
    If Not EsReimpresion Then
        
        ' el campo no debe de estar marcado como finalizado de recolectar
        If Not AnyadirAFormula(cadSelect, "{rcampos.acabadorecol} = 0") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rcampos.acabadorecol} = 0") Then Exit Sub
    
    End If
    
    vSQL = vSQL & " and rcampos.acabadorecol = 0 "
    
    CadParam = CadParam & "pFecha=""" & txtCodigo(138).Text & """|"
    numParam = numParam + 1

    vSQL = vSQL & " and rcampos.fecbajas is null "
    
    cadNombreRPT = "rOrdenRecol.rpt"
    
    indRPT = 96 ' Ordenes de recoleccion
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    ConSubInforme = False
    
    If Not EsReimpresion Then
    
        '[Monica]11/11/2013: indicamos si han entrado o no por campos
        HayRegistros = False
    
        Set frmMens5 = New frmMensajes
        
        frmMens5.OpcionMensaje = 51
        frmMens5.cadWHERE = vSQL
        frmMens5.Show vbModal
        
        Set frmMens5 = Nothing
        
        If Not HayRegistros Then
            MsgBox "No hay datos para mostrar en el Informe.", vbExclamation
            
            '[Monica]10/11/2016: a�ado aqu� el desbloqueo para ver si as� no se queda bloqueado
            'Desbloqueamos ya no estamos imprimiendo ordenes de recoleccion
            DesBloqueoManual ("IMPORD") 'IMPresion ORDenes de recoleccion
            
            Exit Sub
        End If
        
            
         'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(tabla, cadSelect) Then
            cadTitulo = "Orden de Recolecci�n"
            
            If InsertarTemporal(tabla, cadSelect) Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                
                With frmImprimir
                    .FormulaSeleccion = cadFormula
                    .OtrosParametros = CadParam
                    .NumeroParametros = numParam
                    .SoloImprimir = False
                    .EnvioEMail = False
                    .Titulo = cadTitulo
                    .NombreRPT = cadNombreRPT
                    .ConSubInforme = True
                    .Opcion = 0
                    '[Monica]13/02/2017: ahora en el caso de alzira quieren 2 copias
                    If vParamAplic.Cooperativa = 4 Then
                        .NroCopias = 2
                    Else
                        '[Monica]11/09/2013: ahora el nro de copias es 1
                        .NroCopias = 1
                    End If
                    .Show vbModal
                    
                End With
                
                If MsgBox("� Impresi�n correcta para actualizar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                Else
                    If ActualizarDatos Then
                    End If
                End If
            
            End If
        End If
            
        Set vTipoMov = Nothing
            
        'Desbloqueamos ya no estamos imprimiendo ordenes de recoleccion
        DesBloqueoManual ("IMPORD") 'IMPresion ORDenes de recoleccion
    
    Else
        cadSelect = " codcampo in (select codcampo from rcampos where nrocampo in (select nrocampo from rordrecogida where nroorden = " & DBSet(txtCodigo(141).Text, "N") & ")) "

        If InsertarTemporal2(tabla, cadSelect) Then
            cadTitulo = "Reimpresi�n Orden de Recolecci�n"
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            With frmImprimir
                .FormulaSeleccion = cadFormula
                .OtrosParametros = CadParam
                .NumeroParametros = numParam
                .SoloImprimir = False
                .EnvioEMail = False
                .Titulo = cadTitulo
                .NombreRPT = cadNombreRPT
                .ConSubInforme = True
                .Opcion = 0
                '[Monica]11/09/2013: ahora el nro de copias es 1
                .NroCopias = 1
                .Show vbModal
            End With
        End If
    End If
    
End Sub

Private Function ActualizarDatos() As Boolean
Dim sql As String
Dim Sql2 As String
Dim Sql3 As String

Dim Sql4 As String
Dim CadValues As String
Dim RS As ADODB.Recordset

Dim cadCampos As String
Dim Rs2 As ADODB.Recordset

    On Error GoTo eActualizarDatos

    conn.BeginTrans

    ActualizarDatos = False
    
    CadValues = ""
    
    'Sql2 = "insert into rcampos_ordrec (codcampo, nroorden, fecimpre) values "
    Sql2 = "delete from rordrecogida where nroorden = "
    Sql3 = "delete from rordrecogida_incid where nroorden = "
    
    sql = "select * from tmpinformes where codusu = " & vUsu.Codigo & " order by importe3 desc"
    
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
    
    '    CadValues = CadValues & "(" & DBSet(RS!importe2, "N") & "," & DBSet(RS!importe3, "N") & "," & DBSet(txtCodigo(138).Text, "F") & "),"
        conn.Execute Sql3 & DBSet(RS!importe3, "N")
        
        conn.Execute Sql2 & DBSet(RS!importe3, "N")
        
        Sql4 = "update usuarios.stipom set contador = " & DBSet(RS!importe3 - 1, "N") & " where codtipom = " & DBSet(CodTipoMov, "T")
        conn.Execute Sql4
        
        RS.MoveNext
    Wend
    Set RS = Nothing
        
    conn.CommitTrans
    
    ActualizarDatos = True
    Exit Function
    
eActualizarDatos:
    MuestraError Err.Number, "Actualizar Datos", Err.Description
    conn.RollbackTrans
End Function

Private Function InsertarTemporal(cTabla As String, cSelect As String) As Boolean
Dim sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim NumF As Long
Dim devuelve As String
Dim Existe As Boolean

    On Error GoTo eInsertarTemporal

    
    
    InsertarTemporal = False


    conn.BeginTrans


    sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute sql
    
    sql = "insert into tmpinformes (codusu, importe1, importe2) "
    sql = sql & " select " & vUsu.Codigo & ",nrocampo, codcampo from rcampos where " & cSelect  ' copia 1
'    SQL = SQL & " group by 1,2,3 "
'    SQL = SQL & " union "
'    SQL = SQL & " select " & vUsu.Codigo & ",nrocampo, codcampo from rcampos where " & cSelect ' copia 2
'    SQL = SQL & " group by 1,2,3 "
    sql = sql & " order by 1,2,3 "
    
    conn.Execute sql
    
    CodTipoMov = "ORR"
    Set vTipoMov = New CTiposMov
    
    If vTipoMov.Leer(CodTipoMov) Then
        'contador de la orden de recoleccion
        
        
        PriFact = 0
        
        sql = "select distinct importe1 from tmpinformes where codusu = " & vUsu.Codigo & " order by 1 "
        Set RS = New ADODB.Recordset
        RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF
    
            NumF = vTipoMov.ConseguirContador(CodTipoMov)
            
            '[Monica]11/11/2013: a�adido esto por si existe el nro de orden de recogida
            Do
                NumF = vTipoMov.ConseguirContador(CodTipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rordrecogida", "nroorden", "nroorden", CStr(NumF), "N")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (CodTipoMov)
                    NumF = vTipoMov.ConseguirContador(CodTipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            If PriFact = 0 Then PriFact = NumF
          ' hasta aqui
        
        
            Sql2 = "update tmpinformes set importe3 = " & DBSet(NumF, "N")
            Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
            Sql2 = Sql2 & "  and importe1 = " & DBSet(RS!importe1, "N")
            
            conn.Execute Sql2
            
            Sql2 = "insert into rordrecogida (nroorden, fecimpre, nrocampo, codvarie) values ("
            Sql2 = Sql2 & DBSet(NumF, "N") & "," & DBSet(txtCodigo(138).Text, "F") & ","
            Sql2 = Sql2 & DBSet(RS!importe1, "N") & "," & DBSet(txtCodigo(149).Text, "N") & ")"
            
            conn.Execute Sql2
            
            ' lineas de incidencias
            '[Monica]26/11/2013: faltaba el distinct del select
            Sql2 = "insert into rordrecogida_incid (nroorden, idplaga, nivel) "
            Sql2 = Sql2 & " select distinct " & DBSet(NumF, "N") & ", idplaga, nivel from rordrecogida_incid, rordrecogida aaa "
            Sql2 = Sql2 & " where rordrecogida_incid.nroorden = aaa.nroorden and "
            Sql2 = Sql2 & " aaa.nrocampo = " & DBSet(RS!importe1, "N") & " and "
            Sql2 = Sql2 & " aaa.codvarie = " & DBSet(txtCodigo(149).Text, "N") & " and "
            Sql2 = Sql2 & " aaa.fecimpre in (select  max(fecimpre) from rordrecogida bbb where bbb.nrocampo = " & DBSet(RS!importe1, "N") & " and bbb.codvarie = " & DBSet(txtCodigo(149).Text, "N") & " and bbb.nroorden <> " & DBSet(NumF, "N") & " ) "
            
            conn.Execute Sql2


            vTipoMov.IncrementarContador (CodTipoMov)

            RS.MoveNext
        Wend
        Set RS = Nothing
    Else
        InsertarTemporal = False
        Exit Function
    End If
    
    InsertarTemporal = True
    conn.CommitTrans
    Exit Function
    
eInsertarTemporal:
    MuestraError Err.Number, "Insertar en Temporal", Err.Description
    conn.RollbackTrans
End Function


Private Function InsertarTemporal2(cTabla As String, cSelect As String) As Boolean
Dim sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim NumF As Long

    On Error GoTo eInsertarTemporal

    
    
    InsertarTemporal2 = False


    sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute sql
    
    sql = "insert into tmpinformes (codusu, importe1, importe2) "
    sql = sql & " select " & vUsu.Codigo & ",nrocampo, codcampo from rcampos where " & cSelect  ' copia 1
    sql = sql & " order by 1,2,3 "
    
    conn.Execute sql
    
    Sql2 = "update tmpinformes set importe3 = " & DBSet(txtCodigo(141).Text, "N")
    Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
    
    conn.Execute Sql2
    
    
    
    InsertarTemporal2 = True
    
    Exit Function
eInsertarTemporal:
    MuestraError Err.Number, "Insertar en Temporal", Err.Description
End Function





Private Sub cmdAcepPrecios_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String


    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


    'D/H Variedad
    cDesde = Trim(txtCodigo(155).Text)
    cHasta = Trim(txtCodigo(156).Text)
    nDesde = txtNombre(155).Text
    nHasta = txtNombre(156).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rprecios.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
    End If

    'D/H Fecha de precios
    cDesde = Trim(txtCodigo(157).Text)
    cHasta = Trim(txtCodigo(158).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rprecios.fechaini}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= """) Then Exit Sub
    End If

    If Not AnyadirAFormula(cadFormula, "{productos.codgrupo} <> 5 and {productos.codgrupo} <> 6") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "productos.codgrupo <> 5 and productos.codgrupo <> 6") Then Exit Sub
    If Combo1(13).ListIndex <> -1 Then
        If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = " & Combo1(13).ListIndex) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "rprecios.tipofact = " & Combo1(13).ListIndex) Then Exit Sub
    End If
    
    tabla = "(rprecios inner join variedades on rprecios.codvarie = variedades.codvarie) inner join productos on variedades.codprodu = productos.codprodu "
    
    If HayRegParaInforme(tabla, cadSelect) Then
        cadNombreRPT = "rManPrecios.rpt"
        cadTitulo = "Listado de Precios"
        
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Titulo = cadTitulo
            .NombreRPT = cadNombreRPT
            .ConSubInforme = True
            .Opcion = 0
            .Show vbModal
        End With
        
    End If

End Sub

Private Sub cmdAcepRevisionCampos_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String


    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    'D/H Socio
    cDesde = Trim(txtCodigo(163).Text)
    cHasta = Trim(txtCodigo(164).Text)
    nDesde = txtNombre(163).Text
    nHasta = txtNombre(164).Text
    If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If

    'D/H VARIEDAD
    cDesde = Trim(txtCodigo(161).Text)
    cHasta = Trim(txtCodigo(162).Text)
    nDesde = txtNombre(161).Text
    nHasta = txtNombre(162).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcampos.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
    End If

    'D/H Fecha
    cDesde = Trim(txtCodigo(165).Text)
    cHasta = Trim(txtCodigo(166).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        Codigo = "{rcampos_revision.fecha}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
    tabla = "rcampos inner join rcampos_revision on rcampos.codcampo = rcampos_revision.codcampo"
    
    cadNombreRPT = "rRevisionCampos.rpt"
    
    cadTitulo = "Registro Diario de Visitas a Parcelas"
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadSelect) Then
        LlamarImprimir
    End If


End Sub

Private Sub cmdAcepROPAS_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String


    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


    If txtCodigo(62).Text = "" Then
        MsgBox "Debe introducir un valor en el campo ejercicio. Revise.", vbExclamation
        PonerFoco txtCodigo(62)
        Exit Sub
    End If

    If txtCodigo(132).Text = "" Then
        MsgBox "Debe introducir una fecha de envio. Revise.", vbExclamation
        PonerFoco txtCodigo(132)
        Exit Sub
    End If
    
     
     '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtCodigo(58).Text)
    cHasta = Trim(txtCodigo(59).Text)
    nDesde = txtNombre(58).Text
    nHasta = txtNombre(59).Text
    If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rsocios.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If

    '[Monica]11/05/2016: para Picassent la situacion del socio ha de ser 0
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        If Not AnyadirAFormula(cadSelect, "{rsocios.codsitua}=0") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.codsitua}=0") Then Exit Sub
    End If

    cadSelect1 = cadSelect

    'D/H Producto
    cDesde = Trim(txtCodigo(60).Text)
    cHasta = Trim(txtCodigo(61).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codprodu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
    End If
    
    vSQL = ""
    If txtCodigo(60).Text <> "" Then vSQL = vSQL & " and variedades.codprodu >= " & DBSet(txtCodigo(60).Text, "N")
    If txtCodigo(61).Text <> "" Then vSQL = vSQL & " and variedades.codprodu <= " & DBSet(txtCodigo(61).Text, "N")
    
    Set frmMens1 = New frmMensajes
    
    frmMens1.OpcionMensaje = 4
    frmMens1.cadWHERE = vSQL
    frmMens1.Show vbModal
    
    Set frmMens1 = Nothing

    Tabla1 = "rsocios INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & vParamAplic.Seccionhorto
    Tabla1 = Tabla1 & " and rsocios_seccion.fecbaja is null "
    
    tabla = "((" & Tabla1 & ") INNER JOIN rcampos ON rcampos.codsocio = rsocios.codsocio and rcampos.fecbajas is null "
    
    '[Monica]02/04/2014: para el caso de Picassent no miramos que no tenga supcoope <> 0
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        tabla = tabla & ") "
    Else
        tabla = tabla & " and rcampos.supcoope <> 0) "
    End If
    
    tabla = tabla & " INNER JOIN variedades on rcampos.codvarie = variedades.codvarie "
     
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(Tabla1, cadSelect1) Then
        B = GeneraFicheroTraspasoROPAS(Tabla1, cadSelect1, tabla, cadSelect)
        If B Then
            If CopiarFicheroROPAS() Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                CmdCancelROPAS_Click
            End If
        End If
    End If

End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vcad As String

    InicializarVbles
    
    ConSubInforme = False
    
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Select Case Index
       Case 0 'frame de campos
            '======== FORMULA  ====================================
            'D/H Socio
            cDesde = Trim(txtCodigo(2).Text)
            cHasta = Trim(txtCodigo(3).Text)
            nDesde = txtNombre(2).Text
            nHasta = txtNombre(3).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                '[Monica]10/06/2013: el enlace es con el propietario
                If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                    Codigo = "{rcampos.codpropiet}"
                Else
                    Codigo = "{rcampos.codsocio}"
                End If
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
            End If
       
            'D/H Clase
            cDesde = Trim(txtCodigo(0).Text)
            cHasta = Trim(txtCodigo(1).Text)
            nDesde = txtNombre(0).Text
            nHasta = txtNombre(1).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{variedades.codclase}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
            End If
    
            vSQL = ""
            If txtCodigo(0).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(0).Text, "N")
            If txtCodigo(1).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(1).Text, "N")
            
            'D/H Situacion
            cDesde = Trim(txtCodigo(4).Text)
            cHasta = Trim(txtCodigo(5).Text)
            nDesde = txtNombre(4).Text
            nHasta = txtNombre(5).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rcampos.codsitua}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSituacion= """) Then Exit Sub
            End If
    
            'D/H Capataz(Responsable)
            cDesde = Trim(txtCodigo(92).Text)
            cHasta = Trim(txtCodigo(93).Text)
            nDesde = txtNombre(92).Text
            nHasta = txtNombre(93).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rcampos.codcapat}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHResponsable= """) Then Exit Sub
            End If
    
            'D/H Partida
            cDesde = Trim(txtCodigo(94).Text)
            cHasta = Trim(txtCodigo(95).Text)
            nDesde = txtNombre(94).Text
            nHasta = txtNombre(95).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rcampos.codparti}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPartida= """) Then Exit Sub
            End If
            
            '[Monica]20/03/2013:
            'D/H Zonas
            cDesde = Trim(txtCodigo(133).Text)
            cHasta = Trim(txtCodigo(134).Text)
            nDesde = txtNombre(133).Text
            nHasta = txtNombre(134).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rcampos.codzonas}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHZonas= """) Then Exit Sub
            End If
    
    
            tabla = "(((rcampos INNER JOIN rpartida ON rcampos.codparti = rpartida.codparti) "
            tabla = tabla & " INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
            tabla = tabla & " INNER JOIN rzonas ON rcampos.codzonas = rzonas.codzonas) "
            tabla = tabla & " LEFT JOIN rcapataz ON rcampos.codcapat = rcapataz.codcapat "
            
            '[Monica]10/06/2013: a�adimos las condiciones de las cartas de talla Solo Escalona y Utxera
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                tabla = "(" & tabla & ") INNER JOIN rsocios ON rcampos.codpropiet = rsocios.codsocio "
                tabla = "(" & tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
                
                If Not AnyadirAFormula(cadSelect, "{rsituacion.bloqueo} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rsituacion.bloqueo} = 0") Then Exit Sub
                
                If Not AnyadirAFormula(cadSelect, "{rcampos.codsitua} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rcampos.codsitua} = 1") Then Exit Sub
            End If
            
            If Opcion1(0).Value Then numOp = PonerGrupo(1, "Socios")
            If Opcion1(1).Value Then numOp = PonerGrupo(1, "Clases")
            If Opcion1(2).Value Then numOp = PonerGrupo(1, "Terminos")
            If Opcion1(3).Value Then numOp = PonerGrupo(1, "Zonas")
            If Opcion1(7).Value Then numOp = PonerGrupo(1, "Variedad/Zona")
            
            If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub
            
            If Combo1(11).ListIndex < 3 Then
                If Not AnyadirAFormula(cadSelect, "{rcampos.tipocampo}=" & Combo1(11).ListIndex) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rcampos.tipocampo}=" & Combo1(11).ListIndex) Then Exit Sub
            End If
            CadParam = CadParam & "pTipo=" & Combo1(11).ListIndex & "|"
            numParam = numParam + 1
            
            cadTitulo = "Informe de Campos"
            If Opcion1(0).Value Then cadTitulo = cadTitulo & " por Socios"
            If Opcion1(1).Value Then cadTitulo = cadTitulo & " por Clases"
            If Opcion1(2).Value Then cadTitulo = cadTitulo & " por Terminos"
            '[Monica]07/06/2013: Zonas/ Bra�al
            If Opcion1(3).Value Then cadTitulo = cadTitulo & " por " & vParamAplic.NomZonaPOZ 'Zonas
            If Opcion1(4).Value Then cadTitulo = cadTitulo & " por Variedad/Respons./Partida"
            
            '[Monica]20/09/2013: Variedad zona Picassent
            If Opcion1(7).Value Then cadTitulo = cadTitulo & " por Variedad/Zona"
            
            'combo1(0): tipo de has
            CadParam = CadParam & "pTipoHas=" & Combo1(0).ListIndex & "|"
            numParam = numParam + 1
            
            'combo1(1): tipo de kilos 0=aforo 1=real
            CadParam = CadParam & "pKilos=" & Combo1(1).ListIndex & "|"
            numParam = numParam + 1
            
            ' Imprimir cabecera
            CadParam = CadParam & "pCabecera=" & Check4.Value & "|"
            numParam = numParam + 1
            
            '[Monica]06/09/2010: el informe original para todo el mundo es rInfCampos.rpt
            ' para Picassent es PicInfCampos.rpt
            cadNombreRPT = "rInfCampos.rpt"
            
            indRPT = 54 'Informe de campos / huertos
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
              
            'Nombre fichero .rpt a Imprimir
            cadNombreRPT = nomDocu
            ConSubInforme = False
            
            If Me.Check12.Value Then
                cadNombreRPT = Replace(nomDocu, "InfCampos.rpt", "InfCamposRecintos.rpt")
            End If
            
            
            '[Monica]22/12/2011: solo para picassent que tiene los reports con hdas
            CadParam = CadParam & "pHectareas=" & Format(Check16.Value, "0") & "|"
            numParam = numParam + 1
            
            ' resumen o no
            CadParam = CadParam & "pResumen=" & Format(Check1.Value, "0") & "|"
            numParam = numParam + 1
            
            '[Monica]03/06/2016: si se salta pagina por socio
            CadParam = CadParam & "pSalta=" & Format(Check26.Value, "0") & "|"
            numParam = numParam + 1
            
            Set frmMens = New frmMensajes
            
            frmMens.OpcionMensaje = 16
            frmMens.cadWHERE = vSQL
            frmMens.Show vbModal
            
            Set frmMens = Nothing
            
             'Comprobar si hay registros a Mostrar antes de abrir el Informe
            If HayRegParaInforme(tabla, cadSelect) Then
                If Opcion1(4).Value Then
                    If CargarTemporalCampos(tabla, cadSelect) Then
                        cadNombreRPT = "rInfCamposZonas.rpt"
                        indRPT = 66 'Informe de campos / huertos
                        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                          
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        ConSubInforme = True
    
                        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
    
                        With frmImprimir
                            .FormulaSeleccion = cadFormula
                            .OtrosParametros = CadParam
                            .NumeroParametros = numParam
                            .SoloImprimir = False
                            .EnvioEMail = False
                            .Titulo = cadTitulo
                            .NombreRPT = cadNombreRPT
                            .ConSubInforme = True
                            .Opcion = 0
                            .Show vbModal
                        End With
                    End If
                Else
                    '[Monica]23/07/2015: para el caso de picassent sacamos dni y termino municipal si est� marcado para conselleria
                    CadParam = CadParam & "pConselleria=" & Check23.Value & "|"
                    numParam = numParam + 1
                
                    If Me.Check12.Value = 1 Then
                            With frmImprimir
                                .FormulaSeleccion = cadFormula
                                .OtrosParametros = CadParam
                                .NumeroParametros = numParam
                                .SoloImprimir = False
                                .EnvioEMail = False
                                .Titulo = cadTitulo
                                .NombreRPT = cadNombreRPT
                                .ConSubInforme = True
                                .Opcion = 0
                                .Show vbModal
                            End With
                    Else
                        If CargarTemporal(tabla, cadSelect) Then
                            CadParam = CadParam & "pUsu=" & vUsu.Codigo & "|"
                            numParam = numParam + 1
        
                            With frmImprimir
                                .FormulaSeleccion = cadFormula
                                .OtrosParametros = CadParam
                                .NumeroParametros = numParam
                                .SoloImprimir = False
                                .EnvioEMail = False
                                .Titulo = cadTitulo
                                .NombreRPT = cadNombreRPT
                                .ConSubInforme = True
                                .Opcion = 0
                                .Show vbModal
                            End With
                        End If
                    End If
                End If
            End If
      
       Case 1 'Frame Informe de socios por seccion
            '======== FORMULA  ====================================
            'D/H Seccion
            cDesde = Trim(txtCodigo(8).Text)
            cHasta = Trim(txtCodigo(9).Text)
            nDesde = txtNombre(8).Text
            nHasta = txtNombre(9).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rsocios_seccion.codsecci}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSeccion= """) Then Exit Sub
            End If
            
            'D/H Socio
            cDesde = Trim(txtCodigo(10).Text)
            cHasta = Trim(txtCodigo(11).Text)
            nDesde = txtNombre(10).Text
            nHasta = txtNombre(11).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rsocios_seccion.codsocio}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
            End If
            
'[Monica] 16/09/2009 incluir los socios dados de baja
            If Check8.Value = 0 Then
                vcad = "isnull({rsocios_seccion.fecbaja})"
                If AnyadirAFormula(cadFormula, vcad) = False Then Exit Sub
                vcad = "rsocios_seccion.fecbaja is null"
                If AnyadirAFormula(cadSelect, vcad) = False Then Exit Sub
            End If
            
            '[Monica]21/03/2016: imprimir solo los socio de baja
            If Check24.Value Then
                vcad = "not isnull({rsocios_seccion.fecbaja})"
                If AnyadirAFormula(cadFormula, vcad) = False Then Exit Sub
                vcad = "not rsocios_seccion.fecbaja is null"
                If AnyadirAFormula(cadSelect, vcad) = False Then Exit Sub
            End If
            
            
            '[Monica]19/01/2012: insertamos las situaciones de socios que vamos a incluir
            
            Set frmMens4 = New frmMensajes
            
            frmMens4.OpcionMensaje = 36
            frmMens4.Show vbModal
            
            Set frmMens4 = Nothing
            
            tabla = "rsocios_seccion"
        
            '[Monica]08/04/2015: para el caso de catadau miramos el combo1(15)
            If vParamAplic.Cooperativa = 0 And Opcion(0).Value Then
                ' rsocios.tiporelacion puede tomar los valores: 0=socio, 1=asociado, 2=tercero
                Select Case Combo1(15).ListIndex
                    Case 0 ' todos
                    
                    Case 1 ' solo socios
                        vcad = "{rsocios.tiporelacion} = 0"
                        If AnyadirAFormula(cadFormula, vcad) = False Then Exit Sub
                        vcad = "rsocios.tiporelacion = 0"
                        If AnyadirAFormula(cadSelect, vcad) = False Then Exit Sub
                    
                    Case 2 ' solo asociados
                        vcad = "{rsocios.tiporelacion} = 1"
                        If AnyadirAFormula(cadFormula, vcad) = False Then Exit Sub
                        vcad = "rsocios.tiporelacion = 1"
                        If AnyadirAFormula(cadSelect, vcad) = False Then Exit Sub
                
                        CadParam = CadParam & "pAsociado=1|"
                        numParam = numParam + 1
                
                End Select
            End If
            
            
            'Obtener el parametro con el ORDEN del Informe
            '---------------------------------------------
        '    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
        '    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
        
            If Opcion(0).Value Then numOp = PonerGrupo(1, "Seccion")
            If Opcion(1).Value Then numOp = PonerGrupo(1, "Socio")
            
            cadNombreRPT = "rManSocSeccion.rpt"
            
            '[Monica]18/05/2012: personalizacion de informe de socios/seccion
            indRPT = 85
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
              
            'Nombre fichero .rpt a Imprimir
            cadNombreRPT = nomDocu
            
            
            If Opcion(0).Value Then cadTitulo = "Listado de Socios por Secci�n"
            If Opcion(1).Value Then cadTitulo = "Listado de Socios"
            
'[Monica] 23/08/2010: ordenado por socio o alfabeticamente
            ' por codigo
            If Opcion(5).Value Then
                If Opcion(1).Value Then ' por seccion
                    CadParam = CadParam & "pOrden={rsocios_seccion.codsecci}|"
                Else
                    CadParam = CadParam & "pOrden={rsocios_seccion.codsocio}|"
                End If
            End If
            ' alfabetico
            If Opcion(4).Value Then
                If Opcion(1).Value Then ' por seccion
                    CadParam = CadParam & "pOrden={rseccion.nomsecci}|"
                Else ' por socio
                    CadParam = CadParam & "pOrden={rsocios.nomsocio}|"
                End If
            End If
            numParam = numParam + 1
            
            tabla = "rsocios_seccion INNER JOIN rsocios ON rsocios_seccion.codsocio = rsocios.codsocio "
            
            
            '[Monica]10/03/2015: socios o.p. control democr�tico
            
            'Comprobar si hay registros a Mostrar antes de abrir el Informe
            If HayRegParaInforme(tabla, cadSelect) Then
                
                '[Monica]21/05/2012: cargamos los votos si es escalona
                If vParamAplic.Cooperativa = 10 Then
                    If CargarVotos(tabla, cadSelect) Then
                        cadTitulo = "Listado de Propietarios"
                        LlamarImprimir
                        Exit Sub
                    End If
                Else
                    If Check21.Value = 1 Then
                        indRPT = 107
                        cadTitulo = "Listado Socios OP control democr�tico"
                        
                        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
              
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        
                        If CargarTemporalMiembros(tabla, cadSelect) Then
                            CadParam = CadParam & "pUsu=" & vUsu.Codigo & "|"
                            numParam = numParam + 1
                            ConSubInforme = True
                            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                            LlamarImprimir
                            Exit Sub
                        End If
                        
                    End If
                End If
            
                LlamarImprimir
            End If
        
        Case 2 ' informe de calidades
            '======== FORMULA  ====================================
            'D/H Variedad
            cDesde = Trim(txtCodigo(18).Text)
            cHasta = Trim(txtCodigo(19).Text)
            nDesde = txtNombre(18).Text
            nHasta = txtNombre(19).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
            End If
            
            'D/H Calidad
            cDesde = Trim(txtCodigo(16).Text)
            cHasta = Trim(txtCodigo(17).Text)
            nDesde = txtNombre(16).Text
            nHasta = txtNombre(17).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codcalid}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCalidad= """) Then Exit Sub
            End If
            
            'Obtener el parametro con el ORDEN del Informe
            '---------------------------------------------
        '    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
        '    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
            If Opcion(2).Value Then numOp = PonerGrupo(1, "Variedad")
            If Opcion(3).Value Then numOp = PonerGrupo(1, "Calidad")
            
            cadNombreRPT = "rManVarCalidad.rpt"
            
            If Opcion(2).Value Then cadTitulo = "Listado de Calidades por Variedad"
            If Opcion(3).Value Then cadTitulo = "Listado de Calidades"
            
            'Comprobar si hay registros a Mostrar antes de abrir el Informe
            If HayRegParaInforme(tabla, cadSelect) Then
                LlamarImprimir
            End If
            
            
        Case 3 ' informe de entradas de bascula
            '======== FORMULA  ====================================
            'D/H SOCIO
            cDesde = Trim(txtCodigo(12).Text)
            cHasta = Trim(txtCodigo(13).Text)
            nDesde = txtNombre(12).Text
            nHasta = txtNombre(13).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codsocio}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
            End If
            
            '[Monica]25/11/2011: Modificacion para poder seleccionar los campos de cada socio variedad
            vSql2 = ""
            If txtCodigo(12).Text <> "" Then vSql2 = vSql2 & " and rcampos.codsocio >= " & DBSet(txtCodigo(12).Text, "N")
            If txtCodigo(13).Text <> "" Then vSql2 = vSql2 & " and rcampos.codsocio <= " & DBSet(txtCodigo(13).Text, "N")
            
            '[Monica]17/07/2014: a�adido el tipo de socio
            If OpcionListado = 18 Then
                Select Case Combo1(14).ListIndex
                    Case 0, 1, 2, 3
                        vSql2 = vSql2 & " and rcampos.codsocio in (select codsocio from rsocios where tipoprod = " & Combo1(14).ListIndex & ")"
                    Case 4 ' todos
                
                End Select
            End If
            
            
            'D/H CLASE
            cDesde = Trim(txtCodigo(20).Text)
            cHasta = Trim(txtCodigo(21).Text)
            nDesde = txtNombre(20).Text
            nHasta = txtNombre(21).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{variedades.codclase}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
            End If
            
            vSQL = ""
            If txtCodigo(20).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(20).Text, "N")
            If txtCodigo(21).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(21).Text, "N")
                        
            
            'D/H VARIEDAD
            cDesde = Trim(txtCodigo(14).Text)
            cHasta = Trim(txtCodigo(15).Text)
            nDesde = txtNombre(14).Text
            nHasta = txtNombre(15).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
            End If
            
            If txtCodigo(14).Text <> "" Then vSQL = vSQL & " and variedades.codvarie >= " & DBSet(txtCodigo(14).Text, "N")
            If txtCodigo(15).Text <> "" Then vSQL = vSQL & " and variedades.codvarie <= " & DBSet(txtCodigo(15).Text, "N")

            '[Monica]25/11/2011: poder seleccionar los campos
            If vSQL <> "" Then vSql2 = vSql2 & vSQL

            'D/H fecha
            cDesde = Trim(txtCodigo(6).Text)
            cHasta = Trim(txtCodigo(7).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                Select Case OpcionListado
                    Case 10, 14, 16
                        'Cadena para seleccion Desde y Hasta
                        Codigo = "{" & tabla & ".fechaent}"
                    Case 17, 18
                        'Cadena para seleccion Desde y Hasta
                        Codigo = "{" & tabla & ".fecalbar}"
                End Select
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
                
            Select Case OpcionListado
              Case 10 ' Reimpresion de entradas de bascula
                nTabla = "(rentradas INNER JOIN variedades ON rentradas.codvarie = variedades.codvarie) "
    
     
                indRPT = 25 'Ticket de Entrada
     
                If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub
    
                cadNombreRPT = nomDocu
                
                cadTitulo = "Reimpresi�n de Entradas B�scula"
                
                ConSubInforme = True
            
            
                'Comprobar si hay registros a Mostrar antes de abrir el Informe
                If HayRegParaInforme(nTabla, cadSelect) Then
                    LlamarImprimir
                End If
                
              Case 14 ' listado de entradas (rentradas)
                ' resumen o no
                CadParam = CadParam & "pResumen=" & Format(Check2.Value, "0") & "|"
                numParam = numParam + 1
                
                nTabla = "(rentradas INNER JOIN variedades ON rentradas.codvarie = variedades.codvarie) "
    
                '[Monica]20/09/2016: personalizacion del informe de entradas b�scula
                indRPT = 109
    
                If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub
    
                cadNombreRPT = nomDocu ' "rInfEntradas.rpt"
                cadTitulo = "Informe de Entradas B�scula"
                
                ConSubInforme = True
            
            
                'Comprobar si hay registros a Mostrar antes de abrir el Informe
                If HayRegParaInforme(nTabla, cadSelect) Then
                    LlamarImprimir
                End If
            
              Case 16 ' listado de entradas clasificadas (rclasifica)
                nTabla = "(rclasifica INNER JOIN variedades ON rclasifica.codvarie = variedades.codvarie) "
                
                
                '[Monica]25/11/2011: tambien quieren puntear que campos quieren incluir
                Set frmMens3 = New frmMensajes
     
                frmMens3.OpcionMensaje = 34
                frmMens3.cadWHERE = vSql2
                frmMens3.Show vbModal
                
                Set frmMens3 = Nothing
                
                '[Monica]06/10/2016: incluimos los contratos para poder seleccionar
                If vParamAplic.Cooperativa = 16 Then
                   Set frmMens7 = New frmMensajes
        
                   frmMens7.OpcionMensaje = 64
                   frmMens7.Show vbModal
                   
                   Set frmMens7 = Nothing
                End If
                
                indRPT = 56 'Informe de entradas clasificadas
     
                If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub
    
                cadNombreRPT = nomDocu
                
                Select Case Combo1(3).ListIndex
                    Case 0 ' informe normal
                        If CargarTemporal2(nTabla, cadSelect) Then
'                            cadNombreRPT = "rInfEntradasClas.rpt"
                            cadTitulo = "Informe de Entradas Clasificadas"
                            
                            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                            'Comprobar si hay registros a Mostrar antes de abrir el Informe
                            If HayRegParaInforme("tmpinformes", "{tmpinformes.codusu} = " & vUsu.Codigo) Then
                                LlamarImprimir
                            End If
                        End If
                    Case 1 ' informe detalle clasificacion
                        If CargarTemporal3(nTabla, cadSelect) Then
                            cadNombreRPT = Replace(cadNombreRPT, "InfEntradasClas.rpt", "InfEntradasClas1.rpt") '"rInfEntradasClas1.rpt"
                            cadTitulo = "Informe de Entradas Clasificadas"
                            
                            cadFormula = "{tmpclasifica.codusu} = " & vUsu.Codigo
                            'Comprobar si hay registros a Mostrar antes de abrir el Informe
                            If HayRegParaInforme("tmpclasifica", "{tmpclasifica.codusu} = " & vUsu.Codigo) Then
                                With frmImprimir
                                    .FormulaSeleccion = cadFormula
                                    .OtrosParametros = CadParam
                                    .NumeroParametros = numParam
                                    .SoloImprimir = False
                                    .EnvioEMail = False
                                    .Titulo = cadTitulo
                                    .NombreRPT = cadNombreRPT
                                    .ConSubInforme = True
                                    .Opcion = OpcionListado
                                    .Show vbModal
                                End With
                            End If
                        End If
                End Select
                
              Case 17 ' reimpresion de albaranes (rhisfruta)
                nTabla = "(rhisfruta INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
                
                CadParam = CadParam & "pDuplicado=0|"
                numParam = numParam + 1
                
                indRPT = 22 'Impresion de Albaran de clasificacion
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                  
                'Nombre fichero .rpt a Imprimir
'                frmImprimir.NombreRPT = nomDocu
                cadNombreRPT = nomDocu
'                cadNombreRPT = "rInfEntradas.rpt"
                cadTitulo = "Impresion de Albaranes"
'                OpcionListado = 22
                ConSubInforme = True
                
'[Monica]09/06/2010: he sustituido esto por un combo en el que se decide si los no impresos, los impresos o todos
'                If Not AnyadirAFormula(cadFormula, "{rhisfruta.impreso} = 0") Then Exit Sub
'                If Not AnyadirAFormula(cadSelect, "{rhisfruta.impreso} = 0") Then Exit Sub
'
                Select Case Combo1(10).ListIndex
                    Case 0
                        If Not AnyadirAFormula(cadFormula, "{rhisfruta.impreso} = 0") Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, "{rhisfruta.impreso} = 0") Then Exit Sub
                    Case 1
                        If Not AnyadirAFormula(cadFormula, "{rhisfruta.impreso} = 1") Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, "{rhisfruta.impreso} = 1") Then Exit Sub
                    Case 2
                    
                End Select


                'Comprobar si hay registros a Mostrar antes de abrir el Informe
                If HayRegParaInforme(nTabla, cadSelect) Then
                    LlamarImprimir
                    
                    If MsgBox("� Impresi�n correcta para actualizar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        If ActualizarRegistros(tabla, cadSelect) Then
                            MsgBox "Proceso realizado correctamente.", vbExclamation
                            cmdCancel_Click (0)
                        End If
                    End If
                End If
              
              Case 18 ' informe de kilos/gastos
                nTabla = "((rhisfruta INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) INNER JOIN rsocios ON rhisfruta.codsocio=rsocios.codsocio) "
                
                ' a�adido el combo del tipo de entradas recolectadas
                Select Case Combo1(8).ListIndex
                    Case 0, 1
                        If Not AnyadirAFormula(cadSelect, " rhisfruta.recolect = " & Combo1(8).ListIndex) Then Exit Sub
                    Case 2
                    
                End Select
                
                
                ' a�adido el combo del tipo de entradas
                Select Case Combo1(9).ListIndex
                    Case 0, 1, 2, 3, 4, 5
                        If Not AnyadirAFormula(cadSelect, " rhisfruta.tipoentr = " & Combo1(9).ListIndex) Then Exit Sub
                    Case 6
                    
                End Select
                
                '[Monica]17/07/2014: a�adido el tipo de productor (socio)
                ' a�adido el combo del tipo de socio
                Select Case Combo1(14).ListIndex
                    Case 0, 1, 2, 3
                        If Not AnyadirAFormula(cadSelect, " rsocios.tipoprod = " & Combo1(14).ListIndex) Then Exit Sub
                    Case 4
                    
                End Select
                
                '[Monica]01/02/2011: de las variedades quieren puntear cuales quieren incluir
                '******************
                Set frmMens = New frmMensajes
     
                frmMens.OpcionMensaje = 16
                frmMens.cadWHERE = vSQL
                frmMens.Show vbModal
                
                Set frmMens = Nothing

                
                If vParamAplic.Cooperativa <> 12 Then
                       '[Monica]25/11/2011: tambien quieren puntear que campos quieren incluir
                       Set frmMens2 = New frmMensajes
            
                       frmMens2.OpcionMensaje = 34
                       frmMens2.cadWHERE = vSql2
                       frmMens2.Show vbModal
                       
                       Set frmMens2 = Nothing
                End If
                
                '[Monica]30/12/2016: incluimos los contratos para poder seleccionar
                If vParamAplic.Cooperativa = 16 Then
                   Contratos = ""
                
                   Set frmMens8 = New frmMensajes
                   
                   frmMens8.desdeHco = True
                   frmMens8.OpcionMensaje = 64
                   frmMens8.Show vbModal
                   
                   Set frmMens8 = Nothing
                   
                   If Contratos <> "" Then
                        If InStr(UCase(Contratos), "NULL") <> 0 Then
                            vcad = "(rhisfruta.contrato is null or rhisfruta.contrato in (" & Contratos & "))"
                        Else
                            vcad = "(rhisfruta.contrato in (" & Contratos & "))"
                        End If
                        If Not AnyadirAFormula(cadSelect, vcad) Then Exit Sub
                   End If
                   
                End If
                
                If CargarTemporal4New(nTabla, cadSelect) Then
                    CadParam = CadParam & "pRecolectado=" & Combo1(8).ListIndex & "|"
                    numParam = numParam + 1
                    
                    CadParam = CadParam & "pTipoEntrada=" & Combo1(9).ListIndex & "|"
                    numParam = numParam + 1
                                            
                    '[Monica]17/07/2014: a�adido el tipo de socio
                    CadParam = CadParam & "pTipoSocio=" & Combo1(14).ListIndex & "|"
                    numParam = numParam + 1
                                            
                    '[Monica]25/11/2011: he sacado de dentro de check5.value = 1
                    indRPT = 53 'Informe de Kilos Gastos por socio
                    If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub
                    
                    If Check5.Value = 1 Then
                        ' imprimimos por socio
                        CadParam = CadParam & "pSaltar=" & Check6.Value & "|"
                        numParam = numParam + 1
                       '[Monica] 01/10/2009 a�adido el poder detallar las notas
                        CadParam = CadParam & "pDetalleNota=" & Check9.Value & "|"
                        numParam = numParam + 1
                        
                        If Check2.Value = 1 Then
                            '[Monica]01/02/2011: a�adido el caso de Picassent, agrupado por socio/variedad/campo
                            cadNombreRPT = Replace(nomDocu, 2, 4)
                            
                            CadParam = CadParam & "pOmitirGastos=" & Check10.Value & "|"
                            numParam = numParam + 1
                        Else
                            If Check10.Value = 0 Then
                                '[Monica]27/07/2010: personalizado
                                'cadNombreRPT = "rInfHcoEntClas2.rpt"
                                cadNombreRPT = nomDocu
                            Else
                                '[Monica]27/07/2010: personalizado
                                'cadNombreRPT = "rInfHcoEntClas3.rpt"
                                cadNombreRPT = Replace(nomDocu, 2, 3)
                            End If
                        End If
                    Else
                        If Check2.Value = 0 Then
                            cadNombreRPT = Replace(nomDocu, "2.rpt", ".rpt") '"rInfHcoEntClas.rpt"
                            '[Monica] 01/10/2009 a�adido el poder detallar las notas
                             CadParam = CadParam & "pDetalleNota=" & Check9.Value & "|"
                             numParam = numParam + 1
                        Else
                            ' imprimimos un resumen por variedad
                            cadNombreRPT = Replace(nomDocu, "2.rpt", "1.rpt") '"rInfHcoEntClas1.rpt"
                        End If
                    End If
                    cadTitulo = "Informe de Kilos / Gastos"
                    
                    cadFormula = "{tmpclasifica2.codusu} = " & vUsu.Codigo
                    'Comprobar si hay registros a Mostrar antes de abrir el Informe
                    If HayRegParaInforme("tmpclasifica2", "{tmpclasifica2.codusu} = " & vUsu.Codigo) Then
                        With frmImprimir
                            .FormulaSeleccion = cadFormula
                            .OtrosParametros = CadParam
                            .NumeroParametros = numParam
                            .SoloImprimir = False
                            .EnvioEMail = False
                            .Titulo = cadTitulo
                            .NombreRPT = cadNombreRPT
                            .ConSubInforme = True
                            .Opcion = OpcionListado
                            .Show vbModal
                        End With
                    End If
                End If
              
            End Select
    
    
        Case 4 ' Informe de pesadas
        
            'D/H Pesada
            cDesde = Trim(txtCodigo(70).Text)
            cHasta = Trim(txtCodigo(71).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rpesadas.nropesada}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPesada=""") Then Exit Sub
            End If
            
            'D/H CLASE
            cDesde = Trim(txtCodigo(66).Text)
            cHasta = Trim(txtCodigo(67).Text)
            nDesde = txtNombre(66).Text
            nHasta = txtNombre(67).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{variedades.codclase}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
            End If
            
            'D/H VARIEDAD
            cDesde = Trim(txtCodigo(68).Text)
            cHasta = Trim(txtCodigo(69).Text)
            nDesde = txtNombre(68).Text
            nHasta = txtNombre(69).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{variedades.codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
            End If

            'D/H fecha
            cDesde = Trim(txtCodigo(72).Text)
            cHasta = Trim(txtCodigo(73).Text)
            nDesde = ""
            nHasta = ""
            'Cadena para seleccion Desde y Hasta
            If Not (cDesde = "" And cHasta = "") Then
                Codigo = "{" & tabla & ".fecpesada}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
        
            CadParam = CadParam & "pResumen=" & Format(Check13.Value, "0") & "|"
            numParam = numParam + 1
            
            nTabla = "rpesadas INNER JOIN rentradas ON rpesadas.nropesada = rentradas.nropesada "
            nTabla = "(" & nTabla & ") INNER JOIN variedades ON rentradas.codvarie = variedades.codvarie "

            cadNombreRPT = "rInfPesadas.rpt"
            cadTitulo = "Informe de Entradas de Pesadas"
            
            ConSubInforme = True
        
        
            'Comprobar si hay registros a Mostrar antes de abrir el Informe
            If HayRegParaInforme(nTabla, cadSelect) Then
                LlamarImprimir
            End If
    
        
    
    End Select
End Sub

Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView1
End Sub

Private Sub cmdAcepTras_Click()
Dim sql As String
Dim I As Byte
Dim cadWHERE As String
Dim B As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String
Dim Directorio As String
Dim fec As String

Dim File1 As FileSystemObject

On Error GoTo eError

    If Not DatosOK Then Exit Sub

    
    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist


    If Combo1(6).ListIndex = 2 And vParamAplic.Cooperativa = 4 Then
    ' solo para el calibrador de alzira de kaki la extension es diferente
        Me.CommonDialog1.DefaultExt = "pdt"
        CommonDialog1.Filter = "Archivos PTD|*.ptd|"
        CommonDialog1.FilterIndex = 1
        Me.CommonDialog1.FileName = "*.ptd"
    Else
        If vParamAplic.Cooperativa = 2 Then
            Me.CommonDialog1.DefaultExt = "tag"
            CommonDialog1.Filter = "Archivos TAG|*.tag|"
            CommonDialog1.FilterIndex = 1
            Me.CommonDialog1.FileName = "*.tag"
        Else
            Me.CommonDialog1.DefaultExt = "txt"
            CommonDialog1.Filter = "Archivos TXT|*.txt|"
            CommonDialog1.FilterIndex = 1
            Me.CommonDialog1.FileName = "*.txt"
        End If
    End If
    
    '[Monica]21/04/2016: metemos esto dentro del if de castellduc, solo si no es castellduc con combo1(6) = 0 abrimos conexion
    If (vParamAplic.Cooperativa = 5 And (Combo1(6).ListIndex = 0 Or Combo1(6).ListIndex = 1)) Then
        If AbrirConexionSqlSERVER(Combo1(6).ListIndex) Then
            If Not CargarTablaCalibrador Then Exit Sub
        Else
            Exit Sub
        End If
    Else
        Me.CommonDialog1.CancelError = True
        Me.CommonDialog1.ShowOpen
        Set File1 = New FileSystemObject
    
        Directorio = File1.GetParentFolderName(Me.CommonDialog1.FileName)
    End If
        

    Select Case vParamAplic.Cooperativa
        '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
        Case 0, 9  '******* CATADAU *******
'             Directorio = GetFolder("Selecciona directorio")
            If Directorio <> "" Then
                sql = "DROP TABLE IF EXISTS tmpCata; "
                conn.Execute sql
                
                sql = "CREATE TEMPORARY TABLE tmpCata ("
                sql = sql & " codcalid int, kilosnet decimal(10,2)) "
                conn.Execute sql
                
                If Combo1(6).ListIndex = 1 Then ' si calibrador peque�o
                    'creamos la tabla temporal solo si estamos en calibrador peque�o
                    sql = "DROP TABLE IF EXISTS tmpcalibrador; "
                    conn.Execute sql
                    
                    sql = "CREATE TEMPORARY TABLE `tmpcalibrador` ("
                    sql = sql & "`numnota` varchar(10) default NULL, "
                    sql = sql & "`fecnota` varchar(20) default NULL, "
                    sql = sql & "`albaran` varchar(20) default NULL, "
                    sql = sql & "`porcen1` varchar(10) default NULL, "
                    sql = sql & "`porcen2` varchar(10) default NULL, "
                    sql = sql & "`kilos1` varchar(30) default NULL, "
                    sql = sql & "`kilos2` varchar(30) default NULL, "
                    sql = sql & "`kilos3` varchar(30) default NULL, "
                    sql = sql & "`numnota2` varchar(10) default NULL, "
                    sql = sql & "`export` varchar(10) default NULL, "
                    sql = sql & "`nomcalid` varchar(30) default NULL, "
                    sql = sql & "`kilos4` varchar(30) default NULL, "
                    sql = sql & "`kilos5` varchar(30) default NULL "
                    sql = sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
                
                    conn.Execute sql
                    
                    fec = ""
                Else
                    If txtCodigo(63).Text <> "" Then
                        fec = Format(txtCodigo(63).Text, "yyyymmdd")
                    End If
                End If
                
                conn.BeginTrans

                B = ProcesarDirectorioCatadau(Directorio & "\", Combo1(6).ListIndex, fec, Pb1, lblProgres(0), lblProgres(1))
            End If
        
        Case 1 '********* VALSUR *************
            CommonDialog1.FilterIndex = 1
            Me.CommonDialog1.ShowOpen
            
            If Me.CommonDialog1.FileName <> "" Then
                InicializarVbles
        '        InicializarTabla
                    '========= PARAMETROS  =============================
                'A�adir el parametro de Empresa
                CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
                numParam = numParam + 1
                '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
                If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 9 Then
                    sql = "DROP TABLE IF EXISTS tmpCata; "
                    conn.Execute sql
                    
                    sql = "CREATE TEMPORARY TABLE tmpCata ("
                    sql = sql & " codcalid int, kilosnet decimal(10,2)) "
                    conn.Execute sql
                End If
    
                conn.BeginTrans
                ' resto de casos (incluido catadau, calibrador grande)
                B = ProcesarFichero(Me.CommonDialog1.FileName, Combo1(6).ListIndex, Me.Pb1, Me.lblProgres(0), Me.lblProgres(1))
            Else
                MsgBox "No ha seleccionado ning�n fichero", vbExclamation
                Exit Sub
            End If
    
        Case 2 ' ******** PICASSENT **********
            If Directorio <> "" Then

                sql = "DROP TABLE IF EXISTS tmpCata; "
                conn.Execute sql
                
                sql = "CREATE TEMPORARY TABLE tmpCata ("
                sql = sql & " codcalid int, kilosnet decimal(10,2)) "
                conn.Execute sql
                
                
                If Combo1(6).ListIndex = 0 Then
                    'creamos la tabla temporal solo si estamos en precalibrado
                    sql = "DROP TABLE IF EXISTS tmpcalibrador; "
                    conn.Execute sql
                    
                    sql = "CREATE TEMPORARY TABLE `tmpcalibrador` ("
                    sql = sql & "`numnota` varchar(10) default NULL, "
                    sql = sql & "`fecnota` varchar(20) default NULL, "
                    sql = sql & "`nomcalid` varchar(30) default NULL, "
                    sql = sql & "`kilos1` varchar(30) default NULL, "
                    sql = sql & "`kilos2` varchar(30) default NULL, "
                    sql = sql & "`kilos3` varchar(30) default NULL, "
                    sql = sql & "`kilos4` varchar(30) default NULL "
                    sql = sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
                
                    conn.Execute sql
                End If
            
                conn.BeginTrans

                B = ProcesarDirectorioPicassent(Directorio & "\", Combo1(6).ListIndex, Pb1, lblProgres(0), lblProgres(1))
            End If
     
        Case 16 ' ******** COOPIC **********
            If Directorio <> "" Then

                sql = "DROP TABLE IF EXISTS tmpCata; "
                conn.Execute sql
                
                sql = "CREATE TEMPORARY TABLE tmpCata ("
                sql = sql & " codcalid int, kilosnet decimal(10,2)) "
                conn.Execute sql
                
                
                If Combo1(6).ListIndex = 0 Then
                    'creamos la tabla temporal solo si estamos en precalibrado
                    sql = "DROP TABLE IF EXISTS tmpcalibrador; "
                    conn.Execute sql
                    
                    sql = "CREATE TEMPORARY TABLE `tmpcalibrador` ("
                    sql = sql & "`numnota` varchar(10) default NULL, "
                    sql = sql & "`fecnota` varchar(20) default NULL, "
                    sql = sql & "`nomcalid` varchar(30) default NULL, "
                    sql = sql & "`kilos1` varchar(30) default NULL, "
                    sql = sql & "`kilos2` varchar(30) default NULL, "
                    sql = sql & "`kilos3` varchar(30) default NULL, "
                    sql = sql & "`kilos4` varchar(30) default NULL "
                    sql = sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
                
                    conn.Execute sql
                End If
            
                conn.BeginTrans

                B = ProcesarDirectorioCOOPIC(Me.CommonDialog1.FileName, Combo1(6).ListIndex, Pb1, lblProgres(0), lblProgres(1))
            End If
     
     
        Case 4 ' ******** ALZIRA **********
            If Directorio <> "" Then

                sql = "DROP TABLE IF EXISTS tmpCata; "
                conn.Execute sql
                
                sql = "CREATE TEMPORARY TABLE tmpCata ("
                sql = sql & " codcalid int, kilosnet decimal(10,2)) "
                conn.Execute sql
                
                
                If Combo1(6).ListIndex = 0 Then
                    'creamos la tabla temporal solo si estamos en precalibrado
                    sql = "DROP TABLE IF EXISTS tmpcalibrador; "
                    conn.Execute sql
                    
                    sql = "CREATE TEMPORARY TABLE `tmpcalibrador` ("
                    sql = sql & "`numnota` varchar(10) default NULL, "
                    sql = sql & "`fecnota` varchar(20) default NULL, "
                    sql = sql & "`nomcalid` varchar(30) default NULL, "
                    sql = sql & "`kilos1` varchar(30) default NULL, "
                    sql = sql & "`kilos2` varchar(30) default NULL, "
                    sql = sql & "`kilos3` varchar(30) default NULL, "
                    sql = sql & "`kilos4` varchar(30) default NULL "
                    sql = sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
                
                    conn.Execute sql
                End If
            
                conn.BeginTrans

                B = ProcesarDirectorioAlzira(Directorio & "\", Combo1(6).ListIndex, Pb1, lblProgres(0), lblProgres(1))
            End If
    
        Case 5 ' ******** CASTELDUC **********
            If Directorio <> "" Or Combo1(6).ListIndex = 0 Or Combo1(6).ListIndex = 1 Then

                sql = "DROP TABLE IF EXISTS tmpCata; "
                conn.Execute sql
                
                sql = "CREATE TEMPORARY TABLE tmpCata ("
                sql = sql & " codcalid int, kilosnet decimal(10,2)) "
                conn.Execute sql
                
                
                If Combo1(6).ListIndex = 0 Or Combo1(6).ListIndex = 1 Then
                    'creamos la tabla temporal solo si estamos en precalibrado
                    sql = "DROP TABLE IF EXISTS tmpcalibrador; "
                    conn.Execute sql
                    
                    sql = "CREATE TEMPORARY TABLE `tmpcalibrador` ("
                    sql = sql & "`numnota` varchar(10) default NULL, "
                    sql = sql & "`fecnota` varchar(20) default NULL, "
                    sql = sql & "`nomcalid` varchar(30) default NULL, "
                    sql = sql & "`kilos1` varchar(30) default NULL, "
                    sql = sql & "`kilos2` varchar(30) default NULL, "
                    sql = sql & "`kilos3` varchar(30) default NULL, "
                    sql = sql & "`kilos4` varchar(30) default NULL "
                    sql = sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
                
                    conn.Execute sql
                End If
            
                conn.BeginTrans

                B = ProcesarDirectorioCastelduc(Directorio & "\", Combo1(6).ListIndex, Pb1, lblProgres(0), lblProgres(1), txtCodigo(170).Text, txtCodigo(179).Text)
            End If
    
    
    End Select
    
eError:
    If Err.Number = 32755 Then Exit Sub ' le han dado a cancelar

    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso.", vbExclamation
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
'        BorrarArchivo Me.CommonDialog1.FileName
        cmdCancelTras_Click
    End If
    
End Sub

Private Function CargarTablaCalibrador()
Dim sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim CadInsert As String
Dim CadValues As String
Dim fecha1 As Date

    On Error GoTo eCargarTabla



    Pb1.visible = False
    
    Me.Refresh
    DoEvents
    
    
    CargarTablaCalibrador = False

    sql = "delete from tmpcalibradorcast where codusu = " & vUsu.Codigo
    conn.Execute sql


    CadInsert = "insert into tmpcalibradorcast (codusu,numnotac,numcalid,nomcalid,kilos) values "
    CadValues = ""
    
    sql = "select numero_lot, reference, dechet from lotapport where reference >= " & DBSet(txtCodigo(170).Text, "T") & " and reference <= " & DBSet(txtCodigo(179).Text, "T") '(Date_recolte = CONVERT(DATETIME, '" & Format(txtcodigo(63).Text, "yyyy-mm-dd") & " 00:00:00', 102)) "
    sql = sql & " order by 1 "
    
    
    Set RS = New ADODB.Recordset
    RS.Open sql, CnnSqlServer, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not RS.EOF
        lblProgres(0).Caption = "Nota : " & RS!reference
        Me.Refresh
        DoEvents
    
        sql = "select * from lotapportresultat where Numero_lot = " & DBSet(RS!Numero_lot, "N")
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open sql, CnnSqlServer, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(RS!reference, "N") & ",-1,'DESTRIO',"
        CadValues = CadValues & DBSet(RS!dechet, "N") & "),"
        
        
        While Not Rs2.EOF
            CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(RS!reference, "N") & "," & DBSet(Rs2!Num_calibre, "N") & ","
            CadValues = CadValues & DBSet(Rs2!nom_calibre, "T") & "," & DBSet(Rs2!poids, "N") & "),"
            
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If CadValues <> "" Then
        conn.Execute CadInsert & Mid(CadValues, 1, Len(CadValues) - 1)
    End If
    
    Pb1.visible = False
    lblProgres(0).visible = False
    
    CargarTablaCalibrador = True
    Exit Function

eCargarTabla:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Cargar Tabla Intermedia", Err.Description
    End If
End Function




Private Sub cmdAcepTrasCoop_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
    ' Cooperativa
    If Not AnyadirAFormula(cadSelect, "{rsocios.codcoope} = " & DBSet(txtCodigo(45).Text, "N")) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "{rsocios.codcoope} = " & DBSet(txtCodigo(45).Text, "N")) Then Exit Sub
     
    'D/H Socio
    cDesde = Trim(txtCodigo(48).Text)
    cHasta = Trim(txtCodigo(49).Text)
    nDesde = txtNombre(48).Text
    nHasta = txtNombre(49).Text
    If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rfactsoc.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If

    'D/H Fecha de Factura
    cDesde = Trim(txtCodigo(43).Text)
    cHasta = Trim(txtCodigo(44).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= """) Then Exit Sub
    End If

    'D/H Factura
    cDesde = Trim(txtCodigo(41).Text)
    cHasta = Trim(txtCodigo(42).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
     
    ' Tipo de Factura
    If Not AnyadirAFormula(cadSelect, "{rfactsoc.codtipom} = """ & Mid(TextoCombo(Combo1(7)), 1, 3) & """") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "{rfactsoc.codtipom}  = """ & Mid(TextoCombo(Combo1(7)), 1, 3) & """") Then Exit Sub
     
    tabla = "rfactsoc INNER JOIN rsocios ON rfactsoc.codsocio = rsocios.codsocio"
     
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If HayRegParaInforme(tabla, cadSelect) Then
        B = GeneraFicheroTraspasoCoop(tabla, cadSelect)
        If B Then
            If CopiarFicheroCoop(txtCodigo(45).Text) Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                cmdCancelTrasCoop_Click
            End If
        End If
     End If



End Sub

Private Sub CmdAcepTrasRetirada_Click()
Dim sql As String
Dim I As Byte
Dim cadWHERE As String
Dim B As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String
Dim Directorio As String
Dim fec As String
Dim cadTabla As String

Dim File1 As FileSystemObject

On Error GoTo eError

    If Not DatosOK Then Exit Sub


    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist


    Me.CommonDialog1.DefaultExt = "csv"
    CommonDialog1.Filter = "Archivos CSV|*.csv|"
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "*.csv"

    CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowOpen

    If Me.CommonDialog1.FileName <> "" Then
    
            '========= PARAMETROS  =============================
        'A�adir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

        If ProcesarFichero2(Me.CommonDialog1.FileName) Then
              cadTabla = "tmpinformes"
              cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo

              sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo

              If TotalRegistros(sql) <> 0 Then
                  MsgBox "Hay errores en el Traspaso de albaranes de Retirada. Debe corregirlos previamente.", vbExclamation
                  cadTitulo = "Errores de Traspaso Albaranes Retirada"
                  cadNombreRPT = "rErroresTrasAlbaranes.rpt"
                  LlamarImprimir
                  Exit Sub
              Else
                  B = ProcesarFicheroRetirada(Me.CommonDialog1.FileName)
              End If
        End If
    Else
        MsgBox "No ha seleccionado ning�n fichero", vbExclamation
        Exit Sub
    End If

eError:
    If Err.Number = 32755 Then Exit Sub
    cmdCancel_Click (0)
End Sub

Private Function ProcesarFichero2(nomFich As String) As String
Dim sql As String
Dim NF As Integer
Dim cad As String
Dim I As Long
Dim longitud As Long
Dim B As Boolean

    On Error GoTo eProcesarFichero2

    ProcesarFichero2 = False
    
    sql = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute sql
    
    NF = FreeFile
    Open nomFich For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    lblProgres(4).Caption = "Comprobando datos: " & nomFich
    longitud = FileLen(nomFich)
    
    pb9.visible = True
    Me.pb9.Max = longitud
    Me.Refresh
    Me.pb9.Value = 0

    B = True

    CifEmpre = DevuelveValor("select cifcoope from rcoope where codcoope = " & DBSet(txtCodigo(169).Text, "N"))

    While Not EOF(NF) And B
        I = I + 1
        
        B = ComprobarLinea(cad)
        
        Me.pb9.Value = Me.pb9.Value + Len(cad)
        lblProgres(5).Caption = "Linea " & I
        Me.Refresh
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        I = I + 1
        
        Me.pb9.Value = Me.pb9.Value + Len(cad)
        lblProgres(5).Caption = "Linea " & I
        Me.Refresh
        
        B = ComprobarLinea(cad)
        
    End If
    
    pb9.visible = False
    lblProgres(4).Caption = ""
    lblProgres(5).Caption = ""

    ProcesarFichero2 = B
    Exit Function

eProcesarFichero2:

End Function

Private Function ComprobarLinea(vCadena As String) As String
Dim Albaran As String
Dim Fecha As String
Dim sql As String
Dim Sql2 As String
Dim Mens As String
Dim Socio As Long
Dim Cifsocio As String

    On Error GoTo eComprobarLinea


    ComprobarLinea = False
        
    Albaran = RecuperaValorNew(vCadena, ";", 2)
    Fecha = RecuperaValorNew(vCadena, ";", 4)
    
    sql = RecuperaValorNew(vCadena, ";", 1) ' cif de la empresa
    If CifEmpre <> sql Then
        Mens = "El Cif " & sql & " no es de la cooperativa"
        sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute sql
    End If
    
' me da igual el contador de David, tengo que poner yo mi contador de la cooperativa
    ' albaran
    If Not EsNumerico(Albaran) Then
        Mens = "Albar�n no num�rico " & Albaran
        sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"

        conn.Execute sql
    Else
        If Len(Albaran) > 7 Then
            Mens = "Albar�n de m�s de 7 digitos " & Albaran
            sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"

            conn.Execute sql
        Else
            Select Case txtCodigo(169)
                Case 1
                    Albaran = "1" & Mid(Format(Albaran, "0000000"), 2, 6)
                Case 3
                    Albaran = "3" & Mid(Format(Albaran, "0000000"), 2, 6)
                Case 5, 6, 7
                    Albaran = "5" & Mid(Format(Albaran, "0000000"), 2, 6)
            End Select
        
            sql = "select count(*) from rbodalbaran where numalbar = " & DBSet(Albaran, "N")
            If TotalRegistros(sql) <> 0 Then
                Mens = "Albar�n ya existe " & Albaran
                sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
    
                conn.Execute sql
            End If
        End If
    End If
    
    
    ' fecha del albaran
    If Not EsFechaOK(Fecha) Then
        Mens = "Fecha incorrecta"
        sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute sql
    End If
    
    ' socio
    sql = RecuperaValorNew(vCadena, ";", 5) ' socio
    If Not EsNumerico(sql) Then
        Mens = "Socio no num�rico: " & sql
        sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute sql
    Else
        Socio = CInt(sql)
        Cifsocio = RecuperaValorNew(vCadena, ";", 7)
        Select Case CInt(txtCodigo(169).Text)
            Case 1 ' anna
                Sql2 = "select count(*) from rsocios where codcoope = 1 "  '(codsocio = " & DBSet(Socio + 1000, "N") & " or codsocio = " & DBSet(Socio + 11000, "N") & ")  "
                Sql2 = Sql2 & " and nifsocio = " & DBSet(Cifsocio, "T")
                If TotalRegistros(Sql2) = 0 Then
                    Mens = RecuperaValorNew(vCadena, ";", 7) & " " & RecuperaValorNew(vCadena, ";", 6)   '& " o cif err�neo " 'Socio + 1000 & " o " & Socio + 11000 & " o cif erroneo"
                    sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
                    
                    conn.Execute sql
                End If
            
            Case 3 ' bolbaite
                Sql2 = "select count(*) from rsocios where codcoope = 3 " 'codsocio = " & DBSet(Socio + 3000, "N")
                Sql2 = Sql2 & " and nifsocio = " & DBSet(Cifsocio, "T")
                If TotalRegistros(Sql2) = 0 Then
'                    Mens = "Socio no existe: " & Socio & " o cif err�neo " 'Socio  + 3000 & " o cif erroneo"
                    Mens = RecuperaValorNew(vCadena, ";", 7) & " " & RecuperaValorNew(vCadena, ";", 6)  '& " o cif err�neo " 'Socio + 1000 & " o " & Socio + 11000 & " o cif erroneo"
                    sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
                    
                    conn.Execute sql
                End If
            
            Case 5, 6, 7 ' navarres
                Sql2 = "select count(*) from rsocios where codcoope in (5,6,7)  "  'codsocio = " & DBSet(Socio + 5000, "N")
                Sql2 = Sql2 & " and nifsocio = " & DBSet(Cifsocio, "T")
                If TotalRegistros(Sql2) = 0 Then
'                    Mens = "Socio no existe: " & Socio & " o cif err�neo " 'Socio + 5000 & " o cif erroneo"
                    Mens = RecuperaValorNew(vCadena, ";", 7) & " " & RecuperaValorNew(vCadena, ";", 6)   '& " o cif err�neo " 'Socio + 1000 & " o " & Socio + 11000 & " o cif erroneo"
                    sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
                    
                    conn.Execute sql
                End If
            
        End Select
    End If
    
    ' articulo
    Dim AAA As String
    AAA = RecuperaValorNew(vCadena, ";", 8)
    
    sql = "select codvarie from variedades where codvarret = " & DBSet(AAA, "T")
    
    Dim RS As ADODB.Recordset
    
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
'        If Not EsNumerico(Sql) Then
'            Mens = "Variedad no num�rica: " & Sql
'            Sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
'                  vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
'
'            conn.Execute Sql
'        Else
'            Sql2 = "select count(*) from variedades where codvarie = " & DBSet(Sql, "N")
'            If TotalRegistros(Sql2) = 0 Then
                Mens = "Variedad no existe: " & AAA
                sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
                
                conn.Execute sql
'            End If
'        End If
    End If
    Set RS = Nothing
    
    ComprobarLinea = True
    Exit Function
    
eComprobarLinea:

End Function


Private Function ProcesarFicheroRetirada(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim sql As String
Dim Sql1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim NomFic As String

    On Error GoTo eProcesarFicheroRetirada

    
    ProcesarFicheroRetirada = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    
    conn.BeginTrans
    
    lblProgres(4).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    pb9.visible = True
    Me.pb9.Max = longitud
    Me.Refresh
    Me.pb9.Value = 0
        
    AlbaranAnterior = 0
        
    B = True
    While Not EOF(NF)
        I = I + 1
        
        Me.pb9.Value = Me.pb9.Value + Len(cad)
        lblProgres(5).Caption = "Linea " & I
        Me.Refresh
        
        cad = cad & ";"
        
        B = InsertarLinea(cad)
        If Not B Then
            ProcesarFicheroRetirada = False
            Exit Function
        End If
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        cad = cad & ";"
        B = InsertarLinea(cad)
    
        If Not B Then
            ProcesarFicheroRetirada = False
            Exit Function
        End If
    End If
    
    
    ProcesarFicheroRetirada = B
    
    pb9.visible = False
    lblProgres(4).Caption = ""
    lblProgres(5).Caption = ""



eProcesarFicheroRetirada:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
    End If

End Function

Private Function InsertarLinea(vCadena As String) As Boolean
Dim Albaran As String
Dim Linea As String
Dim Fecha As String
Dim Socio As String
Dim Variedad As String
Dim cantidad As String
Dim CodIva As String
Dim sql As String
Dim cSocio As Long
Dim Cifsocio As String
Dim Mens As String
Dim Mc As Contadores
Dim Existe As Boolean
Dim devuelve As String

    On Error GoTo EInsertarLinea

    InsertarLinea = True
    
    Albaran = RecuperaValorNew(vCadena, ";", 2)
    Linea = RecuperaValorNew(vCadena, ";", 3)
    Fecha = RecuperaValorNew(vCadena, ";", 4)
    Socio = RecuperaValorNew(vCadena, ";", 5)
    Variedad = RecuperaValorNew(vCadena, ";", 8)
    
    Variedad = DevuelveValor("select codvarie from variedades where codvarret = " & DBSet(Variedad, "T"))

'    Select Case Mid(Variedad, 1, 2)
'        Case "AB"
'            Variedad = "60"
'        Case "AA"
'            Variedad = "61"
'    End Select
    
    cantidad = ImporteSinFormato(RecuperaValorNew(vCadena, ";", 10)) '/ 100
    Cifsocio = RecuperaValorNew(vCadena, ";", 7)
    
'    If Albaran <> AlbaranAnterior Then
'        Set vTipoMov = New CTiposMov
'        If vTipoMov.Leer(CodTipoMov) Then
'            Albaran = vTipoMov.ConseguirContador(CodTipoMov)
'
'            Do
'                devuelve = DevuelveDesdeBDNew(cAgro, "rbodalbaran", "numalbar", "numalbar", CStr(Albaran), "N")
'                If devuelve <> "" Then
'                    'Ya existe el contador incrementarlo
'                    Existe = True
'                    vTipoMov.IncrementarContador (CodTipoMov)
'                    Albaran = vTipoMov.ConseguirContador(CodTipoMov)
'                Else
'                    Existe = False
'                End If
'            Loop Until Not Existe
'            vTipoMov.IncrementarContador (CodTipoMov)
'            Albaran2 = Albaran
'        End If
'        Set vTipoMov = Nothing
'    Else
'        Albaran = Albaran2
'    End If
    
    'el numero del albaran es el que me viene cambiando el primer d�gito
    Select Case CInt(txtCodigo(169))
        Case 1
            Albaran = "1" & Mid(Format(Albaran, "0000000"), 2, 6)
        Case 3
            Albaran = "3" & Mid(Format(Albaran, "0000000"), 2, 6)
        Case 5, 6, 7
            Albaran = "5" & Mid(Format(Albaran, "0000000"), 2, 6)
    End Select
    
    Select Case CInt(txtCodigo(169).Text)
        Case 1 'anna
            sql = "select codsocio from rsocios where codcoope = 1"
            sql = sql & " and nifsocio = " & DBSet(Cifsocio, "T")
            If TotalRegistrosConsulta(sql) <> 0 Then
                cSocio = CLng(DevuelveValor(sql))
            End If
        Case 3 'bolbaite
            sql = "select codsocio from rsocios where codcoope = 3"
            sql = sql & " and nifsocio = " & DBSet(Cifsocio, "T")
            If TotalRegistrosConsulta(sql) <> 0 Then
                cSocio = CLng(DevuelveValor(sql))
            End If
        Case 5, 6, 7
            sql = "select codsocio from rsocios where codcoope in (5,6,7)"
            sql = sql & " and nifsocio = " & DBSet(Cifsocio, "T")
            If TotalRegistrosConsulta(sql) <> 0 Then
                cSocio = CLng(DevuelveValor(sql))
            End If
    End Select
    
    sql = "select count(*) from rbodalbaran where numalbar = " & DBSet(Albaran, "N")
    If TotalRegistros(sql) = 0 Then
        sql = "insert into rbodalbaran (numalbar, fechaalb, codsocio, observac) values ("
        sql = sql & DBSet(Albaran, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(cSocio, "N") & "," & ValorNulo & ")"
            
        conn.Execute sql
    End If
    
    CodIva = DevuelveValor("select codigiva from variedades where codvarie = " & DBSet(Variedad, "N"))
    
    ' insertamos en la tabla de lineas
    sql = "insert into rbodalbaran_variedad (numalbar, numlinea, codvarie, unidades, cantidad, precioar, dtolinea, importel, codigiva) values ("
    sql = sql & DBSet(Albaran, "N") & "," & DBSet(Linea, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(cantidad, "N") & "," & DBSet(cantidad, "N") & ",0,0,0,"
    sql = sql & DBSet(CodIva, "N") & ")"
    
    conn.Execute sql
    
    AlbaranAnterior = Albaran 'RecuperaValor(vCadena, 2)
    
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function

Private Sub CmdAcepTraza_Click()
Dim sql As String
Dim I As Byte
Dim cadWHERE As String
Dim B As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String
Dim nompath As String
Dim Fichero1 As String
Dim Fichero2 As String
Dim cadTabla As String

On Error GoTo eError

'    nomPath = GetFolder("Selecciona directorio")
    If ExistenFicheros Then
        Fichero1 = vParamAplic.PathTraza & "\clasific.txt"
        Fichero2 = vParamAplic.PathTraza & "\entrada.txt"
        
        ' la creacion de las tablas temporales se hace dentro de la transaccion
        
'monica:08052009
    sql = "DROP TABLE IF EXISTS tmpEntrada; "
    conn.Execute sql
    
    sql = "DROP TABLE IF EXISTS tmpClasific; "
    conn.Execute sql
    
    
    sql = "CREATE TEMPORARY TABLE tmpEntrada ("
    sql = sql & " codsocio int, codcampo int, numalbar int, codvarie int, fecalbar date, "
    sql = sql & " horalbar datetime, kilosbru int, kilosnet int, numcajon int) "
    conn.Execute sql
    
    sql = "CREATE TEMPORARY TABLE tmpClasific ("
    sql = sql & " numalbar int, codvarie int, codcalir int, porcenta decimal(5,2)) "
    conn.Execute sql
'08052009
        
        
        conn.BeginTrans
    
        If CargarTablasTemporales(Fichero1, Fichero2) Then
            InicializarVbles
                
                '========= PARAMETROS  =============================
            'A�adir el parametro de Empresa
            CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
    
            If ComprobarErrores() Then
                    cadTabla = "tmpinformes"
                    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                    
                    sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
                    
                    If TotalRegistros(sql) <> 0 Then
                        MsgBox "Hay errores en el Traspaso de Trazabilidad. Debe corregirlos previamente.", vbExclamation
                        cadTitulo = "Errores de Traspaso de TRAZABILIDAD"
                        cadNombreRPT = "rErroresTraza.rpt"
                        LlamarImprimir
                        conn.RollbackTrans
                        Exit Sub
                    Else
                        B = CargarClasificacion()
                    End If
            Else
                B = False
            End If
        Else
            B = False
        End If
    Else
        CmdCancelTraza_Click
        Exit Sub
    End If
    
eError:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
        BorrarArchivo Fichero1
        BorrarArchivo Fichero2
        CmdCancelTraza_Click
    End If
End Sub

Private Sub CmdAcepVtaFruta_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vcad As String

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H SOCIO
    cDesde = Trim(txtCodigo(113).Text)
    cHasta = Trim(txtCodigo(114).Text)
    nDesde = txtNombre(113).Text
    nHasta = txtNombre(114).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    'D/H CLIENTE
    cDesde = Trim(txtCodigo(117).Text)
    cHasta = Trim(txtCodigo(118).Text)
    nDesde = txtNombre(117).Text
    nHasta = txtNombre(118).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codclien}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente=""") Then Exit Sub
    End If
    
    'D/H VARIEDAD
    cDesde = Trim(txtCodigo(115).Text)
    cHasta = Trim(txtCodigo(116).Text)
    nDesde = txtNombre(115).Text
    nHasta = txtNombre(116).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{vtafrutalin.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
    End If

    'D/H Fecha
    cDesde = Trim(txtCodigo(109).Text)
    cHasta = Trim(txtCodigo(110).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        Codigo = "{" & tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
            
    nTabla = "(((vtafrutalin INNER JOIN variedades ON vtafrutalin.codvarie = variedades.codvarie) "
    nTabla = nTabla & " INNER JOIN vtafrutacab ON vtafrutalin.codtipom = vtafrutacab.codtipom and vtafrutalin.numalbar = vtafrutacab.numalbar and vtafrutalin.fecalbar = vtafrutacab.fecalbar) "
    nTabla = nTabla & " LEFT JOIN clientes ON vtafrutacab.codclien = clientes.codclien) "
    nTabla = nTabla & " LEFT JOIN rsocios ON vtafrutacab.codsocio = rsocios.codsocio "
    
    If Check15.Value = 0 Then
        CadParam = CadParam & "pResumen=0|"
    Else
        CadParam = CadParam & "pResumen=1|"  ' imprimir resumen
    End If
    
    If HayRegParaInforme(nTabla, cadSelect) Then
        If CargarTemporalVtaFruta(nTabla, cadSelect) Then
            If Check18.Value = 0 Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                cadNombreRPT = "rInfCompVtaFruta.rpt"
                cadTitulo = "Listado Comprobaci�n Venta Fruta"
                LlamarImprimir
            Else
                Shell App.Path & "\clasificacion.exe /L|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|1|", vbNormalFocus
            End If
        End If
    End If

End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdSubir_Click()
    SubirItemList Me.ListView1
End Sub

Private Sub CmdCancelAgri_Click()
    Unload Me
End Sub

Private Sub cmdCancelBajaSocio_Click()
    Unload Me
End Sub

Private Sub cmdCancelInf_Click()
    Unload Me
End Sub

Private Sub CmdCancelKilosSoc_Click()
    Unload Me
End Sub

Private Sub CmdCancelROPAS_Click()
    Unload Me
End Sub

Private Sub cmdCancelTras_Click()
    Unload Me
End Sub

Private Sub cmdCancelTrasCoop_Click()
    Unload Me
End Sub

Private Sub CmdCancelTraza_Click()
    Unload Me
End Sub


Private Sub Combo1_Click(Index As Integer)

    VisualizarFecha Index

End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 6 Then
        VisualizarFecha Index
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim B As Boolean
    If Index = 3 Then
        If Combo1(Index).ListIndex = 1 Then ' si el tipo de listado es detalle clasificacion
            Combo1(2).ListIndex = 1
            Combo1(2).Enabled = False
            cmdAceptar(3).SetFocus
        Else
            Combo1(2).Enabled = True
            Combo1(2).SetFocus
        End If
    End If
    
    VisualizarFecha Indice
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 10 ' Reimpresion de entradas en bascula
                PonerFoco txtCodigo(12)
            
            Case 11 ' Listado de Entradas de Pesadas
            
            Case 12 ' Listado de Calidades
                PonerFoco txtCodigo(18)
        
            Case 13 ' Listado de Socios por seccion
                PonerFoco txtCodigo(8)
                
            Case 14, 16, 17, 18 '14 = Listado de entradas en bascula
                                '16 = Listado de Entradas clasificadas
                                '17 = Reimpresion de Albaranes de Clasificacion
                                '18 = Informe de Kilos/gastos
                PonerFoco txtCodigo(12)
                
                Select Case OpcionListado
                    Case 17
                        Combo1(10).ListIndex = 0
                    Case 18
                        Combo1(8).ListIndex = 0
                        Combo1(9).ListIndex = 0
                        '[Monica]17/07/2014: a�adido el tipo de socio
                        Combo1(14).ListIndex = 4 ' por defecto todos
                End Select
                
            Case 15 ' Listado de campos huertos
                PonerFoco txtCodigo(2)
                
            Case 19 ' grabacion de fichero agriweb
                PonerFoco txtCodigo(23)
                txtCodigo(27).Text = Format(Year(Now), "0000")
                
            Case 20 ' informe de kilos por producto
                PonerFoco txtCodigo(33)
                
            Case 21 ' traspaso desde el calibrador
                Combo1(6).SetFocus
                
            Case 23 ' baja de socio
                PonerFoco txtCodigo(46)
            
            Case 24 ' traspaso factura cooperativa
                Combo1(7).ListIndex = 0
                PonerFoco txtCodigo(45)
                
            Case 25 ' informe de kilos recolectados por socio/cooperativa
                PonerFoco txtCodigo(54)
            
            Case 26 ' traspaso de ROPAS
                PonerFoco txtCodigo(58)
                txtCodigo(62).Text = Format(Year(Now), "0000")
                txtCodigo(132).Text = Format(Now, "dd/mm/yyyy")
            
            
            Case 28 ' alta masiva de bonificaciones
                PonerFoco txtCodigo(75)
            
            Case 29 ' baja masiva de bonificaciones
                PonerFoco txtCodigo(75)
            
            Case 30 ' Generacion automatica de clasificacion (Picassent)
                PonerFoco txtCodigo(83)
            
            Case 32
                PonerFoco txtCodigo(86)
            
            Case 33
                PonerFoco txtCodigo(100)
                Me.Opcion1(8).Value = True
                Me.Check11.Value = 1
                
            Case 34 ' cambio de socio de un campo
                txtCodigo(106).Text = Format(Now, "dd/mm/yyyy")
                PonerFoco txtCodigo(111)
            
            Case 35 ' informe de comprobacion de venta fruta
                PonerFoco txtCodigo(120)
                ConexionConta
                
            Case 36 ' informe de comprobacion de venta fruta
                PonerFoco txtCodigo(129)
        
            Case 38 ' cambio de numero de factura
                PonerFoco txtCodigo(129)
                FecFacInicial = DevuelveValor("select fecfactu from rfactsoc where " & NumCod)
                txtCodigo(131).Text = Format(CDate(FecFacInicial), "dd/mm/yyyy")
                
            Case 40 ' orden recoleccion
                PonerFoco txtCodigo(147)
                txtCodigo(138).Text = Format(Now, "dd/mm/yyyy")
                
                Check19.Value = 0
                EsReimpresion = False
                
            Case 41 ' informe de ordenes de recoleccion emitidas
                PonerFoco txtCodigo(139)
                
            Case 42 ' informe de socios
                Opcion(8).Value = True
                Opcion(7).Value = True
                PonerFoco txtCodigo(145)
                
            Case 43 ' informe atria
                PonerFoco txtCodigo(153)
            
            Case 44 ' informe de precios
                PonerFoco txtCodigo(155)
            
            Case 45 ' informe de revision de campos
                PonerFoco txtCodigo(163)
        
            Case 46 ' informe de registros fitosanitarios
                PonerFoco txtCodigo(73)
        
            Case 48 ' traspaso de albaranes de retirada para abn
                PonerFoco txtCodigo(169)
        
            Case 50 ' informe de diferencias
                PonerFoco txtCodigo(186)
                Me.Opcion1(14).Value = True
        
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me
    
    ConSubInforme = False

    For h = 0 To 65
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 70 To 78
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 80 To 122
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 125 To 128
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    
    For h = 181 To 181
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    
    For h = 0 To imgAyuda.Count - 1
        imgAyuda(h).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next h
    
    'Ocultar todos los Frames de Formulario
    FrameSociosSeccion.visible = False
    FrameCalidades.visible = False
    FrameCampos.visible = False
    FrameEntradasCampo.visible = False
    FrameGrabacionAgriweb.visible = False
    FrameKilosProducto.visible = False
    FrameTraspasoCalibrador.visible = False
    FrameTrazabilidad.visible = False
    Me.FrameBajaSocios.visible = False
    Me.FrameTraspasoFactCoop.visible = False
    Me.FrameKilosRecolect.visible = False
    Me.FrameTraspasoROPAS.visible = False
    Me.FrameTraspasoAAlmazara.visible = False
    Me.FrameEntradasPesada.visible = False
    Me.FrameBonificaciones.visible = False
    Me.FrameGeneraClasifica.visible = False
    Me.FrameInformeFases.visible = False
    Me.FrameControlDestrio.visible = False
    Me.FrameGastosporConcepto.visible = False
    Me.FrameCambioSocio.visible = False
    Me.FrameVentaFruta.visible = False
    Me.FrameGastosCampos.visible = False
    Me.FrameContabGastos.visible = False
    Me.FrameCambioNroFactura.visible = False
    Me.FrameGeneracionEntradasSIN.visible = False
    Me.FrameOrdenRecoleccion.visible = False
    Me.FrameListOrdenesEmitidas.visible = False
    Me.FrameInformeSocios.visible = False
    Me.FrameInfATRIA.visible = False
    Me.FramePrecios.visible = False
    Me.FrameRevisionCampos.visible = False
    Me.FrameRegFitosanitario.visible = False
    Me.FrameTraspDatosATrazabilidad.visible = False
    Me.FrameTraspasoAlbRetirada.visible = False
    Me.FrameAsignacionGlobalgap.visible = False
    Me.FrameDiferenciaKilos.visible = False
    
    '[Monica]07/06/2013: Zona / bra�al
    Label2(188).Caption = vParamAplic.NomZonaPOZ
    imgBuscar(82).ToolTipText = "Buscar " & vParamAplic.NomZonaPOZ
    imgBuscar(83).ToolTipText = "Buscar " & vParamAplic.NomZonaPOZ
    Me.Opcion1(3).Caption = vParamAplic.NomZonaPOZ
    
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 10 ' Reimpresion de entradas en bascula
        Label3.Caption = "Reimpresi�n de Entradas B�scula"
        FrameEntradaBasculaVisible True, h, w
        indFrame = 1
        tabla = "rentradas"
        Check2.visible = False
        Check2.Enabled = False
        Check5.visible = False
        Check5.Enabled = False
        Check6.visible = False
        Check6.Enabled = False
        '[Monica] 01/10/2009 a�adido el poder detallar las notas
        Check9.visible = False
        Check9.Enabled = False
        Check10.visible = False
        Check10.Enabled = False
        
        FrameTipo.Enabled = False
        FrameTipo.visible = False
        
        FrameRecolectado.Enabled = False
        FrameRecolectado.visible = False
        
        FrameTipoAlbaran.Enabled = False
        FrameTipoAlbaran.visible = False
        
    Case 11 ' Listado de entradas de pesadas
        FrameEntradaPesadaVisible True, h, w
        indFrame = 2
        tabla = "rpesadas"
    
    Case 12 ' Listado de Calidades
        FrameCalidadesVisible True, h, w
        CargarListViewOrden (2)
        indFrame = 2
        tabla = "rcalidad"
    
    Case 13 ' Listado de Socios por Seccion
        FrameSociosSeccionVisible True, h, w
        CargaCombo
        Opcion(0).Value = True
        Opcion(5).Value = True
        CargarListViewOrden (3)
        indFrame = 1
        tabla = "rsocios_seccion"
        
        '[Monica]08/04/2015: tipo de socio por catadau
        Label2(233).visible = (vParamAplic.Cooperativa = 0)
        Combo1(15).Enabled = (vParamAplic.Cooperativa = 0)
        Combo1(15).visible = (vParamAplic.Cooperativa = 0)
        Combo1(15).ListIndex = 0
        
        Check24.Enabled = (Check8.Value = 1)
        
        
        
    Case 14 ' Listado de entradas en bascula
        FrameEntradaBasculaVisible True, h, w
'        Opcion(0).Value = True
        indFrame = 1
        tabla = "rentradas"
        Check2.visible = True
        Check2.Enabled = True
        Check5.visible = False
        Check5.Enabled = False
        Check6.visible = False
        Check6.Enabled = False
        '[Monica] 01/10/2009 a�adido el poder detallar las notas
        Check9.visible = False
        Check9.Enabled = False
        Check10.visible = False
        Check10.Enabled = False
        
        FrameTipo.Enabled = False
        FrameTipo.visible = False
        
        FrameRecolectado.Enabled = False
        FrameRecolectado.visible = False
        
        FrameTipoAlbaran.Enabled = False
        FrameTipoAlbaran.visible = False
        
    Case 15 ' Listado de Campos
        CargaCombo
        Combo1(0).ListIndex = 0
        Combo1(1).ListIndex = 0
        Combo1(11).ListIndex = 0
        FrameCamposVisible True, h, w
        Opcion1(0).Value = True
        tabla = "rcampos"
        
        '[Monica]22/12/2011: solo para picassent pq tiene los informes en hanegadas
        Check16.Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
        Check16.visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
        Opcion1(7).Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
        Opcion1(7).visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
        '[Monica]23/07/2015: informe para Conselleria
        Check23.Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16) And Opcion1(1).Value
        Check23.visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
        imgAyuda(3).Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
        imgAyuda(3).visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
        
        
        Check26.Enabled = (vParamAplic.Cooperativa = 0)
        Check26.visible = (vParamAplic.Cooperativa = 0)
        
        '[Monica]29/09/2016: orden por partida o por socio
        Check27.visible = (Opcion1(1).Value And vParamAplic.Cooperativa = 16)
        Check27.Enabled = (Opcion1(1).Value And vParamAplic.Cooperativa = 16)
        
        
    Case 16, 17, 18 '16= Listado de entradas clasificacion
                    '17= reimpresion de albaranes de clasificacion
                    '18= informe de kilos/gastos
        CargaCombo
        Combo1(2).ListIndex = 0
        Combo1(3).ListIndex = 0
        FrameEntradaBasculaVisible True, h, w
        Opcion(0).Value = True
        indFrame = 1
        Select Case OpcionListado
            Case 16
                tabla = "rclasifica"
                Check2.visible = False
                Check2.Enabled = False
                Check5.visible = False
                Check5.Enabled = False
                Check6.visible = False
                Check6.Enabled = False
               '[Monica] 01/10/2009 a�adido el poder detallar las notas
                Check9.visible = False
                Check9.Enabled = False
                Check10.visible = False
                Check10.Enabled = False
                FrameTipo.Enabled = True
                FrameTipo.visible = True
                
                Label3.Caption = "Informe de Entradas"
                
                FrameRecolectado.Enabled = False
                FrameRecolectado.visible = False
                
                FrameTipoAlbaran.Enabled = False
                FrameTipoAlbaran.visible = False
            
            Case 17, 18
                tabla = "rhisfruta"
                FrameTipo.Enabled = False
                FrameTipo.visible = False
                If OpcionListado = 17 Then
                    Check2.visible = False
                    Check2.Enabled = False
                    Check5.visible = False
                    Check5.Enabled = False
                    Check6.visible = False
                    Check6.Enabled = False
                    '[Monica] 01/10/2009 a�adido el poder detallar las notas
                    Check9.visible = False
                    Check9.Enabled = False
                    Check10.visible = False
                    Check10.Enabled = False

                    Label3.Caption = "Reimpresi�n de Albaranes"
                    
                    FrameRecolectado.Enabled = False
                    FrameRecolectado.visible = False
                
                    FrameTipoAlbaran.Enabled = True
                    FrameTipoAlbaran.visible = True
                    
                Else
                    Check2.visible = True
                    Check2.Enabled = True
                    Check5.visible = True
                    Check5.Enabled = True
                    Check6.visible = True
                    Check6.Enabled = True And (Check5.Value = 1)
                    '[Monica] 01/10/2009 a�adido el poder detallar las notas
                    Check9.visible = True
                    Check9.Enabled = True
                    Check10.visible = True
                    Check10.Enabled = True And (Check5.Value = 1)
                    Label3.Caption = "Informe de Kilos/Gastos"
                    
                    FrameRecolectado.Enabled = True
                    FrameRecolectado.visible = True
                    
                    FrameTipoAlbaran.Enabled = False
                    FrameTipoAlbaran.visible = False
                    
                End If
        End Select
    
    Case 19 ' grabacion de fichero agriweb
        CargaCombo
        Combo1(4).ListIndex = 0
        FrameGrabacionFicheroVisible True, h, w
    
    Case 20 ' informe de kilos por producto
        CargaCombo
        Combo1(5).ListIndex = 0
        FrameKilosProductoVisible True, h, w
        
        '[Monica]20/07/2015: nuevo informe para Mogente
        Check22.visible = (vParamAplic.Cooperativa = 3)
        Check22.Enabled = (vParamAplic.Cooperativa = 3)
        
    Case 21 ' traspaso desde el calibrador
        CargaCombo
        Combo1(6).ListIndex = 0
        FrameTraspasoCalibradorVisible True, h, w
        Pb1.visible = False
        '[Monica]21/04/2016: a�adida la fecha para Castellduc (cooperativa=5)
        FrameNota.visible = vParamAplic.Cooperativa = 5
        FrameNota.Enabled = vParamAplic.Cooperativa = 5
        
        FrameFecha.visible = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 9)
        FrameFecha.Enabled = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 9)
        
    Case 22 ' traspaso de trazabilidad
        FrameTraspasoTrazaVisible True, h, w
        pb2.visible = False
        lblProgres(2).Caption = ""
        lblProgres(3).Caption = ""
        
    Case 23 ' baja de socios
        FrameBajaSociosVisible True, h, w
    
    Case 24 ' traspaso facturas cooperativa (VALSUR)
        CargaCombo
        FrameTraspasoFactCoopVisible True, h, w
    
    Case 25 ' informe de kilos recolectados por socio cooperativa
        CargarListViewTEntrada
        FrameKilosRecolectVisible True, h, w
    
    
    Case 26 ' traspaso ROPAS
        FrameTraspasoROPASVisible True, h, w
    
    
    Case 27 ' traspaso datos a Almazara
        FrameTraspasoAAlmazaraVisible True, h, w
    
    Case 28
        FrameBonificacionesVisible True, h, w
        
        Frame1.visible = True
        Frame1.Enabled = True
    
    Case 29
        FrameBonificacionesVisible True, h, w
        Label15.Caption = "Baja Masiva Bonificaciones"
        
        Frame1.visible = False
        Frame1.Enabled = False
    
    Case 30
        FrameGeneraClasificaVisible True, h, w
        
    Case 31
        CargaCombo
        FrameInformeFasesVisible True, h, w
    
    Case 32
        tabla = "rcontrol"
        FrameControlDestrioVisible True, h, w
    
    Case 33
        tabla = "rhisfruta"
        FrameGastosporConceptoVisible True, h, w
    
    Case 34
        tabla = "rcampos"
        FrameCambioSocioVisible True, h, w
    
    Case 35 ' informe de comprobacion de venta fruta
        tabla = "vtafrutacab"
        FrameVentaFrutaVisible True, h, w
    
    Case 36 ' informe de gastos pendientes de integrar
        tabla = "rcampos"
        FrameGastosCamposVisible True, h, w
        Opcion1(5).Value = True
        
    Case 37 ' Contabilizacion de gastos de campo
        tabla = "rcampos"
        FrameContabGastosCamposVisible True, h, w
    
        ConexionConta
    
    Case 38 ' cambio de nro de factura de socio
        h = FrameCambioNroFactura.Height
        w = FrameCambioNroFactura.Width
        
        PonerFrameVisible FrameCambioNroFactura, True, h, w
    
    Case 39 ' generacion de entradas a partir de las facturas de siniestro
        h = FrameGeneracionEntradasSIN.Height
        w = FrameGeneracionEntradasSIN.Width
        
        PonerFrameVisible FrameGeneracionEntradasSIN, True, h, w
    
    Case 40 ' impresion de ordenes de recoleccion
        h = FrameOrdenRecoleccion.Height
        w = FrameOrdenRecoleccion.Width
        
        PonerFrameVisible FrameOrdenRecoleccion, True, h, w
    
    Case 41 ' Informe de ordenes de recoleccion emitidas
        h = FrameListOrdenesEmitidas.Height
        w = FrameListOrdenesEmitidas.Width
        
        PonerFrameVisible FrameListOrdenesEmitidas, True, h, w
    
    Case 42 ' Informe de Socios/
        h = FrameInformeSocios.Height
        w = FrameInformeSocios.Width
    
        PonerFrameVisible FrameInformeSocios, True, h, w
    
    Case 43 ' Informe de Atria
        h = FrameInfATRIA.Height
        w = FrameInfATRIA.Width
    
        PonerFrameVisible FrameInfATRIA, True, h, w
    
    Case 44 ' Informe de precios
        CargaCombo
    
        h = FramePrecios.Height
        w = FramePrecios.Width
    
        PonerFrameVisible FramePrecios, True, h, w
    
    Case 43 ' Informe de Atria
        h = FrameInfATRIA.Height
        w = FrameInfATRIA.Width
    
        PonerFrameVisible FrameInfATRIA, True, h, w
    
    Case 45 ' Informe de revisiones de campos
        h = FrameRevisionCampos.Height
        w = FrameRevisionCampos.Width
    
        PonerFrameVisible FrameRevisionCampos, True, h, w
    
    Case 46 ' Informe de registros fitosanitarios
        h = FrameRegFitosanitario.Height
        w = FrameRegFitosanitario.Width
    
        PonerFrameVisible FrameRegFitosanitario, True, h, w
        '[Monica]03/05/2016
        txtCodigo(180).Text = "CAMPA�A " & Year(CDate(vParam.FecIniCam)) & "/" & Year(CDate(vParam.FecFinCam))
            
            
    Case 47 ' traspaso datos a trazabilidad (Castelduc)
        FrameTraspDatosATrazabilidadVisible True, h, w
      
    Case 48 ' traspaso albaranes de retirada de cooperativas a ABN
        FrameTraspasoAlbRetiradaVisible True, h, w
        CodTipoMov = "ALB"
    
    Case 49 ' asignacion de globalgap
        h = FrameAsignacionGlobalgap.Height
        w = FrameAsignacionGlobalgap.Width
    
        PonerFrameVisible FrameAsignacionGlobalgap, True, h, w
    
    Case 50 ' informe de diferencias de kilos
        FrameDiferenciaKilosVisible True, h, w
        
    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If OpcionListado = 37 Then
        If Not vSeccion Is Nothing Then
            vSeccion.CerrarConta
            Set vSeccion = Nothing
        End If
    End If
End Sub




Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(Indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de calidades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCampos_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00000000")
End Sub

Private Sub frmCConta_DatoSeleccionado(CadenaSeleccion As String)
'conceptos contables
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") ' codigo de concepto contable
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") ' codigo de cliente
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' nombre del cliente
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") ' codigo de concepto
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion del concepto
End Sub

Private Sub frmCoop_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") ' codigo de cooperativa
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion de la cooperativa
End Sub

Private Sub frmCtaConta_DatoSeleccionado(CadenaSeleccion As String)
' cuentas contables
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") ' codigo de diario
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmDConta_DatoSeleccionado(CadenaSeleccion As String)
' diario contable
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") ' codigo de diario
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") ' codigo de incidencia
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion de la incidencia
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        sql = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
        vSql2 = vSql2 & " and variedades.codvarie in (" & CadenaSeleccion & ")"
    Else
        sql = " {variedades.codvarie} = -1 "
        vSql2 = vSql2 & " and variedades.codvarie is null"
    End If
    If Not AnyadirAFormula(cadSelect, sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMens1_DatoSeleccionado(CadenaSeleccion As String)
Dim sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        sql = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
        
        If Not AnyadirAFormula(cadSelect, sql) Then Exit Sub
    Else
        sql = " {rsocios.codsocio} = -1 "
        
        If Not AnyadirAFormula(cadSelect1, sql) Then Exit Sub
    End If
End Sub

Private Sub frmMens2_DatoSeleccionado(CadenaSeleccion As String)
Dim sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        sql = " {rhisfruta.codcampo} in (" & CadenaSeleccion & ")"
        Sql2 = " {rhisfruta.codcampo} in [" & CadenaSeleccion & "]"
        
        If Not AnyadirAFormula(cadSelect, sql) Then Exit Sub
    Else
        '[Monica]08/04/2014: quito esto para que si es un campo de un coopropietario salga si no tiene campos
        'SQL = " {rhisfruta.codcampo} = -1 "
        'if Not AnyadirAFormula(cadSelect, SQL) Then Exit Sub
    End If
End Sub

Private Sub frmMens3_datoseleccionado(CadenaSeleccion As String)
Dim sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        sql = " {rclasifica.codcampo} in (" & CadenaSeleccion & ")"
        Sql2 = " {rclasifica.codcampo} in [" & CadenaSeleccion & "]"
        
        If Not AnyadirAFormula(cadSelect, sql) Then Exit Sub
    Else
        sql = " {rclasifica.codcampo} = -1 "
        
        If Not AnyadirAFormula(cadSelect, sql) Then Exit Sub
    End If
End Sub

Private Sub frmMens4_DatoSeleccionado(CadenaSeleccion As String)
Dim sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        sql = " {rsocios.codsitua} in (" & CadenaSeleccion & ")"
        Sql2 = " {rsocios.codsitua} in [" & CadenaSeleccion & "]"
    Else
        sql = " {rsocios.codsitua} = -1 "
        Sql2 = " {rsocios.codsitua} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMens5_DatoSeleccionado(CadenaSeleccion As String)
Dim sql As String
Dim Sql2 As String

    HayRegistros = True
    
    If CadenaSeleccion <> "" Then
        sql = " {rcampos.nrocampo} in (" & CadenaSeleccion & ")"
        Sql2 = " {rcampos.nrocampo} in [" & CadenaSeleccion & "]"
    Else
        sql = " {rcampos.nrocampo} = -1 "
        Sql2 = " {rcampos.nrocampo} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMens6_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCodigo(141).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000000")
    End If
End Sub

Private Sub frmMens7_datoseleccionado(CadenaSeleccion As String)
Dim sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        sql = " {rclasifica.contrato} in (" & CadenaSeleccion & ")"
        Sql2 = " {rclasifica.contrato} in [" & CadenaSeleccion & "]"
        
        If Not AnyadirAFormula(cadSelect, sql) Then Exit Sub
    Else
        sql = " {rclasifica.contrato} = '' "
        
        If Not AnyadirAFormula(cadSelect, sql) Then Exit Sub
    End If
End Sub

Private Sub frmMens8_datoseleccionado(CadenaSeleccion As String)
Dim sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Contratos = CadenaSeleccion
    Else
        Contratos = ""
    End If
End Sub


Private Sub frmPar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmProd_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
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

Private Sub frmSitu_DatoSeleccionado(CadenaSeleccion As String)
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


Private Sub frmCapa_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmZon_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "El informe saca los datos de los recintos." & vbCrLf & _
                      "Las �nicas hectareas que hay en recintos son sigpac y catastro." & vbCrLf & vbCrLf & _
                      "" & vbCrLf
                      
                      
        Case 1
           ' "____________________________________________________________"
            vCadena = "En caso de reimpresi�n de ordenes de recolecci�n s�lo se tiene " & vbCrLf & _
                      "en cuenta el nro de orden a reimprimir." & vbCrLf & vbCrLf & _
                      "" & vbCrLf
                      
        Case 2
           ' "____________________________________________________________"
            vCadena = "Tipo de Socio se corresponde con el Tipo de Productor" & vbCrLf & _
                      "de la ficha del socio." & vbCrLf & vbCrLf & _
                      "" & vbCrLf
                      
        Case 3
           ' "____________________________________________________________"
            vCadena = "El informe de campos para Conselleria, s�lo est� activo" & vbCrLf & _
                      "si est� ordenado por clase/variedad y no est� marcado " & vbCrLf & _
                      "imprimir resumen. " & vbCrLf & _
                      "En ese caso saca el DNI del socio y el t�rmino municipal" & vbCrLf & _
                      "en lugar de la partida y la zona. " & vbCrLf & vbCrLf & _
                      "" & vbCrLf
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripci�n de Ayuda"
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1, 6, 7, 25, 26, 35, 36  ' Clase
            AbrirFrmClase (Index)
        
        Case 31, 32 'clase
            AbrirFrmClase (Index + 25)
        
        Case 55, 56 'clase
            AbrirFrmClase (Index + 26)
        
        Case 63, 64 'clase
            AbrirFrmClase (Index + 41)
        
        Case 4, 5 ' Situacion de campo
            AbrirFrmSituacionCampo (Index)
            
        Case 8, 9 'SECCION
            AbrirFrmSeccion (Index)
        
        Case 2, 3, 10, 11, 12, 13, 23, 24, 33, 34 'SOCIOS
            AbrirFrmSocios (Index)
        
        Case 29, 30 'socios
            AbrirFrmSocios (Index + 19)
        
        Case 27, 28 'socios
            AbrirFrmSocios (Index + 27)
        
        Case 42, 43 'socios
            AbrirFrmSocios (Index + 16)
        
        Case 47 'socios
            AbrirFrmSocios (Index + 53)
        Case 50 'socios
            AbrirFrmSocios (Index + 51)
            
        Case 53, 54 'socios
            AbrirFrmSocios (Index + 33)
            
        Case 73, 74 'socios
            AbrirFrmSocios (Index + 40)
        
        Case 14, 15, 18, 19 'VARIEDADES
            AbrirFrmVariedad (Index)
    
        Case 51, 52 'VARIEDADES
            AbrirFrmVariedad (Index + 33)
        
        Case 75, 76 'VARIEDADES
            AbrirFrmVariedad (Index + 40)
            
        Case 61, 62 'VARIEDADES
            AbrirFrmVariedad (Index + 41)
        
        Case 77, 78 'CLIENTES
            indCodigo = Index + 40
            
            Set frmCli = New frmBasico
            
            AyudaClienteCom frmCli, Index + 40
            
            Set frmCli = Nothing
            
        Case 20 ' cooperativa
            AbrirFrmCooperativa (45)
            
        Case 16, 17 'CALIDADES
            AbrirFrmCalidad (Index)
            
        Case 22, 37, 38 'Producto
            AbrirFrmProducto (Index)
            
        Case 39 'Socios
             AbrirFrmSocios (Index + 44)
            
        Case 40, 41 'Producto
            AbrirFrmProducto (Index + 10)
            
        Case 44, 45 'Producto
            AbrirFrmProducto (Index + 16)
        
        Case 46 ' situacion de socio
            AbrirFrmSituacion (Index)
        
        Case 48, 49 'socios
            AbrirFrmSocios (Index + 16)
            
        Case 59, 60 'capataz (responsable campo)
            AbrirFrmCapataz (Index + 33)
            
        Case 57, 58 'partidas
            AbrirFrmPartidas (Index + 37)
            
        Case 65 ' Concepto
            AbrirFrmConceptos (Index + 31)
        Case 70 ' Concepto
            AbrirFrmConceptos (Index + 27)
            
        Case 71 'socios
            AbrirFrmSocios (111)
        Case 72 'incidencia
            AbrirFrmIncidencias (107)

        ' informe de gastos de pendientes
        Case 90, 91 'socio
            AbrirFrmSocios (Index + 30)
            
        Case 94, 95 ' concepto de gastos
            AbrirFrmConceptos (Index + 30)
            
        Case 92, 93 ' campos
            AbrirFrmCampos (Index + 30)
        
        ' contabilizacion de gastos de campos
        Case 79 ' diario
            AbrirFrmDiariosConta (Index + 29)
        Case 80 ' concepto conta
            AbrirFrmConceptosConta (Index + 32)
        Case 81 ' cta de contrapartida
            AbrirFrmCuentasConta (Index + 47)
           
        Case 82, 83
            AbrirFrmZonas (Index + 51)
           
        'impresion de ordenes de recoleccion
        Case 84 'capataz
            AbrirFrmCapataz (Index + 63)
        
        Case 85 'variedad
            AbrirFrmVariedad (Index + 64)
           
        Case 86 'partida
            AbrirFrmPartidas (Index + 56)
           
        Case 89 'nro de ordenes de recoleccion
            Set frmMens6 = New frmMensajes
            
            frmMens6.OpcionMensaje = 52
            frmMens6.cadWHERE = ""
            frmMens6.Show vbModal
            
            Set frmMens6 = Nothing
                    
        Case 87, 88 'VARIEDADES
            AbrirFrmVariedad (Index + 56)
        
        Case 96, 97 'socio
            AbrirFrmSocios (Index + 49)
        
        ' Informe de ATRIA
        Case 104, 105 'socio
            AbrirFrmSocios (Index + 49)
        Case 98 'Producto
            AbrirFrmProducto (Index + 50)
        Case 99 'Producto
            AbrirFrmProducto (Index + 51)
        Case 100, 101 'VARIEDADES
            AbrirFrmVariedad (Index + 51)
        
        Case 102, 103 ' variedades
            AbrirFrmVariedad (Index + 53)
        
        ' Informe de Revision Campos
        Case 110, 111 'socio
            AbrirFrmSocios (Index + 53)
        Case 108, 109 'VARIEDADES
            AbrirFrmVariedad (Index + 53)
                    
        ' Informe de registro de Fitosanitarios
        Case 118, 119 'socio
            AbrirFrmSocios (Index + 55)
        Case 120, 121 'Producto
            AbrirFrmProducto (Index + 55)
        Case 112, 113 'partidas
            AbrirFrmPartidas (Index + 55)
        Case 106, 107 'termino municipal
            AbrirFrmPueblos (Index + 53)
            
        ' traspaso de datos a trazabilidad
        Case 114, 115 'socio
            AbrirFrmSocios (Index + 57)
        Case 116, 117 'VARIEDADES
            AbrirFrmVariedad (Index + 61)
        
        Case 122 'cooperativa
            AbrirFrmCooperativa (169)
                    
        ' situacion de baja de campo (dentro de baja de socio)
        Case 181 ' situacion de campo
            AbrirFrmSituacionCampo (181)
                    
                    
        ' informe de diferencia de kilos
        Case 125, 126 'socios
            AbrirFrmSocios (Index + 61)
            
        Case 127, 128 'clase
            AbrirFrmClase (Index + 61)
        
                    
                    
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim B As Boolean
Dim TotalArray As Integer
    If Index < 2 Then
        'En el listview3
        B = Index = 1
        For TotalArray = 1 To ListView1(0).ListItems.Count
            ListView1(0).ListItems(TotalArray).Checked = B
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
    End If
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

    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    Select Case Index
        Case 0, 1
            Indice = Index + 6
        Case 11
            Indice = 30
        Case 2, 3
            Indice = Index + 37
        Case 5
            Indice = 47
        Case 7, 8
            Indice = Index + 45
        Case 9
            Indice = 63
        Case 10
            Indice = 72
        Case 12
            Indice = 73
        Case 14, 15
            Indice = Index + 74
        Case 16, 17
            Indice = Index + 82
        Case 20
            Indice = 106
        Case 18, 19
            Indice = Index + 91
        Case 21, 22
            Indice = Index + 105
        Case 23
            Indice = 131
        Case 24
            Indice = 132
        Case 25, 26
            Indice = Index + 111
        Case 27
            Indice = 135
        Case 28
            Indice = 138
        Case 29, 30
            Indice = Index + 110
        Case 31, 32
            Indice = Index + 126
        Case 33, 34
            Indice = Index + 132
            
        Case 35, 36
            Indice = Index + 149
    End Select


    imgFec(0).Tag = Indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Indice).Text <> "" Then frmC.NovaData = txtCodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(Indice) '<===
    ' ********************************************

End Sub


Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Opcion_Click(Index As Integer)
    If Index = 0 Then
        If vParamAplic.Cooperativa = 0 And Opcion(0).Value Then
            Label2(233).visible = True
            Combo1(15).Enabled = True
            Combo1(15).visible = True
            Combo1(15).ListIndex = 0
        End If
    Else
        Label2(233).visible = False
        Combo1(15).Enabled = False
        Combo1(15).visible = False
    End If
End Sub

Private Sub Opcion1_Click(Index As Integer)
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        Check23.Enabled = (Opcion1(1).Value)
    End If
    
    '[Monica]03/06/2016: si es por socio se puede pedir que salte p�gina
    If vParamAplic.Cooperativa = 0 Then
        Check26.Enabled = (Opcion1(0).Value)
        If Opcion1(0).Value = 0 Then Check26.Value = 0
    End If
    
    '[Monica]29/09/2016: para el caso de de que sea por clase/variedad
    Check27.visible = (Opcion1(1).Value And vParamAplic.Cooperativa = 16)
    Check27.Enabled = (Opcion1(1).Value And vParamAplic.Cooperativa = 16)
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
            Case 0: KEYBusqueda KeyAscii, 0 'clase desde
            Case 1: KEYBusqueda KeyAscii, 1 'clase hasta
            
            Case 2: KEYBusqueda KeyAscii, 2 'socio desde
            Case 3: KEYBusqueda KeyAscii, 3 'socio hasta
            Case 4: KEYBusqueda KeyAscii, 4 'situacion desde
            Case 5: KEYBusqueda KeyAscii, 5 'situacion hasta
            Case 6: KEYFecha KeyAscii, 0 'fecha entrada
            Case 7: KEYFecha KeyAscii, 1 'fecha entrada
            Case 8: KEYBusqueda KeyAscii, 8 'seccion desde
            Case 9: KEYBusqueda KeyAscii, 9 'seccion hasta
            Case 10: KEYBusqueda KeyAscii, 10 'socio desde
            Case 11: KEYBusqueda KeyAscii, 11 'socio hasta
            Case 12: KEYBusqueda KeyAscii, 12 'socio desde
            Case 13: KEYBusqueda KeyAscii, 13 'socio hasta
            Case 14: KEYBusqueda KeyAscii, 14 'variedad desde
            Case 15: KEYBusqueda KeyAscii, 15 'variedad hasta
            Case 16: KEYBusqueda KeyAscii, 16 'calidad desde
            Case 17: KEYBusqueda KeyAscii, 17 'calidad desde
            Case 18: KEYBusqueda KeyAscii, 18 'variedad desde
            Case 19: KEYBusqueda KeyAscii, 19 'variedad desde
            Case 20: KEYBusqueda KeyAscii, 6 'clase desde
            Case 21: KEYBusqueda KeyAscii, 7 'clase hasta
            Case 22: KEYBusqueda KeyAscii, 22 'producto
            Case 23: KEYBusqueda KeyAscii, 23 'socio desde
            Case 24: KEYBusqueda KeyAscii, 24 'sosio hasta
            Case 25: KEYBusqueda KeyAscii, 25 'clase desde
            Case 26: KEYBusqueda KeyAscii, 26 'clase hasta
            Case 30: KEYFecha KeyAscii, 11 'fecha formalizacion
            
            Case 33: KEYBusqueda KeyAscii, 33 'socio desde
            Case 34: KEYBusqueda KeyAscii, 34 'socio hasta
            Case 35: KEYBusqueda KeyAscii, 35 'clase desde
            Case 36: KEYBusqueda KeyAscii, 36 'clase hasta
            Case 37: KEYBusqueda KeyAscii, 37 'producto desde
            Case 38: KEYBusqueda KeyAscii, 38 'producto hasta
            Case 39: KEYFecha KeyAscii, 2 'fecha desde
            Case 40: KEYFecha KeyAscii, 3 'fecha hasta
            
            Case 43: KEYFecha KeyAscii, 4 'fecha desde
            Case 44: KEYFecha KeyAscii, 6 'fecha hasta
            
            Case 45: KEYBusqueda KeyAscii, 20 ' cooperativa
            
            Case 46: KEYBusqueda KeyAscii, 46 'situacion de baja de socio
            Case 47: KEYFecha KeyAscii, 5 'fecha de baja de socio
            '[Monica]21/09/2016: situacion de baja de socio
            Case 181: KEYBusqueda KeyAscii, 181 'situacion de baja de campo
            
            
            Case 48: KEYBusqueda KeyAscii, 29 'socio desde
            Case 49: KEYBusqueda KeyAscii, 30 'socio hasta
            
            Case 50: KEYBusqueda KeyAscii, 40 ' producto desde
            Case 51: KEYBusqueda KeyAscii, 41 ' producto hasta
            
            Case 54: KEYBusqueda KeyAscii, 27 ' socio desde
            Case 55: KEYBusqueda KeyAscii, 28 ' socio hasta
            
            Case 56: KEYBusqueda KeyAscii, 31 ' clase desde
            Case 57: KEYBusqueda KeyAscii, 32 ' clase hasta

            Case 52: KEYFecha KeyAscii, 7 'fecha desde
            Case 53: KEYFecha KeyAscii, 8 'fecha hasta
            Case 63: KEYFecha KeyAscii, 9 'fecha de calibrado
            Case 64: KEYBusqueda KeyAscii, 48 'socio desde
            Case 65: KEYBusqueda KeyAscii, 49 'socio hasta
            
            Case 66: KEYBusqueda KeyAscii, 66 ' clase desde
            Case 67: KEYBusqueda KeyAscii, 67 ' clase hasta

            Case 68: KEYBusqueda KeyAscii, 68 'variedad desde
            Case 69: KEYBusqueda KeyAscii, 69 'variedad desde

            Case 72: KEYFecha KeyAscii, 10 'fecha desde
            Case 73: KEYFecha KeyAscii, 12 'fecha hasta
            
            Case 75: KEYBusqueda KeyAscii, 21 'variedad
            Case 74: KEYFecha KeyAscii, 13 'fecha
            
            Case 83: KEYBusqueda KeyAscii, 39 'socio
            
            Case 86: KEYBusqueda KeyAscii, 53 'socio desde
            Case 87: KEYBusqueda KeyAscii, 54 'socio hasta
            
            Case 81: KEYBusqueda KeyAscii, 55 ' clase desde
            Case 82: KEYBusqueda KeyAscii, 56 ' clase hasta
            
            Case 84: KEYBusqueda KeyAscii, 51 'variedad desde
            Case 85: KEYBusqueda KeyAscii, 52 'variedad hasta
            
            Case 88: KEYFecha KeyAscii, 14 'fecha desde
            Case 89: KEYFecha KeyAscii, 15 'fecha hasta
            
            Case 92: KEYBusqueda KeyAscii, 59 'capataz desde
            Case 93: KEYBusqueda KeyAscii, 60 'capataz hasta
            Case 94: KEYBusqueda KeyAscii, 57 'partida desde
            Case 95: KEYBusqueda KeyAscii, 58 'partida hasta
            
            'listado de gastos por conceptos
            Case 100: KEYBusqueda KeyAscii, 47 'socio desde
            Case 101: KEYBusqueda KeyAscii, 50 'socio hasta
            
            Case 104: KEYBusqueda KeyAscii, 63 ' clase desde
            Case 105: KEYBusqueda KeyAscii, 64 ' clase hasta
            
            Case 102: KEYBusqueda KeyAscii, 61 'variedad desde
            Case 103: KEYBusqueda KeyAscii, 62 'variedad hasta
            
            Case 98: KEYFecha KeyAscii, 16 'fecha desde
            Case 99: KEYFecha KeyAscii, 17 'fecha hasta
            
            Case 96: KEYBusqueda KeyAscii, 65 'concepto desde
            Case 97: KEYBusqueda KeyAscii, 70 'concepto hasta
            
            ' cambio del socio del campo
            Case 111: KEYBusqueda KeyAscii, 71 'socio
            Case 106: KEYFecha KeyAscii, 20 'fecha
            Case 107: KEYFecha KeyAscii, 72 'codigo de incidencia
        
            ' listado de comprobacion de venta fruta
            Case 113: KEYBusqueda KeyAscii, 73 'socio desde
            Case 114: KEYBusqueda KeyAscii, 74 'socio hasta
            Case 117: KEYBusqueda KeyAscii, 77 'cliente desde
            Case 118: KEYBusqueda KeyAscii, 78 'cliente hasta
            Case 115: KEYBusqueda KeyAscii, 75 'variedad desde
            Case 116: KEYBusqueda KeyAscii, 76 'variedad hasta
            Case 109: KEYFecha KeyAscii, 18 'fecha desde
            Case 110: KEYFecha KeyAscii, 19 'fecha hasta
            
            ' listado de gastos pendientes de integrar en la contabilidad
            Case 120: KEYBusqueda KeyAscii, 90 'socio desde
            Case 121: KEYBusqueda KeyAscii, 91 'socio hasta
            Case 124: KEYBusqueda KeyAscii, 94 'concepto desde
            Case 125: KEYBusqueda KeyAscii, 95 'concepto hasta
            
            Case 112: KEYBusqueda KeyAscii, 80 'concepto hasta
            Case 128: KEYBusqueda KeyAscii, 81 'cuenta contrapartida
            
            Case 131: KEYFecha KeyAscii, 23 ' nueva fecha de factura socio
        
            Case 132: KEYFecha KeyAscii, 24 'fecha
            
            Case 133: KEYBusqueda KeyAscii, 82 ' zona desde
            Case 134: KEYBusqueda KeyAscii, 83 ' zona hasta
        
            Case 136: KEYFecha KeyAscii, 25 'fecha desde
            Case 137: KEYFecha KeyAscii, 26 'fecha hasta
            
            Case 135: KEYFecha KeyAscii, 27 'fecha desde
        
            'Impresion de Ordenes de Recoleccion
            Case 147: KEYBusqueda KeyAscii, 84 ' responsable
            Case 149: KEYBusqueda KeyAscii, 85 ' variedad
            Case 142: KEYBusqueda KeyAscii, 86 ' partida
            Case 138: KEYFecha KeyAscii, 28 'fecha de recogida
        
            'Informes de ordenes de recoleccion impresas
            Case 139: KEYFecha KeyAscii, 29 'fecha desde
            Case 140: KEYFecha KeyAscii, 30 'fecha hasta
            Case 143: KEYBusqueda KeyAscii, 87 'variedad desde
            Case 144: KEYBusqueda KeyAscii, 88 'variedad hasta
            
            'Informes de socios
            Case 145: KEYBusqueda KeyAscii, 96 'socio hasta
            Case 146: KEYBusqueda KeyAscii, 97 'socio desde
        
            'Informe de miembros ATRIA
            Case 153: KEYBusqueda KeyAscii, 104 'socio desde
            Case 154: KEYBusqueda KeyAscii, 105 'socio hasta
            Case 148: KEYBusqueda KeyAscii, 98  'producto desde
            Case 150: KEYBusqueda KeyAscii, 99  'producto hasta
            Case 151: KEYBusqueda KeyAscii, 100 'variedad desde
            Case 152: KEYBusqueda KeyAscii, 101 'variedad hasta
        
            ' Informe de precios
            Case 155: KEYBusqueda KeyAscii, 102 'variedad desde
            Case 156: KEYBusqueda KeyAscii, 103 'variedad hasta
        
            'Informe de revision de campos
            Case 163: KEYBusqueda KeyAscii, 110 'socio desde
            Case 164: KEYBusqueda KeyAscii, 111 'socio hasta
            Case 161: KEYBusqueda KeyAscii, 108 'variedad desde
            Case 162: KEYBusqueda KeyAscii, 109 'variedad hasta
            Case 165: KEYFecha KeyAscii, 33 'fecha desde
            Case 166: KEYFecha KeyAscii, 34 'fecha hasta
        
            'Informe de registro de aplicacion de fitosanitarios
            Case 173: KEYBusqueda KeyAscii, 118 'socio desde
            Case 174: KEYBusqueda KeyAscii, 119 'socio hasta
            Case 175: KEYBusqueda KeyAscii, 120  'producto desde
            Case 176: KEYBusqueda KeyAscii, 121  'producto hasta
            Case 167: KEYBusqueda KeyAscii, 112 'partida desde
            Case 168: KEYBusqueda KeyAscii, 113 'partida hasta
            Case 159: KEYBusqueda KeyAscii, 106 'poblacion desde
            Case 160: KEYBusqueda KeyAscii, 107 'poblacion hasta
            
            'traspaso de datos a trazabilidad
            Case 171: KEYBusqueda KeyAscii, 114 'socio desde
            Case 172: KEYBusqueda KeyAscii, 115 'socio hasta
            Case 177: KEYBusqueda KeyAscii, 116 'variedad desde
            Case 178: KEYBusqueda KeyAscii, 117 'variedad hasta
        
            Case 169: KEYBusqueda KeyAscii, 122 ' cooperativa
        
            ' listado de diferencia de kilos
            Case 186: KEYBusqueda KeyAscii, 125 'socio desde
            Case 187: KEYBusqueda KeyAscii, 126 'socio hasta
            
            Case 188: KEYBusqueda KeyAscii, 127 ' clase desde
            Case 189: KEYBusqueda KeyAscii, 128 ' clase hasta
            
            Case 184: KEYFecha KeyAscii, 35 'fecha desde
            Case 185: KEYFecha KeyAscii, 36 'fecha hasta
        
        
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
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1, 20, 21, 25, 26, 35, 36, 56, 57, 66, 67, 81, 82, 104, 105, 188, 189 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
    
        Case 8, 9 'SECCIONES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rseccion", "nomsecci", "codsecci", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
        Case 2, 3, 10, 11, 12, 13, 23, 24, 33, 34, 48, 49, 54, 55, 58, 59, 64, 65, 83, 86, 87, 100, 101, 113, 114, 120, 121, 153, 154, 163, 164, 173, 174, 171, 172, 186, 187 'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
        Case 4, 5, 181 'SITUACION
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsituacioncampo", "nomsitua", "codsitua", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
        Case 6, 7, 30, 39, 40, 47, 43, 44, 52, 53, 63, 72, 73, 74, 88, 89, 98, 99, 132, 136, 137, 135, 138, 139, 140, 157, 158, 165, 166, 184, 185 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 16, 17 'CALIDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rcalidad", "nomcalid", "codcalid", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
        Case 14, 15, 18, 19, 68, 69, 75, 84, 85, 102, 103, 115, 116, 149, 143, 144, 151, 152, 155, 156, 161, 162, 177, 178 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 22, 37, 38, 50, 51, 60, 61, 148, 150, 175, 176 'PRODUCTOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "productos", "nomprodu", "codprodu", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
        Case 27, 29 ' datos de agroweb
            txtCodigo(Index).Text = Format(txtCodigo(Index).Text, FormatoCampo(txtCodigo(Index)))
            
        Case 31
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index) = Format(TransformaPuntosComas(txtCodigo(Index).Text), "#,##0.00")
            
        Case 32 ' datos de agroweb
            PonerFormatoDecimal txtCodigo(Index), 4
    
        Case 45, 169 ' cooperativa
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rcoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
        
        Case 46 'SITUACION de socio
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsituacion", "nomsitua", "codsitua", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
    
        Case 62 ' Ejercicio
            txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
    
        Case 70, 71 ' nro de pesada
            txtCodigo(Index).Text = Format(txtCodigo(Index).Text, FormatoCampo(txtCodigo(Index)))
            
        ' Alta masiva de bonificaciones
        Case 76 ' nro de dias
            PonerFormatoEntero txtCodigo(Index)
            
        Case 77 ' porcentaje de aumento
            PonerFormatoDecimal txtCodigo(Index), 4
        
        Case 78 ' indice de correccion
            PonerFormatoDecimal txtCodigo(Index), 4
    
        Case 79 ' porcentaje de destrio
            PonerFormatoDecimal txtCodigo(Index), 4
    
        Case 92, 93, 147 'CAPATAZ
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rcapataz", "nomcapat", "codcapat", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
    
        Case 94, 95, 142, 167, 168 'Partidas
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rpartida", "nomparti", "codparti", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
    
        Case 159, 160 'Poblacion
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rpueblos", "despobla", "codpobla", "T")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = txtCodigo(Index).Text
    
    
        Case 16, 17, 124, 125 'CONCEPTOS DE GASTOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rconcepgasto", "nomgasto", "codgasto", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
    
        Case 106 ' fecha de cambio
            PonerFormatoFecha txtCodigo(Index)
            
        Case 107 ' Incidencia
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rincidencia", "nomincid", "codincid", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
        
        Case 111 ' socio
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
                    
        Case 117, 118 ' clientes
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
    
        Case 109, 110, 126, 127  'Fechas
            PonerFormatoFecha txtCodigo(Index), True
    
        Case 112 ' concepto contable
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "conceptos", "nomconce", "codconce", "N", cConta)
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
        Case 128 ' Cuentas contables
            If vSeccion Is Nothing Then Exit Sub
        
            If txtCodigo(Index).Text <> "" Then
                txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2)
                If txtNombre(Index).Text = "" Then PonerFoco txtCodigo(Index)
            Else
                MsgBox "N�mero de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
      
        Case 131 ' fecha de cambio
            PonerFormatoFecha txtCodigo(Index)
    
        Case 133, 134 ' zonas
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rzonas", "nomzonas", "codzonas", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
    
        Case 141 ' nro de orden de recoleccion
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            
        Case 145, 146 ' socio
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
    
    
        Case 170, 179
            If Not IsNumeric(txtCodigo(Index)) Then
                MsgBox "El n�mero de nota ha de ser num�rico. Reintroduzca.", vbExclamation
                PonerFoco txtCodigo(Index)
            End If
    
    End Select
End Sub

Private Sub FrameCalidadesVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de clientes
    Me.FrameCalidades.visible = visible
    If visible = True Then
        Me.FrameCalidades.Top = -90
        Me.FrameCalidades.Left = 0
        Me.FrameCalidades.Height = 4820
        Me.FrameCalidades.Width = 8600
        w = Me.FrameCalidades.Width
        h = Me.FrameCalidades.Height
    End If
End Sub

Private Sub FrameSociosSeccionVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameSociosSeccion.visible = visible
    If visible = True Then
        Me.FrameSociosSeccion.Top = -90
        Me.FrameSociosSeccion.Left = 0
        Me.FrameSociosSeccion.Height = 5655
        Me.FrameSociosSeccion.Width = 6600
        w = Me.FrameSociosSeccion.Width
        h = Me.FrameSociosSeccion.Height
    End If
End Sub

Private Sub FrameEntradaBasculaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameEntradasCampo.visible = visible
    If visible = True Then
        Me.FrameEntradasCampo.Top = -90
        Me.FrameEntradasCampo.Left = 0
        Me.FrameEntradasCampo.Height = 6480
        Me.FrameEntradasCampo.Width = 6615
        w = Me.FrameEntradasCampo.Width
        h = Me.FrameEntradasCampo.Height
    End If
End Sub

Private Sub FrameEntradaPesadaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameEntradasPesada.visible = visible
    If visible = True Then
        Me.FrameEntradasPesada.Top = -90
        Me.FrameEntradasPesada.Left = 0
        Me.FrameEntradasPesada.Height = 5715
        Me.FrameEntradasPesada.Width = 6615
        w = Me.FrameEntradasPesada.Width
        h = Me.FrameEntradasPesada.Height
    End If
End Sub

Private Sub FrameCamposVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameCampos.visible = visible
    If visible = True Then
        Me.FrameCampos.Top = -90
        Me.FrameCampos.Left = 0
        Me.FrameCampos.Height = 9795 '9465
        Me.FrameCampos.Width = 6600
        w = Me.FrameCampos.Width
        h = Me.FrameCampos.Height
        
    End If
End Sub

Private Sub FrameGrabacionFicheroVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameGrabacionAgriweb.visible = visible
    If visible = True Then
        Me.FrameGrabacionAgriweb.Top = -90
        Me.FrameGrabacionAgriweb.Left = 0
        Me.FrameGrabacionAgriweb.Height = 6975
        Me.FrameGrabacionAgriweb.Width = 6675
        w = Me.FrameGrabacionAgriweb.Width
        h = Me.FrameGrabacionAgriweb.Height
    End If
End Sub

Private Sub FrameKilosProductoVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameKilosProducto.visible = visible
    If visible = True Then
        Me.FrameKilosProducto.Top = -90
        Me.FrameKilosProducto.Left = 0
        Me.FrameKilosProducto.Height = 6480
        Me.FrameKilosProducto.Width = 6615
        w = Me.FrameKilosProducto.Width
        h = Me.FrameKilosProducto.Height
    End If
End Sub

Private Sub FrameDiferenciaKilosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameDiferenciaKilos.visible = visible
    If visible = True Then
        Me.FrameDiferenciaKilos.Top = -90
        Me.FrameDiferenciaKilos.Left = 0
        Me.FrameDiferenciaKilos.Height = 5670
        Me.FrameDiferenciaKilos.Width = 6615
        w = Me.FrameDiferenciaKilos.Width
        h = Me.FrameDiferenciaKilos.Height
    End If
End Sub




Private Sub FrameTraspasoCalibradorVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameTraspasoCalibrador.visible = visible
    If visible = True Then
        Me.FrameTraspasoCalibrador.Top = -90
        Me.FrameTraspasoCalibrador.Left = 0
        Me.FrameTraspasoCalibrador.Height = 4665
        Me.FrameTraspasoCalibrador.Width = 6555
        w = Me.FrameTraspasoCalibrador.Width
        h = Me.FrameTraspasoCalibrador.Height
    End If
End Sub


Private Sub FrameTraspasoTrazaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el trapaso de trazabilidad
    Me.FrameTrazabilidad.visible = visible
    If visible = True Then
        Me.FrameTrazabilidad.Top = -90
        Me.FrameTrazabilidad.Left = 0
        Me.FrameTrazabilidad.Height = 4665
        Me.FrameTrazabilidad.Width = 6555
        w = Me.FrameTrazabilidad.Width
        h = Me.FrameTrazabilidad.Height
    End If
End Sub

Private Sub FrameBajaSociosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameBajaSocios.visible = visible
    If visible = True Then
        Me.FrameBajaSocios.Top = -90
        Me.FrameBajaSocios.Left = 0
        Me.FrameBajaSocios.Height = 4050
        Me.FrameBajaSocios.Width = 7785
        w = Me.FrameBajaSocios.Width
        h = Me.FrameBajaSocios.Height
    End If
End Sub


Private Sub FrameTraspasoFactCoopVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameTraspasoFactCoop.visible = visible
    If visible = True Then
        Me.FrameTraspasoFactCoop.Top = -90
        Me.FrameTraspasoFactCoop.Left = 0
        Me.FrameTraspasoFactCoop.Height = 5490
        Me.FrameTraspasoFactCoop.Width = 6615
        w = Me.FrameTraspasoFactCoop.Width
        h = Me.FrameTraspasoFactCoop.Height
    End If
End Sub


Private Sub FrameTraspasoROPASVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameTraspasoROPAS.visible = visible
    If visible = True Then
        Me.FrameTraspasoROPAS.Top = -90
        Me.FrameTraspasoROPAS.Left = 0
        Me.FrameTraspasoROPAS.Height = 5490
        Me.FrameTraspasoROPAS.Width = 6615
        w = Me.FrameTraspasoROPAS.Width
        h = Me.FrameTraspasoROPAS.Height
    End If
End Sub


Private Sub FrameTraspDatosATrazabilidadVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameTraspDatosATrazabilidad.visible = visible
    If visible = True Then
        Me.FrameTraspDatosATrazabilidad.Top = -90
        Me.FrameTraspDatosATrazabilidad.Left = 0
        Me.FrameTraspDatosATrazabilidad.Height = 4320
        Me.FrameTraspDatosATrazabilidad.Width = 6615
        w = Me.FrameTraspDatosATrazabilidad.Width
        h = Me.FrameTraspDatosATrazabilidad.Height
    End If
End Sub

Private Sub FrameTraspasoAlbRetiradaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameTraspasoAlbRetirada.visible = visible
    If visible = True Then
        Me.FrameTraspasoAlbRetirada.Top = -90
        Me.FrameTraspasoAlbRetirada.Left = 0
        Me.FrameTraspasoAlbRetirada.Height = 4665
        Me.FrameTraspasoAlbRetirada.Width = 6655
        w = Me.FrameTraspasoAlbRetirada.Width
        h = Me.FrameTraspasoAlbRetirada.Height
    End If
End Sub


Private Sub FrameTraspasoAAlmazaraVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameTraspasoAAlmazara.visible = visible
    If visible = True Then
        Me.FrameTraspasoAAlmazara.Top = -90
        Me.FrameTraspasoAAlmazara.Left = 0
        Me.FrameTraspasoAAlmazara.Height = 3450
        Me.FrameTraspasoAAlmazara.Width = 6615
        w = Me.FrameTraspasoAAlmazara.Width
        h = Me.FrameTraspasoAAlmazara.Height
    End If
End Sub

Private Sub FrameBonificacionesVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameBonificaciones.visible = visible
    If visible = True Then
        Me.FrameBonificaciones.Top = -90
        Me.FrameBonificaciones.Left = 0
        Me.FrameBonificaciones.Height = 4800
        Me.FrameBonificaciones.Width = 6615
        w = Me.FrameBonificaciones.Width
        h = Me.FrameBonificaciones.Height
    End If
End Sub


Private Sub FrameGeneraClasificaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameGeneraClasifica.visible = visible
    If visible = True Then
        Me.FrameGeneraClasifica.Top = -90
        Me.FrameGeneraClasifica.Left = 0
        Me.FrameGeneraClasifica.Height = 3390
        Me.FrameGeneraClasifica.Width = 6615
        w = Me.FrameGeneraClasifica.Width
        h = Me.FrameGeneraClasifica.Height
    End If
End Sub


Private Sub FrameInformeFasesVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameInformeFases.visible = visible
    If visible = True Then
        Me.FrameInformeFases.Top = -90
        Me.FrameInformeFases.Left = 0
        Me.FrameInformeFases.Height = 3390
        Me.FrameInformeFases.Width = 6615
        w = Me.FrameInformeFases.Width
        h = Me.FrameInformeFases.Height
    End If
End Sub


Private Sub FrameControlDestrioVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameControlDestrio.visible = visible
    If visible = True Then
        Me.FrameControlDestrio.Top = -90
        Me.FrameControlDestrio.Left = 0
        Me.FrameControlDestrio.Height = 6690
        Me.FrameControlDestrio.Width = 6615
        w = Me.FrameControlDestrio.Width
        h = Me.FrameControlDestrio.Height
    End If
End Sub


Private Sub FrameGastosporConceptoVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameGastosporConcepto.visible = visible
    If visible = True Then
        Me.FrameGastosporConcepto.Top = -90
        Me.FrameGastosporConcepto.Left = 0
        Me.FrameGastosporConcepto.Height = 7680
        Me.FrameGastosporConcepto.Width = 6615
        w = Me.FrameGastosporConcepto.Width
        h = Me.FrameGastosporConcepto.Height
    End If
End Sub

Private Sub FrameCambioSocioVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameCambioSocio.visible = visible
    If visible = True Then
        Me.FrameCambioSocio.Top = -90
        Me.FrameCambioSocio.Left = 0
        Me.FrameCambioSocio.Height = 4890
        Me.FrameCambioSocio.Width = 6615
        w = Me.FrameCambioSocio.Width
        h = Me.FrameCambioSocio.Height
    End If
End Sub

Private Sub FrameVentaFrutaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameVentaFruta.visible = visible
    If visible = True Then
        Me.FrameVentaFruta.Top = -90
        Me.FrameVentaFruta.Left = 0
        Me.FrameVentaFruta.Height = 6690
        Me.FrameVentaFruta.Width = 6615
        w = Me.FrameVentaFruta.Width
        h = Me.FrameVentaFruta.Height
    End If
End Sub

Private Sub FrameGastosCamposVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameGastosCampos.visible = visible
    If visible = True Then
        Me.FrameGastosCampos.Top = -90
        Me.FrameGastosCampos.Left = 0
        Me.FrameGastosCampos.Height = 6720
        Me.FrameGastosCampos.Width = 6765
        w = Me.FrameGastosCampos.Width
        h = Me.FrameGastosCampos.Height
    End If
End Sub

Private Sub FrameContabGastosCamposVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameContabGastos.visible = visible
    If visible = True Then
        Me.FrameContabGastos.Top = -90
        Me.FrameContabGastos.Left = 0
        Me.FrameContabGastos.Height = 5220
        Me.FrameContabGastos.Width = 6615
        w = Me.FrameContabGastos.Width
        h = Me.FrameContabGastos.Height
    End If
End Sub

Private Sub FrameKilosRecolectVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para baja de socios
    Me.FrameKilosRecolect.visible = visible
    If visible = True Then
        Me.FrameKilosRecolect.Top = -90
        Me.FrameKilosRecolect.Left = 0
        Me.FrameKilosRecolect.Height = 6840
        Me.FrameKilosRecolect.Width = 6615
        w = Me.FrameKilosRecolect.Width
        h = Me.FrameKilosRecolect.Height
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
            ItmX.Text = "Alfab�tico"
        Case 1
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Clase"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Producto"
        Case 2
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Variedad"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Calidad"
        Case 3
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Seccion"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Socio"
        Case 4
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Trabajador"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Fecha"
    End Select
        
End Sub

Private Sub CargarListViewTEntrada()
Dim ItmX As ListItem

    'Los encabezados
    ListView1(0).ColumnHeaders.Clear
    ListView1(0).ColumnHeaders.Add , , "Tipo Entrada", 1890

    Set ItmX = ListView1(0).ListItems.Add
    ItmX.Text = "Normal"
    Set ItmX = ListView1(0).ListItems.Add
    ItmX.Text = "Venta Campo"
    Set ItmX = ListView1(0).ListItems.Add
    ItmX.Text = "Producto Integrado"
    Set ItmX = ListView1(0).ListItems.Add
    ItmX.Text = "Industria Directo"
    Set ItmX = ListView1(0).ListItems.Add
    ItmX.Text = "Retirada"
    Set ItmX = ListView1(0).ListItems.Add
    ItmX.Text = "Venta Comercio"
        
    imgCheck_Click (1)


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
'A�ade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y a�ade a cadParam la cadena para mostrar en la cabecera informe:
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
        .outTipoDocumento = 101
        .outClaveNombreArchiv = "Resultado"
        .outCodigoCliProv = 0
        If OpcionListado = 17 Then
            If txtCodigo(12).Text = txtCodigo(13).Text And txtCodigo(12).Text <> "" Then
                .outCodigoCliProv = txtCodigo(12).Text
            End If
        End If
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = ConSubInforme
        .Opcion = OpcionListado
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
            CadParam = CadParam & campo & "{" & tabla & ".codclase}" & "|"
            CadParam = CadParam & nomCampo & " {" & "clases" & ".nomclase}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Producto""" & "|"
            numParam = numParam + 3
            
        Case "Producto"
            CadParam = CadParam & campo & "{" & tabla & ".codprodu}" & "|"
            CadParam = CadParam & nomCampo & " {" & "productos" & ".nomprodu}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3

        'Informe de calibres
        Case "Seccion"
            CadParam = CadParam & campo & "{" & tabla & ".codsecci}" & "|"
            CadParam = CadParam & nomCampo & "{rseccion.nomsecci}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Seccion""" & "|"
            numParam = numParam + 3
            
        Case "Socio"
            CadParam = CadParam & campo & "{" & tabla & ".codsocio}" & "|"
            CadParam = CadParam & nomCampo & " {" & "rsocios" & ".nomsocio}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Socio""" & "|"
            numParam = numParam + 3
            
        'Informe de calidades
        Case "Variedad"
            CadParam = CadParam & campo & "{" & tabla & ".codvarie}" & "|"
            CadParam = CadParam & nomCampo & "{variedades.nomvarie}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Variedad""" & "|"
            numParam = numParam + 3
            
        Case "Calidad"
            CadParam = CadParam & campo & "{" & tabla & ".codcalid}" & "|"
            CadParam = CadParam & nomCampo & " {" & "rcalidad" & ".nomcalid}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Calidad""" & "|"
            numParam = numParam + 3
            
            
        'Informe de campos
        Case "Socios"
            CadParam = CadParam & campo & "{rcampos.codsocio}" & "|"
            CadParam = CadParam & nomCampo & "{rsocios.nomsocio}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Socio""" & "|"
            numParam = numParam + 3
            
        Case "Clases"
            CadParam = CadParam & campo & "{variedades.codclase}" & "|"
            CadParam = CadParam & nomCampo & " {clases.nomclase}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3
            CadParam = CadParam & "pOrden={rcampos.codvarie}|"
            numParam = numParam + 1
            
            If vParamAplic.Cooperativa = 16 Then
                If Check27.Value = 1 Then
                    CadParam = CadParam & "pOrden1={rpartida.nomparti}|"
                    numParam = numParam + 1
                    CadParam = CadParam & "pOrden2={rcampos.codsocio}|"
                    numParam = numParam + 1
                Else
                    CadParam = CadParam & "pOrden1={rsocios.nomsocio}|"
                    numParam = numParam + 1
                    CadParam = CadParam & "pOrden2={rcampos.codsocio}|"
                    numParam = numParam + 1
                End If
            Else
                CadParam = CadParam & "pOrden1={rcampos.codsocio}|"
                numParam = numParam + 1
                CadParam = CadParam & "pOrden2={rcampos.codcampo}|"
                numParam = numParam + 1
            End If
        Case "Terminos"
            CadParam = CadParam & campo & "{rpartida.codpobla}" & "|"
            CadParam = CadParam & nomCampo & " {" & "rpueblos" & ".despobla}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Termino Municipal""" & "|"
            numParam = numParam + 3
            
        Case "Zonas"
            CadParam = CadParam & campo & "{rcampos.codzonas}" & "|"
            CadParam = CadParam & nomCampo & " {" & "rzonas" & ".nomzonas}" & "|"
            '[Monica]10/06/2013: Cambiamos zona por bra�al
            CadParam = CadParam & "pTitulo1=""" & vParamAplic.NomZonaPOZ & """|"     ' "=""Zonas""" & "|"
            numParam = numParam + 3
            
        Case "Variedad/Zona"
            CadParam = CadParam & campo & "{rcampos.codvarie}" & "|"
            CadParam = CadParam & nomCampo & " {variedades.nomvarie}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Variedad/Zona""" & "|"
            numParam = numParam + 3
            CadParam = CadParam & "pOrden={rcampos.codzonas}|"
            numParam = numParam + 1
            CadParam = CadParam & "pOrden1={rcampos.codsocio}|"
            numParam = numParam + 1
            CadParam = CadParam & "pOrden2={rcampos.codcampo}|"
            numParam = numParam + 1


End Select

End Function

Private Function PonerOrden(cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    PonerOrden = 0

    Select Case cadgrupo
        Case "Codigo"
            CadParam = CadParam & "Orden" & "= {" & tabla
            Select Case OpcionListado
                Case 10
                    CadParam = CadParam & ".codclien}|"
                Case 11
                    CadParam = CadParam & ".codprove}|"
            End Select
            Tipo = "C�digo"
        Case "Alfab�tico"
            CadParam = CadParam & "Orden" & "= {" & tabla
            Select Case OpcionListado
                Case 10
                    CadParam = CadParam & ".nomclien}|"
                Case 11
                    CadParam = CadParam & ".nomprove}|"
            End Select
            Tipo = "Alfab�tico"
    End Select
    
    numParam = numParam + 1

End Function

Private Sub AbrirFrmProducto(Indice As Integer)
    indCodigo = Indice
    Set frmProd = New frmComercial
    
    AyudaProductosCom frmProd, txtCodigo(Indice).Text
    
    Set frmProd = Nothing
End Sub

Private Sub AbrirFrmCalidad(Indice As Integer)
    indCodigo = Indice
    Set frmCal = New frmManCalidades
    frmCal.DatosADevolverBusqueda = "2|3|"
    frmCal.Show vbModal
    Set frmCal = Nothing
End Sub


Private Sub AbrirFrmSeccion(Indice As Integer)
    indCodigo = Indice
    Set frmSec = New frmManSeccion
    frmSec.DatosADevolverBusqueda = "0|1|"
    frmSec.Show vbModal
    Set frmSec = Nothing
End Sub

Private Sub AbrirFrmSocios(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmCampos(Indice As Integer)
    indCodigo = Indice
    Set frmCampos = New frmManCampos
    frmCampos.DatosADevolverBusqueda = "0|1|"
    frmCampos.Show vbModal
    Set frmCampos = Nothing
End Sub

Private Sub AbrirFrmIncidencias(Indice As Integer)
    indCodigo = Indice
    
    Set frmInc = New frmManInciden
    frmInc.DatosADevolverBusqueda = "0|1|"
    frmInc.Show vbModal
    
    Set frmInc = Nothing
End Sub

Private Sub AbrirFrmSituacionCampo(Indice As Integer)
    indCodigo = Indice
    Set frmSit = New frmManSituCamp
    frmSit.DatosADevolverBusqueda = "0|1|"
    frmSit.Show vbModal
    Set frmSit = Nothing
End Sub

Private Sub AbrirFrmSituacion(Indice As Integer)
    indCodigo = Indice
    Set frmSitu = New frmManSituacion
    frmSitu.DatosADevolverBusqueda = "0|1|"
    frmSitu.Show vbModal
    Set frmSitu = Nothing
End Sub

Private Sub AbrirFrmCapataz(Indice As Integer)
    indCodigo = Indice
    Set frmCapa = New frmManCapataz
    frmCapa.DatosADevolverBusqueda = "0|1|"
    frmCapa.Show vbModal
    Set frmCapa = Nothing
End Sub

Private Sub AbrirFrmPartidas(Indice As Integer)
    indCodigo = Indice
    Set frmPar = New frmManPartidas
    frmPar.DatosADevolverBusqueda = "0|1|"
    frmPar.Show vbModal
    Set frmPar = Nothing
End Sub

Private Sub AbrirFrmPueblos(Indice As Integer)
    indCodigo = Indice
    Set frmPue = New frmManPueblos
    frmPue.DatosADevolverBusqueda = "0|1|"
    frmPue.Show vbModal
    Set frmPue = Nothing
End Sub



Private Sub AbrirFrmConceptos(Indice As Integer)
    indCodigo = Indice
    Set frmCon = New frmManConcepGasto
    frmCon.DatosADevolverBusqueda = "0|1|"
    frmCon.Show vbModal
    Set frmCon = Nothing
End Sub

Private Sub AbrirFrmConceptosConta(Indice As Integer)
    indCodigo = Indice
    Set frmCConta = New frmConceConta
    frmCConta.DatosADevolverBusqueda = "0|1|"
    frmCConta.Show vbModal
    Set frmCConta = Nothing
End Sub

Private Sub AbrirFrmCuentasConta(Indice As Integer)
    indCodigo = Indice
    Set frmCtaConta = New frmCtasConta
    frmCtaConta.DatosADevolverBusqueda = "0|1|"
    frmCtaConta.Show vbModal
    Set frmCtaConta = Nothing
End Sub


Private Sub AbrirFrmDiariosConta(Indice As Integer)
    indCodigo = Indice
    Set frmDConta = New frmDiaConta
    frmDConta.DatosADevolverBusqueda = "0|1|"
    frmDConta.Show vbModal
    Set frmDConta = Nothing
End Sub


Private Sub AbrirFrmClase(Indice As Integer)
    If Indice = 6 Or Indice = 7 Then
        indCodigo = Indice + 14
    Else
        indCodigo = Indice
    End If
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtCodigo(Indice).Text
    
    Set frmCla = Nothing
End Sub



Private Sub AbrirFrmVariedad(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub


Private Sub AbrirFrmCooperativa(Indice As Integer)
    indCodigo = Indice
    Set frmCoop = New frmManCoope
    frmCoop.DatosADevolverBusqueda = "0|1|"
    frmCoop.Show vbModal
    Set frmCoop = Nothing
End Sub


Private Sub AbrirFrmZonas(Indice As Integer)
    indCodigo = Indice
    Set frmZon = New frmManZonas
    frmZon.DatosADevolverBusqueda = "0|1|"
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        frmZon.Caption = "Bra�als"
    End If
    frmZon.DeInformes = True
    frmZon.Show vbModal
    Set frmZon = Nothing
End Sub


' ********* si n'hi han combos a la cap�alera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer
Dim RS As ADODB.Recordset
Dim sql As String


    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    'tipo de hectareas
    Combo1(0).AddItem "Cooperativa"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Sigpac"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Catastro"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Cultivable"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    
    'tipo de produccion
    Combo1(1).AddItem "Esperada"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Real"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
  
    'tipo de campo
    Combo1(11).AddItem "Normal"
    Combo1(11).ItemData(Combo1(11).NewIndex) = 0
    Combo1(11).AddItem "Comercio"
    Combo1(11).ItemData(Combo1(11).NewIndex) = 1
    Combo1(11).AddItem "Industria"
    Combo1(11).ItemData(Combo1(11).NewIndex) = 2
    Combo1(11).AddItem "Todos"
    Combo1(11).ItemData(Combo1(11).NewIndex) = 3
  
    'tipo de informe de entradas clasificadas
    Combo1(2).AddItem "Todas"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "S�lo Clasificadas"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Pendientes"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
    
    'tipo de registros de entradas clasificadas
    Combo1(3).AddItem "Normal"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 0
    Combo1(3).AddItem "Detalle Clasif."
    Combo1(3).ItemData(Combo1(3).NewIndex) = 1

    'produccion segun tabla
    Combo1(4).AddItem "NZ"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 0
    Combo1(4).AddItem "MZ"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 1
    Combo1(4).AddItem "CZ"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 2
    Combo1(4).AddItem "LZ"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 3
    Combo1(4).AddItem "TZ"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 4
    Combo1(4).AddItem "PZ"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 5
    Combo1(4).AddItem "CG"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 6
    Combo1(4).AddItem "SG"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 7

    'tipo de hectareas
    Combo1(5).AddItem "Cooperativa"
    Combo1(5).ItemData(Combo1(5).NewIndex) = 0
    Combo1(5).AddItem "Sigpac"
    Combo1(5).ItemData(Combo1(5).NewIndex) = 1
    Combo1(5).AddItem "Catastro"
    Combo1(5).ItemData(Combo1(5).NewIndex) = 2
    Combo1(5).AddItem "Cultivable"
    Combo1(5).ItemData(Combo1(5).NewIndex) = 3

    'tipo de calibrador
    Select Case vParamAplic.Cooperativa
        '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
        Case 0, 9 ' 0=catadau 9=natural
            Combo1(6).AddItem "Calibrador Grande"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 0
            Combo1(6).AddItem "Calibrador Peque�o"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 1
        Case 1 ' 1=valsur
            Combo1(6).AddItem "Calibrador 1"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 0
        Case 2, 16 ' Picassent, 20/09/2016: a�ado Coopic 16
            Combo1(6).AddItem "Calibrador 1"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 0
        Case 4 '4=alzira
            Combo1(6).AddItem "Precalibrador"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 0
            Combo1(6).AddItem "Escandalladora"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 1
            Combo1(6).AddItem "Calibrador Kaki"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 2
        Case 5 '5=castelduc
            Combo1(6).AddItem "Calibrador 1"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 0
            Combo1(6).AddItem "Calibrador 2"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 1
            Combo1(6).AddItem "Castello de Rugat"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 2
    End Select
    
    ' tipo de factura a traspasar
    'tipo de factura
    sql = "select codtipom, nomtipom from usuarios.stipom where tipodocu > 0 "

    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    While Not RS.EOF
'        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
        sql = RS.Fields(1).Value
        sql = RS.Fields(0).Value & " - " & sql
        Combo1(7).AddItem sql 'campo del codigo
        Combo1(7).ItemData(Combo1(7).NewIndex) = I
        I = I + 1
        RS.MoveNext
    Wend
    Set RS = Nothing
    

    ' recolectada por
    Combo1(8).AddItem "Cooperativa"
    Combo1(8).ItemData(Combo1(8).NewIndex) = 0
    Combo1(8).AddItem "Socio"
    Combo1(8).ItemData(Combo1(8).NewIndex) = 1
    Combo1(8).AddItem "Ambos"
    Combo1(8).ItemData(Combo1(8).NewIndex) = 2


    ' tipo de entrada
    Combo1(9).AddItem "Normal"
    Combo1(9).ItemData(Combo1(9).NewIndex) = 0
    Combo1(9).AddItem "Venta Campo"
    Combo1(9).ItemData(Combo1(9).NewIndex) = 1
    Combo1(9).AddItem "Prod.Integrado"
    Combo1(9).ItemData(Combo1(9).NewIndex) = 2
    Combo1(9).AddItem "Industria Directo"
    Combo1(9).ItemData(Combo1(9).NewIndex) = 3
    Combo1(9).AddItem "Retirada"
    Combo1(9).ItemData(Combo1(9).NewIndex) = 4
    Combo1(9).AddItem "Venta Comercio"
    Combo1(9).ItemData(Combo1(9).NewIndex) = 5
    Combo1(9).AddItem "Todas"
    Combo1(9).ItemData(Combo1(9).NewIndex) = 6
    
    '[Monica]17/07/2014: a�adido el tipo de socio al listado de clasificacion (NATURAL)
    'tipo de socio tipoprod
    Combo1(14).AddItem "Socio"
    Combo1(14).ItemData(Combo1(14).NewIndex) = 0
    Combo1(14).AddItem "Tercero"
    Combo1(14).ItemData(Combo1(14).NewIndex) = 1
    Combo1(14).AddItem "Otra OPA"
    Combo1(14).ItemData(Combo1(14).NewIndex) = 2
    Combo1(14).AddItem "Aportacionista"
    Combo1(14).ItemData(Combo1(14).NewIndex) = 3
    Combo1(14).AddItem "Todos"
    Combo1(14).ItemData(Combo1(14).NewIndex) = 4
       
    
    

    'tipo de albaran
    Combo1(10).AddItem "No Impresos"
    Combo1(10).ItemData(Combo1(10).NewIndex) = 0
    Combo1(10).AddItem "Impresos"
    Combo1(10).ItemData(Combo1(10).NewIndex) = 1
    Combo1(10).AddItem "Todos"
    Combo1(10).ItemData(Combo1(10).NewIndex) = 2


    
    Combo1(12).AddItem "Todos" 'campo del codigo
    Combo1(12).ItemData(Combo1(12).NewIndex) = 0
    
    sql = "select distinct numfases from rsocios_pozos order by 1"

    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    While Not RS.EOF
'        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
        sql = RS.Fields(0).Value
        Combo1(12).AddItem sql 'campo del codigo
        Combo1(12).ItemData(Combo1(12).NewIndex) = I
        I = I + 1
        RS.MoveNext
    Wend
    Set RS = Nothing

    ' tipo de precios
    Combo1(13).AddItem "Anticipo"
    Combo1(13).ItemData(Combo1(13).NewIndex) = 0
    Combo1(13).AddItem "Liquidacion"
    Combo1(13).ItemData(Combo1(13).NewIndex) = 1
    
    ' solo hay industria directa y complementaria en horto
    Combo1(13).AddItem "Industria Directa"
    Combo1(13).ItemData(Combo1(13).NewIndex) = 2
    Combo1(13).AddItem "Complementaria"
    Combo1(13).ItemData(Combo1(13).NewIndex) = 3
    Combo1(13).AddItem "Anticipo Gen�rico"
    Combo1(13).ItemData(Combo1(13).NewIndex) = 4
    Combo1(13).AddItem "Anticipo Retirada"
    Combo1(13).ItemData(Combo1(13).NewIndex) = 5

    '[Monica]08/04/2015
    ' tipo de socio (listado de socios por seccion en el caso de catadau)
    Combo1(15).AddItem "Todos"
    Combo1(15).ItemData(Combo1(15).NewIndex) = 0
    Combo1(15).AddItem "S�lo Socios"
    Combo1(15).ItemData(Combo1(15).NewIndex) = 1
    Combo1(15).AddItem "S�lo Asociados"
    Combo1(15).ItemData(Combo1(15).NewIndex) = 2

End Sub

Private Function ConcatenarCampos(cTabla As String, cWhere As String) As String
Dim RS As ADODB.Recordset
Dim sql As String
Dim Sql1 As String

    ConcatenarCampos = ""

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    
    
    sql = "select distinct rcampos.codcampo  from " & cTabla & " where " & cWhere
    Set RS = New ADODB.Recordset
    
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not RS.EOF
        Sql1 = Sql1 & DBLet(RS.Fields(0).Value, "N") & ","
        RS.MoveNext
    Wend
    Set RS = Nothing
    'quitamos el ultimo or
    ConcatenarCampos = Mid(Sql1, 1, Len(Sql1) - 1)
    
End Function

Private Function CargarTemporal(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String

    
    On Error GoTo eCargarTemporal
    
    CargarTemporal = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    
    sql = "select distinct rcampos.codcampo  from " & cTabla & " where " & cWhere
    
'[Monica]14/12/2010: en los siguientes selects anteriormente estaba el sum sobre la tabla rentradas. Ahora lo cambio a rhisfruta
    
    If Opcion1(0) Then ' socios
        Sql1 = "select " & vUsu.Codigo & ",rcampos.codsocio, sum(kilosnet) from rhisfruta right join rcampos on rhisfruta.codcampo = rcampos.codcampo "
        Sql1 = Sql1 & " where rcampos.codcampo in (" & sql & ")"
        Sql1 = Sql1 & " group by 1,2"
        
        Sql2 = "insert into tmpinformes (codusu, codigo1, importe1) " & Sql1
        conn.Execute Sql2
    End If
    
    If Opcion1(1) Then ' clases
        
        If Combo1(1).ListIndex = 1 Then
            Sql1 = "select " & vUsu.Codigo & ",variedades.codclase, variedades.codvarie, sum(if(kilosnet is null,0,kilosnet)), 0, 0"
'            '[Monica]23/09/2011: agrupamos por variedad tambien en el resumen
'            Select Case Combo1(0).ListIndex
'                Case 0
'                    Sql1 = Sql1 & "sum(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supcoope} / 0.0831,2)
'                Case 1
'                    Sql1 = Sql1 & "sum(round(rcampos.supsigpa) / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supsigpa} / 0.0831,2)
'                Case 2
'                    Sql1 = Sql1 & "sum(round(rcampos.supcatas) / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supcatas} / 0.0831,2)
'                Case 3
'                    Sql1 = Sql1 & "sum(round(rcampos.supculti) / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supculti} / 0.0831,2)
'            End Select
'
'
'            Sql1 = Sql1 & " from rhisfruta right join (rcampos inner join variedades on rcampos.codvarie = variedades.codvarie) on rhisfruta.codcampo = rcampos.codcampo and rhisfruta.codvarie = rcampos.codvarie where rcampos.codcampo in (" & Sql & ")"
'            Sql1 = Sql1 & " group by 1,2,3 "
            
            Sql1 = Sql1 & " from rhisfruta inner join variedades on rhisfruta.codvarie = variedades.codvarie where rhisfruta.codcampo in (" & sql & ")"
            Sql1 = Sql1 & " group by 1,2,3 "
            
            Sql2 = "insert into tmpinformes (codusu, codigo1, importe2, importe1, importe3, precio1) " & Sql1
            conn.Execute Sql2
            
            
            '[Monica]28/07/2014: en elcaso de ser por clase si no tiene existencia real tiene que aparecer con 0 y con superficie
            Sql1 = "select " & vUsu.Codigo & ",variedades.codclase, variedades.codvarie, 0, 0, 0"
            Sql1 = Sql1 & " from rcampos inner join variedades on rcampos.codvarie = variedades.codvarie where rcampos.codcampo in (" & sql & ")"
            Sql1 = Sql1 & " and not (variedades.codclase, variedades.codvarie) in (select codigo1, importe2 from tmpinformes where codusu = " & vUsu.Codigo & ")"
            Sql1 = Sql1 & " group by 1,2,3 "
            Sql2 = "insert into tmpinformes (codusu, codigo1, importe2, importe1, importe3, precio1) " & Sql1
            conn.Execute Sql2
            
            Sql1 = "update tmpinformes set precio1 = (select "
            If cadNombreRPT = "rInfCampos.rpt" Then
                Select Case Combo1(0).ListIndex
                    Case 0
                        Sql1 = Sql1 & "sum(rcampos.supcoope) "    '{rcampos.supcoope} / 0.0831,2)
                    Case 1
                        Sql1 = Sql1 & "sum(rcampos.supsigpa) "    '{rcampos.supsigpa} / 0.0831,2)
                    Case 2
                        Sql1 = Sql1 & "sum(rcampos.supcatas) "    '{rcampos.supcatas} / 0.0831,2)
                    Case 3
                        Sql1 = Sql1 & "sum(rcampos.supculti) "    '{rcampos.supculti} / 0.0831,2)
                End Select
            Else
                Select Case Combo1(0).ListIndex
                    Case 0
                        Sql1 = Sql1 & "sum(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supcoope} / 0.0831,2)
                    Case 1
                        Sql1 = Sql1 & "sum(round(rcampos.supsigpa / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supsigpa} / 0.0831,2)
                    Case 2
                        Sql1 = Sql1 & "sum(round(rcampos.supcatas / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supcatas} / 0.0831,2)
                    Case 3
                        Sql1 = Sql1 & "sum(round(rcampos.supculti / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supculti} / 0.0831,2)
                End Select
            End If
            Sql1 = Sql1 & " from rcampos where rcampos.codcampo in (" & sql & ") and tmpinformes.importe2 = rcampos.codvarie )"
            
            conn.Execute Sql1
            
            
        Else
            Sql1 = "select " & vUsu.Codigo & ",variedades.codclase, variedades.codvarie, 0, sum(if(canaforo is null,0,canaforo)), "
            '[Monica]23/09/2011: agrupamos por variedad tambien en el resumen
            If cadNombreRPT = "rInfCampos.rpt" Then
                Select Case Combo1(0).ListIndex
                    Case 0
                        Sql1 = Sql1 & "sum(rcampos.supcoope) "    '{rcampos.supcoope} / 0.0831,2)
                    Case 1
                        Sql1 = Sql1 & "sum(rcampos.supsigpa) "    '{rcampos.supsigpa} / 0.0831,2)
                    Case 2
                        Sql1 = Sql1 & "sum(rcampos.supcatas) "    '{rcampos.supcatas} / 0.0831,2)
                    Case 3
                        Sql1 = Sql1 & "sum(rcampos.supculti) "    '{rcampos.supculti} / 0.0831,2)
                End Select
            Else
                Select Case Combo1(0).ListIndex
                    Case 0
                        Sql1 = Sql1 & "sum(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supcoope} / 0.0831,2)
                    Case 1
                        Sql1 = Sql1 & "sum(round(rcampos.supsigpa / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supsigpa} / 0.0831,2)
                    Case 2
                        Sql1 = Sql1 & "sum(round(rcampos.supcatas / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supcatas} / 0.0831,2)
                    Case 3
                        Sql1 = Sql1 & "sum(round(rcampos.supculti / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supculti} / 0.0831,2)
                End Select
            End If
            
            Sql1 = Sql1 & " from (rcampos inner join variedades on rcampos.codvarie = variedades.codvarie)  where rcampos.codcampo in (" & sql & ")"
            Sql1 = Sql1 & " group by 1,2,3,4 "
        
            Sql2 = "insert into tmpinformes (codusu, codigo1, importe2, importe1, importe3, precio1) " & Sql1
            conn.Execute Sql2
        
        End If
            
    End If
    
    If Opcion1(2) Then ' terminos
        Sql1 = "select " & vUsu.Codigo & ", rpartida.codpobla, sum(kilosnet) from rhisfruta right join (rcampos inner join rpartida on rcampos.codparti = rpartida.codparti)  on rhisfruta.codcampo = rcampos.codcampo where rcampos.codcampo in (" & sql & ")"
        Sql1 = Sql1 & " group by 1,2"
    
        Sql2 = "insert into tmpinformes (codusu, nombre1, importe1) " & Sql1
        conn.Execute Sql2
    End If
    
    If Opcion1(3) Then ' zonas
        Sql1 = "select " & vUsu.Codigo & ", rpartida.codzonas, sum(kilosnet) from rhisfruta right join (rcampos inner join rpartida on rcampos.codparti = rpartida.codparti)  on rhisfruta.codcampo = rcampos.codcampo where rcampos.codcampo in (" & sql & ")"
        Sql1 = Sql1 & " group by 1,2"
    
        Sql2 = "insert into tmpinformes (codusu, codigo1, importe1) " & Sql1
        conn.Execute Sql2
    End If
    
    '[Monica]29/09/2014:
    If Opcion1(7) Then ' variedad/zona
        If Combo1(1).ListIndex = 1 Then
            Sql1 = "select " & vUsu.Codigo & ", variedades.codvarie, rcampos.codzonas, sum(if(kilosnet is null,0,kilosnet)), 0, 0"
            
            Sql1 = Sql1 & " from (rhisfruta inner join variedades on rhisfruta.codvarie = variedades.codvarie) inner join rcampos on rhisfruta.codcampo = rcampos.codcampo where rhisfruta.codcampo in (" & sql & ")"
            Sql1 = Sql1 & " group by 1,2,3 "
            
            Sql2 = "insert into tmpinformes (codusu, codigo1, importe2, importe1, importe3, precio1) " & Sql1
            conn.Execute Sql2
            
            
            '[Monica]28/07/2014: en elcaso de ser por clase si no tiene existencia real tiene que aparecer con 0 y con superficie
            Sql1 = "select " & vUsu.Codigo & ", variedades.codvarie, rcampos.codzonas, 0, 0, 0"
            Sql1 = Sql1 & " from rcampos inner join variedades on rcampos.codvarie = variedades.codvarie where rcampos.codcampo in (" & sql & ")"
            Sql1 = Sql1 & " and not (variedades.codvarie, rcampos.codzonas) in (select codigo1, importe2 from tmpinformes where codusu = " & vUsu.Codigo & ")"
            Sql1 = Sql1 & " group by 1,2,3 "
            Sql2 = "insert into tmpinformes (codusu, codigo1, importe2, importe1, importe3, precio1) " & Sql1
            conn.Execute Sql2
            
            Sql1 = "update tmpinformes set precio1 = (select "
            Select Case Combo1(0).ListIndex
                Case 0
                    Sql1 = Sql1 & "sum(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supcoope} / 0.0831,2)
                Case 1
                    Sql1 = Sql1 & "sum(round(rcampos.supsigpa / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supsigpa} / 0.0831,2)
                Case 2
                    Sql1 = Sql1 & "sum(round(rcampos.supcatas / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supcatas} / 0.0831,2)
                Case 3
                    Sql1 = Sql1 & "sum(round(rcampos.supculti / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supculti} / 0.0831,2)
            End Select
            Sql1 = Sql1 & " from rcampos where rcampos.codcampo in (" & sql & ") and tmpinformes.importe2 = rcampos.codvarie )"
            
            conn.Execute Sql1
            
            
        Else
            Sql1 = "select " & vUsu.Codigo & ", variedades.codvarie, rcampos.codzonas, 0, sum(if(canaforo is null,0,canaforo)), "
            Select Case Combo1(0).ListIndex
                Case 0
                    Sql1 = Sql1 & "sum(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supcoope} / 0.0831,2)
                Case 1
                    Sql1 = Sql1 & "sum(round(rcampos.supsigpa / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supsigpa} / 0.0831,2)
                Case 2
                    Sql1 = Sql1 & "sum(round(rcampos.supcatas / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supcatas} / 0.0831,2)
                Case 3
                    Sql1 = Sql1 & "sum(round(rcampos.supculti / " & DBSet(vParamAplic.Faneca, "N") & ",2)) "    '{rcampos.supculti} / 0.0831,2)
            End Select
            
            Sql1 = Sql1 & " from (rcampos inner join variedades on rcampos.codvarie = variedades.codvarie)  where rcampos.codcampo in (" & sql & ")"
            Sql1 = Sql1 & " group by 1,2,3,4 "
        
            Sql2 = "insert into tmpinformes (codusu, codigo1, importe2, importe1, importe3, precio1) " & Sql1
            conn.Execute Sql2
        
        End If
            
    
    End If
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporal2(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String
    
    On Error GoTo eCargarTemporal
    
    CargarTemporal2 = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "Select rclasifica.numnotac FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    If cWhere <> "" Then
        sql = "select distinct rclasifica.numnotac  from " & cTabla & " where " & cWhere
    Else
        sql = "select distinct rclasifica.numnotac  from " & cTabla
    End If
    
    Select Case Combo1(2).ListIndex
        Case 0: 'todas
            Sql1 = "select " & vUsu.Codigo & ", rclasifica.numnotac, 0 from rclasifica where numnotac in (" & sql & ")"
        
        Case 1: ' solo clasificadas
            Sql1 = "select " & vUsu.Codigo & ",rclasifica_clasif.numnotac, sum(rclasifica_clasif.kilosnet) from rclasifica_clasif inner join rclasifica on rclasifica_clasif.numnotac = rclasifica.numnotac "
            Sql1 = Sql1 & " where rclasifica.numnotac in (" & sql & ")"
            Sql1 = Sql1 & " group by 1,2 "
            Sql1 = Sql1 & " having not sum(rclasifica_clasif.kilosnet)  is null "
            
        
        Case 2: ' pendientes
            Sql1 = "select " & vUsu.Codigo & ",rclasifica.numnotac, sum(rclasifica_clasif.kilosnet) from rclasifica left join rclasifica_clasif on rclasifica_clasif.numnotac = rclasifica.numnotac "
            Sql1 = Sql1 & " where rclasifica.numnotac in (" & sql & ")"
            Sql1 = Sql1 & " group by 1,2 "
            Sql1 = Sql1 & " having sum(rclasifica_clasif.kilosnet) is null "
    End Select
    
    Sql2 = "insert into tmpinformes (codusu, codigo1, importe1) " & Sql1
    conn.Execute Sql2
    
    CargarTemporal2 = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function



Private Function CargarTemporal3(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim vSQL As String
Dim res As String
Dim Res1 As String
Dim I As Integer
Dim Clase As String

    On Error GoTo eCargarTemporal
    
    CargarTemporal3 = False

    Sql2 = "delete from tmpclasifica where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "Select rclasifica.numnotac FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    If cWhere <> "" Then
        sql = "select distinct rclasifica.numnotac  from " & cTabla & " where " & cWhere
    Else
        sql = "select distinct rclasifica.numnotac  from " & cTabla
    End If
    
  ' solo clasificadas
    Sql1 = "select rclasifica.numnotac, rclasifica.codvarie, rclasifica.codcampo, rclasifica.codsocio,sum(rclasifica_clasif.kilosnet) from rclasifica inner join rclasifica_clasif on rclasifica.numnotac = rclasifica_clasif.numnotac "
    Sql1 = Sql1 & " where rclasifica.numnotac in (" & sql & ")"
    Sql1 = Sql1 & " group by 1,2,3,4 "
    Sql1 = Sql1 & " having not sum(rclasifica_clasif.kilosnet) is null "
        
    Set RS = New ADODB.Recordset
    
    RS.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not RS.EOF
        Clase = DevuelveDesdeBDNew(cAgro, "variedades", "codclase", "codvarie", RS!codvarie, "N")
        Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " order by 1 "
        
        Set Rs1 = New ADODB.Recordset
        
        res = ""
        Res1 = ""
        I = 0
        Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs1.EOF
            I = I + 1
            vSQL = "select kilosnet from rclasifica_clasif where numnotac= " & DBSet(RS!numnotac, "N")
            vSQL = vSQL & " and codcalid = " & DBSet(Rs1!codcalid, "N")
            
            res = res & "cal" & I & "," 'Format(Rs1!codcalid, "00") & ","
            Res1 = Res1 & DBSet(TotalRegistros(vSQL), "N") & ","
            
            Rs1.MoveNext
        Wend
        
        Set Rs1 = Nothing
        
        
        Sql2 = "insert into tmpclasifica (codusu, codcampo, codsocio, numnotac, codvarie, codclase, "
        Sql2 = Sql2 & Mid(res, 1, Len(res) - 1) & ") values ("
        Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(RS!codcampo, "N") & "," & DBSet(RS!Codsocio, "N") & ","
        Sql2 = Sql2 & DBSet(RS!numnotac, "N") & "," & DBSet(RS!codvarie, "N") & "," & DBSet(Clase, "N") & ","
        Sql2 = Sql2 & Mid(Res1, 1, Len(Res1) - 1) & ")"
        
        conn.Execute Sql2
        
        RS.MoveNext
        
    Wend
    
    Set RS = Nothing
    
    CargarTemporal3 = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function

'
' carga temporal para sacar informe de kilos / gastos de la rhisfruta
'
Private Function CargarTemporal4(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim vSQL As String
Dim res As String
Dim Res1 As String
Dim I As Integer
Dim Clase As String

    On Error GoTo eCargarTemporal
    
    CargarTemporal4 = False

    Sql2 = "delete from tmpclasifica where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "select rhisfruta.numalbar, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.codsocio, rhisfruta.kilosnet "
    sql = sql & " from " & cTabla
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    
        
    Set RS = New ADODB.Recordset
    
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not RS.EOF
        Clase = DevuelveDesdeBDNew(cAgro, "variedades", "codclase", "codvarie", RS!codvarie, "N")
        Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " order by 1 "
        
        Set Rs1 = New ADODB.Recordset
        
        res = ""
        Res1 = ""
        I = 0
        Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs1.EOF
            I = I + 1
            vSQL = "select kilosnet from rhisfruta_clasif where numalbar= " & DBSet(RS!numalbar, "N")
            vSQL = vSQL & " and codvarie = " & DBSet(RS!codvarie, "N")
            vSQL = vSQL & " and codcalid = " & DBSet(Rs1!codcalid, "N")
            
            res = res & "cal" & I & ","
            Res1 = Res1 & DBSet(TotalRegistros(vSQL), "N") & ","
            
            Rs1.MoveNext
        Wend
        
        Set Rs1 = Nothing
        
        
        Sql2 = "insert into tmpclasifica (codusu, codcampo, codsocio, numnotac, codvarie, codclase, "
        Sql2 = Sql2 & Mid(res, 1, Len(res) - 1) & ") values ("
        Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(RS!codcampo, "N") & "," & DBSet(RS!Codsocio, "N") & ","
        Sql2 = Sql2 & DBSet(RS!numalbar, "N") & "," & DBSet(RS!codvarie, "N") & "," & DBSet(Clase, "N") & ","
        Sql2 = Sql2 & Mid(Res1, 1, Len(Res1) - 1) & ")"
        
        conn.Execute Sql2
        
        RS.MoveNext
        
    Wend
    
    Set RS = Nothing
    
    CargarTemporal4 = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function

'
' carga temporal para sacar informe de kilos / gastos de la rhisfruta
'
Private Function CargarTemporal4New(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim vSQL As String
Dim res As String
Dim Res1 As String
Dim I As Integer
Dim J As Integer
Dim Clase As String
Dim m As Integer

    On Error GoTo eCargarTemporal
    
    CargarTemporal4New = False

    sql = "DROP TABLE IF EXISTS tmp; "
    conn.Execute sql
    
    
    Sql2 = "delete from tmpclasifica2 where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "select rhisfruta.numalbar, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.codsocio, rhisfruta.kilosnet "
    sql = sql & " from " & cTabla
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    
    Set RS = New ADODB.Recordset
    
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not RS.EOF
        Clase = DevuelveDesdeBDNew(cAgro, "variedades", "codclase", "codvarie", RS!codvarie, "N")
        Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " order by 1 "
        
        Set Rs1 = New ADODB.Recordset
        
        res = ""
        Res1 = ""
        I = 0
        Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs1.EOF Then
            While Not Rs1.EOF
                I = I + 1
                vSQL = "select kilosnet from rhisfruta_clasif where numalbar= " & DBSet(RS!numalbar, "N")
                vSQL = vSQL & " and codvarie = " & DBSet(RS!codvarie, "N")
                vSQL = vSQL & " and codcalid = " & DBSet(Rs1!codcalid, "N")
                
                res = res & "nomcal" & I & "," & "kilcal" & I & ","
                Res1 = Res1 & DBSet(NombreCalidad(CStr(RS!codvarie), CStr(Rs1!codcalid)), "T") & "," & DBSet(TotalRegistros(vSQL), "N") & ","
                
                Rs1.MoveNext
            Wend
            
            Set Rs1 = Nothing
            
            
            Sql2 = "insert into tmpclasifica2 (codusu, codcampo, codsocio, numnotac, codvarie, codclase, "
            Sql2 = Sql2 & Mid(res, 1, Len(res) - 1) & ") values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(RS!codcampo, "N") & "," & DBSet(RS!Codsocio, "N") & ","
            Sql2 = Sql2 & DBSet(RS!numalbar, "N") & "," & DBSet(RS!codvarie, "N") & "," & DBSet(Clase, "N") & ","
            Sql2 = Sql2 & Mid(Res1, 1, Len(Res1) - 1) & ")"
            
            conn.Execute Sql2
        End If
        
        RS.MoveNext
        
    Wend
    
    Set RS = Nothing
    
    sql = "select codvarie, sum(kilcal1), sum(kilcal2) as kilos2, sum(kilcal3) as kilos3, sum(kilcal4) as kilos4, sum(kilcal5), sum(kilcal6), sum(kilcal7), sum(kilcal8), "
    sql = sql & " sum(kilcal9), sum(kilcal10), sum(kilcal11), sum(kilcal12), sum(kilcal13), sum(kilcal14), sum(kilcal15), sum(kilcal16),"
    sql = sql & " sum(kilcal17), sum(kilcal18), sum(kilcal19), sum(kilcal20) from tmpclasifica2 "
    sql = sql & " where codusu = " & vUsu.Codigo
    sql = sql & " group by 1 "
    
    
    Set Rs1 = New ADODB.Recordset
    
    Rs1.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs1.EOF
        m = 1 ' para evitar que sean todos ceros y haya un bucle infinito
        I = 1
        
        While I < 20 And m < 40
            sql = "select codvarie, sum(kilcal1), sum(kilcal2) as kilos2, sum(kilcal3) as kilos3, sum(kilcal4) as kilos4, sum(kilcal5), sum(kilcal6), sum(kilcal7), sum(kilcal8), "
            sql = sql & " sum(kilcal9), sum(kilcal10), sum(kilcal11), sum(kilcal12), sum(kilcal13), sum(kilcal14), sum(kilcal15), sum(kilcal16),"
            sql = sql & " sum(kilcal17), sum(kilcal18), sum(kilcal19), sum(kilcal20) from tmpclasifica2 "
            sql = sql & " where codusu = " & vUsu.Codigo
            sql = sql & " and codvarie = " & DBSet(Rs1!codvarie, "N")
            sql = sql & " group by 1 "
        
            Set Rs2 = New ADODB.Recordset
            
            Rs2.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            If DBLet(Rs2.Fields(I).Value, "N") = 0 Then
                sql = "update tmpclasifica2 set kilcal" & I & "=kilcal" & I + 1 & ","
                sql = sql & " nomcal" & I & "=nomcal" & I + 1
                
                For J = I + 1 To 19
                    sql = sql & ", kilcal" & J & "=kilcal" & J + 1
                    sql = sql & ", nomcal" & J & "=nomcal" & J + 1
                Next J
                
                sql = sql & ", kilcal20=" & ValorNulo
                sql = sql & ", nomcal20=" & ValorNulo
                sql = sql & " where codvarie = " & DBSet(Rs1.Fields(0).Value, "N")
                sql = sql & " and codusu = " & vUsu.Codigo
                
                conn.Execute sql
                
            Else
                I = I + 1
          
            End If
            
            m = m + 1
            
            Set Rs2 = Nothing
            
        Wend
    
        Rs1.MoveNext
    Wend
    
    Set Rs1 = Nothing
    
    CargarTemporal4New = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporal5(cTabla As String, cWhere As String, cTabla2 As String) As Boolean
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql3 As String
Dim SocioAnt As Long
Dim CampoAnt As Long
Dim SocioAct As Long
Dim CampoAct As Long
Dim SQLaux As String
Dim SqlAux2 As String
Dim Ha As Currency
Dim Producto As String
    
'[Monica]13/11/2013: prorrateamos segun los coopropietarios
Dim Porcen As Currency
Dim Canaforo As String
Dim Hanegadas As Currency
Dim Hectareas As Currency
Dim Arboles As Long
                    
Dim DCanaforo As Long
Dim DHanegada As Currency
Dim DHectarea As Currency
Dim DNroArbol As Long
                    
    
    On Error GoTo eCargarTemporal
    
    CargarTemporal5 = False

    Sql2 = "delete from tmpinfkilos where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
    End If
    
    sql = "select distinct rcampos.codsocio, rcampos.codcampo "
    sql = sql & " from " & cTabla
    sql = sql & " where rcampos.fecbajas is null "
    If cWhere <> "" Then
        sql = sql & " and " & cWhere
    End If
    sql = sql & " union "
    sql = sql & " select distinct rhisfruta.codsocio, rhisfruta.codcampo "
    sql = sql & " from (" & cTabla & ") inner join rhisfruta on rcampos.codcampo = rhisfruta.codcampo and rcampos.codsocio = rhisfruta.codsocio "
    If cWhere <> "" Then
        sql = sql & " where " & cWhere
    End If
    If txtCodigo(39).Text <> "" Then sql = sql & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(39).Text, "F")
    If txtCodigo(40).Text <> "" Then sql = sql & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(40).Text, "F")
    
    '[Monica]13/11/2013: faltan los medieros para sacar los kilos de las entradas
    sql = sql & " union "
    sql = sql & " select distinct rhisfruta.codsocio, rhisfruta.codcampo "
    sql = sql & " from (" & cTabla2 & ") inner join rhisfruta on rcampos_cooprop.codcampo = rhisfruta.codcampo and rcampos_cooprop.codsocio = rhisfruta.codsocio "
    If cWhere <> "" Then
        sql = sql & " where " & cWhere
    End If
    If txtCodigo(39).Text <> "" Then sql = sql & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(39).Text, "F")
    If txtCodigo(40).Text <> "" Then sql = sql & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(40).Text, "F")
    
    
    sql = sql & " order by 1, 2"
    
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql2 = "insert into tmpinfkilos (codusu, codprodu, codsocio, codcampo, kilosnet, porcen,"
    Sql2 = Sql2 & "canaforo, hanegada, hectarea, rdtohane, rdtohecta, nroarbol) values "
    
    While Not RS.EOF
        SocioAct = DBLet(RS.Fields(0).Value, "N")
        CampoAct = DBLet(RS.Fields(1).Value, "N")
            
        Producto = ProductoCampo(DBLet(RS.Fields(1).Value, "N"))
            
        Sql3 = "(" & vUsu.Codigo & "," & DBSet(Producto, "N") & "," & DBSet(RS.Fields(0).Value, "N") & "," & DBSet(RS.Fields(1).Value, "N") & ","
        
        SQLaux = "select sum(kilosnet) from rhisfruta where codsocio = " & DBSet(RS.Fields(0).Value, "N")
        SQLaux = SQLaux & " and codcampo = " & DBSet(RS.Fields(1).Value, "N")
        If txtCodigo(39).Text <> "" Then SQLaux = SQLaux & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(39).Text, "F")
        If txtCodigo(40).Text <> "" Then SQLaux = SQLaux & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(40).Text, "F")
        
        Sql3 = Sql3 & DBSet(DevuelveValor(SQLaux), "N") & ",0," 'kilosnet
        
        SqlAux2 = "select canaforo, "
        Select Case Combo1(5).ListIndex
            Case 0
                SqlAux2 = SqlAux2 & " supcoope, nroarbol"
            Case 1
                SqlAux2 = SqlAux2 & " supsigpa, nroarbol"
            Case 2
                SqlAux2 = SqlAux2 & " supcatas, nroarbol"
            Case 3
                SqlAux2 = SqlAux2 & " supculti, nroarbol"
        
        End Select
        SqlAux2 = SqlAux2 & " from rcampos where codcampo = " & DBSet(RS.Fields(1).Value, "N")
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open SqlAux2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs1.EOF Then
            '[Monica]13/11/2013: a�adimos el porcentaje de coopropiedad
            Porcen = PorCoopropiedadCampo(RS.Fields(1).Value, RS.Fields(0).Value) / 100
            If Porcen <> 0 Then
        
                Canaforo = Round2(DBLet(Rs1.Fields(0).Value, "N") * Porcen, 0)
                Hanegadas = Round2(Round2(DBLet(Rs1.Fields(1).Value, "N") / vParamAplic.Faneca, 2) * Porcen, 2)
                Hectareas = Round2(DBLet(Rs1.Fields(1).Value, "N") * Porcen, 4)
                Arboles = Round2(DBLet(Rs1.Fields(2).Value, "N") * Porcen, 0)
                
                Sql3 = Sql3 & DBSet(Canaforo, "N") & ","
                Sql3 = Sql3 & DBSet(Hanegadas, "N") & ","
                Sql3 = Sql3 & DBSet(Hectareas, "N") & ",0,0,"
                Sql3 = Sql3 & DBSet(Arboles, "N") & "),"
                
        
            Else
                ' si no hay coopropietarios es todo suyo
            
                Sql3 = Sql3 & DBSet(Rs1.Fields(0).Value, "N") & "," 'canaforo
                Ha = Round2(DBLet(Rs1.Fields(1).Value, "N") / vParamAplic.Faneca, 2)
                Sql3 = Sql3 & DBSet(Ha, "N") & "," 'hanegadas
                Sql3 = Sql3 & DBSet(Rs1.Fields(1).Value, "N") & ",0,0," 'hectareas
                Sql3 = Sql3 & DBSet(Rs1.Fields(2).Value, "N") 'arboles
                Sql3 = Sql3 & "),"
        
            End If
            
        
        Else
            Sql3 = Sql3 & "0,0,0,0,0,0),"
        End If
        
        Sql2 = Sql2 & Sql3
        
        Set Rs1 = Nothing
        
        
        RS.MoveNext
    Wend

    'quitamos la ultima coma
    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
    conn.Execute Sql2
    
    '[Monica]13/11/2013: puede que hayan errores en el prorrateo de hectareas, hanegadas, arboles y canaforo, se lo daremos al
    sql = "select codcampo, sum(canaforo) canaforo, sum(hanegada) hanegada, sum(hectarea) hectarea, sum(nroarbol) nroarbol from tmpinfkilos where codusu = " & vUsu.Codigo & " group by codcampo order by codcampo "
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        sql = "select codsocio, canaforo, "
        Select Case Combo1(5).ListIndex
            Case 0
                sql = sql & " supcoope, nroarbol"
            Case 1
                sql = sql & " supsigpa, nroarbol"
            Case 2
                sql = sql & " supcatas, nroarbol"
            Case 3
                sql = sql & " supculti, nroarbol"
        End Select
        sql = sql & " from rcampos where codcampo = " & DBSet(RS!codcampo, "N")
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs1.EOF Then
            DCanaforo = DBLet(RS!Canaforo, "N") - DBLet(Rs1!Canaforo, "N")
            DHectarea = DBLet(RS!hectarea, "N") - DBLet(Rs1.Fields(2).Value, "N")
            DHanegada = Round2(DHectarea / vParamAplic.Faneca, 2)
            DNroArbol = DBLet(RS!nroarbol, "N") - DBLet(Rs1!nroarbol, "N")
        
            sql = "update tmpinfkilos set "
            sql = sql & " canaforo = canaforo + " & DBSet(DCanaforo, "N")
            sql = sql & " ,hanegada = hanegada + " & DBSet(DHanegada, "N")
            sql = sql & " ,hectarea = hectarea + " & DBSet(DHectarea, "N")
            sql = sql & " ,nroarbol = nroarbol + " & DBSet(DNroArbol, "N")
            sql = sql & " where codusu = " & vUsu.Codigo
            sql = sql & " and codcampo = " & DBSet(RS!codcampo, "N")
            sql = sql & " and codsocio = " & DBSet(Rs1!Codsocio, "N")
        
            conn.Execute sql
        End If
        
        RS.MoveNext
    Wend
    
    CargarTemporal5 = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalVtaFruta(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim SqlValues As String
Dim Nombre As String
Dim AlbaranAnt As Long
Dim Primero As Boolean
Dim TipoAlb As Integer

    On Error GoTo eCargarTemporalVtaFruta
    
    CargarTemporalVtaFruta = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "select  vtafrutacab.numalbar, vtafrutacab.fecalbar, vtafrutacab.codsocio, vtafrutacab.codclien, vtafrutalin.codvarie, vtafrutalin.descalibre, vtafrutalin.numcajon, vtafrutalin.numpalet, vtafrutalin.pesonetoreal, vtafrutacab.numpalot, vtafrutacab.tarapalot, vtafrutacab.tipoalbaran FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SqlValues = ""
    
    If Not RS.EOF Then
        AlbaranAnt = DBLet(RS!numalbar, "N")
        Primero = True
    End If
    
    While Not RS.EOF
        SqlValues = SqlValues & "(" & vUsu.Codigo & ","
        
        If DBLet(RS.Fields(2).Value, "N") <> 0 Then 'es socio
            Nombre = DevuelveDesdeBDNew(cAgro, "rsocios", "nomsocio", "codsocio", DBLet(RS.Fields(2).Value, "N"), "N")
            SqlValues = SqlValues & DBLet(RS.Fields(2).Value, "N") & "," & DBSet(Nombre, "T") & ",0,"
        Else
            If DBLet(RS.Fields(3).Value, "N") <> 0 Then 'es cliente
                Nombre = DevuelveDesdeBDNew(cAgro, "clientes", "nomclien", "codclien", DBLet(RS.Fields(3).Value, "N"), "N")
                SqlValues = SqlValues & DBSet(RS.Fields(3).Value, "N") & "," & DBSet(Nombre, "T") & ",1,"
            End If
        End If
        
        
        SqlValues = SqlValues & DBSet(RS.Fields(1).Value, "F") & "," & DBSet(RS.Fields(4).Value, "N") & "," & DBSet(RS.Fields(0).Value, "N") & ","
        
        If AlbaranAnt = DBLet(RS!numalbar, "N") Then
            TipoAlb = DBLet(RS!tipoalbaran)
            If Primero Then
                SqlValues = SqlValues & DBSet(RS.Fields(5).Value, "T") & "," & DBSet(RS.Fields(6).Value, "N") & "," & DBSet(RS.Fields(7).Value, "N") & "," & DBSet(RS.Fields(8).Value, "N") & "," & DBSet(DBLet(RS.Fields(9).Value, "N"), "N") & "," & DBSet(DBLet(RS.Fields(10).Value, "N"), "N") & "," & DBSet(TipoAlb, "N") & "),"
                Primero = False
            Else
                SqlValues = SqlValues & DBSet(RS.Fields(5).Value, "T") & "," & DBSet(RS.Fields(6).Value, "N") & "," & DBSet(RS.Fields(7).Value, "N") & "," & DBSet(RS.Fields(8).Value, "N") & ",0,0," & DBSet(TipoAlb, "N") & "),"
            End If
        Else
            AlbaranAnt = DBLet(RS!numalbar, "N")
            TipoAlb = DBLet(RS!tipoalbaran)
            SqlValues = SqlValues & DBSet(RS.Fields(5).Value, "T") & "," & DBSet(RS.Fields(6).Value, "N") & "," & DBSet(RS.Fields(7).Value, "N") & "," & DBSet(RS.Fields(8).Value, "N") & "," & DBSet(DBLet(RS.Fields(9).Value, "N"), "N") & "," & DBSet(DBLet(RS.Fields(10).Value, "N"), "N") & "," & DBSet(TipoAlb, "N") & "),"
            Primero = False
        End If
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If SqlValues <> "" Then
        'quitamos la ultima coma
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        
        Sql2 = "insert into tmpinformes (codusu, codigo1, nombre1, campo1, fecha1, importe1, importe2, nombre2, importe3, importe4, importe5, importeb1, importeb2, importeb3) values " & SqlValues
    End If
    
    conn.Execute Sql2
    
    CargarTemporalVtaFruta = True
    Exit Function
    
eCargarTemporalVtaFruta:
    MuestraError "Cargando temporal Venta Fruta", Err.Description
End Function

Private Function ActualizarRegistros(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim sql As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False
'[Monica]20/12/2013: fallaba cuando metiamos desde/hasta clase
'    cTabla = QuitarCaracterACadena(cTabla, "{")
'    cTabla = QuitarCaracterACadena(cTabla, "}")
'    SQL = "update " & QuitarCaracterACadena(cTabla, "_1") & " set impreso = 1 "
    sql = "update rhisfruta, variedades set rhisfruta.impreso = 1 where rhisfruta.codvarie = variedades.codvarie "

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        'SQL = SQL & " WHERE " & cWhere
        sql = sql & " and " & cWhere
    End If
    
    conn.Execute sql
    
    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function


Private Function NombreCalidad(Var As String, Calid As String) As String
Dim sql As String

    NombreCalidad = ""

    sql = "select nomcalab from rcalidad where codvarie = " & DBSet(Var, "N")
    sql = sql & " and codcalid = " & DBSet(Calid, "N")
    
    NombreCalidad = DevuelveValor(sql)
    
End Function



Private Function GeneraFicheroAgriweb(pTabla As String, pWhere As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim vSocio As cSocio
Dim B As Boolean
Dim Nregs As Long
Dim Total As Variant

Dim cTabla As String
Dim vWhere As String


    On Error GoTo EGen
    GeneraFicheroAgriweb = False
    
    cTabla = pTabla
    vWhere = pWhere
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = QuitarCaracterACadena(cTabla, "_1")
    If vWhere <> "" Then
        vWhere = QuitarCaracterACadena(vWhere, "{")
        vWhere = QuitarCaracterACadena(vWhere, "}")
        vWhere = QuitarCaracterACadena(vWhere, "_1")
    End If
    
    NFic = FreeFile
    
    Open App.Path & "\fichero.txt" For Output As #NFic
    
    'CABECERA
    CabeceraAgriweb NFic
    
    Set RS = Nothing
    
    'Imprimimos las lineas
    Aux = "select  rcampos.codsocio, sum(rcampos.supsigpa) "
    Aux = Aux & " from " & cTabla
    Aux = Aux & " where " & vWhere
    Aux = Aux & " group by 1 "
    Aux = Aux & " order by 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then
        'No hayningun registro
    Else
        B = True
        Regs = 0
        While Not RS.EOF And B
            Regs = Regs + 1
            Set vSocio = New cSocio
            
            If vSocio.LeerDatos(DBLet(RS!Codsocio, "N")) Then
                LineaAgriweb NFic, vSocio, RS
            Else
                B = False
            End If
            
            Set vSocio = Nothing
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
            
    Close (NFic)
    
    If Regs > 0 Then GeneraFicheroAgriweb = True
    Exit Function
    
EGen:
    Set RS = Nothing
    Close (NFic)
    MuestraError Err.Number, Err.Description
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

Private Sub CabeceraAgriweb(NFich As Integer)
Dim cad As String
      
    ' multibase
    'column  0 ,"C",";",        column  2 ,f_campa,";",
    'column  7 ,"17",";",       column  10,cifempre[1,9],";",
    'column  20,"OP",";",       column  23,f_cifemp,";",
    'column  33,f_produc,";",   column  36,f_kilost using "&&&&&&&&&&",";",
    'column  47,f_fecont using "ddmmyyyy",";",       column  56,f_superf using "&&&&&&",";",
    'column  63,f_precio using "&&.&&",";",

    cad = "C"                                                  'p.1 tipo de registro
    cad = cad & Format(txtCodigo(27).Text, "0000")             'p.2 a�o ejercicio
    cad = cad & "17"                                           'p.6 comunidad autonoma
    cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 9)    'p.8 cif empresa
    cad = cad & "OP"                                           'p.17 tipo de vendedor
    cad = cad & RellenaABlancos(txtCodigo(28).Text, True, 9)   'p.19 cif industria transformadora
    cad = cad & RellenaABlancos(Combo1(4).Text, True, 2)       'p.28 producto segun tabla
    cad = cad & RellenaAceros(ImporteSinFormato(txtCodigo(29).Text), True, 10)    'p.30 kilos contratados
    cad = cad & Format(txtCodigo(30).Text, "ddmmyyyy")         'p.40 fecha de contratacion
    cad = cad & RellenaAceros(ImporteSinFormato(CCur(txtCodigo(31).Text) * 100), False, 6)    'p.48 superficie
    cad = cad & Format(txtCodigo(32).Text, "00.00")            'p.54 precio
    
    Print #NFich, cad
End Sub

Private Sub LineaAgriweb(NFich As Integer, vSocio As cSocio, ByRef RS As ADODB.Recordset)
Dim cad As String
Dim Areas As Long

    cad = "P"                                                'p.1 tipo de registro
    cad = cad & Format(txtCodigo(27).Text, "0000")           'p.2 a�o ejercicio
    cad = cad & "17"                                         'p.6 comunidad autonoma
    cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 9)  'p.8 cif empresa
    cad = cad & "OP"                                         'p.17 tipo de vendedor
    cad = cad & RellenaABlancos(txtCodigo(28).Text, True, 9) 'p.19 cif de la empresa transformadora
    cad = cad & RellenaABlancos(Combo1(4).Text, True, 2)     'p.28 codigo del producto
    cad = cad & RellenaABlancos(vSocio.Nombre, True, 40)     'p.30 nombre socio
    cad = cad & RellenaABlancos(vSocio.nif, True, 9)         'p.70 nif socio
    
    ' modificacion de Alzira (no es lo mismo socio que tercero)
    ' si es socio PA el resto es PI
    If vSocio.TipoProd = 0 Then
        cad = cad & "PA"                                         'p.79 tipo productor
    Else
        cad = cad & "PI"
    End If
    
    cad = cad & RellenaAceros(ImporteSinFormato(CStr(Round2(DBLet(RS.Fields(1).Value, "N") * 100, 0))), False, 6)   'p.81 superficie amparada
    
    Print #NFich, cad
End Sub

Private Function ProductoCampo(campo As String) As String
Dim sql As String

    ProductoCampo = ""
    
    sql = "select variedades.codprodu from rcampos inner join variedades on rcampos.codvarie = variedades.codvarie "
    sql = sql & " where rcampos.codcampo = " & DBSet(campo, "N")
    
    ProductoCampo = DevuelveValor(sql)

End Function





Private Function ComprobarErrores() As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim sql As String
Dim Sql1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim Mens As String
Dim Tipo As Integer


    On Error GoTo eComprobarErrores

    ComprobarErrores = False
    
    sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute sql
    

    I = 0
    lblProgres(2).Caption = "Comprobando errores Tabla temporal entradas "
    
    sql = "select count(*) from tmpentrada"
    longitud = TotalRegistros(sql)

    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0
    
    
    sql = "select * from tmpentrada"
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText


    B = True
    I = 0
    While Not RS.EOF And B
        I = I + 1

        Me.pb2.Value = Me.pb2.Value + 1
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh

        ' comprobamos que no exista el albaran en rclasifica
        sql = "select count(*) from rclasifica where numnotac = " & DBSet(RS!numalbar, "N")
        If TotalRegistros(sql) > 0 Then
            Mens = "Nro. de Nota ya existe"
            sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Albar�n:" & RS!numalbar, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute sql
        End If

        ' comprobamos que no exista el albaran en el historico
        sql = "select numalbar from rhisfruta_entradas where numnotac = " & DBSet(RS!numalbar, "N")
        If DevuelveValor(sql) <> 0 Then
            Mens = "Nro.Nota existe en hco. albar�n:" & DevuelveValor(sql)
            sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Albar�n:" & RS!numalbar, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute sql
        End If



        ' comprobamos que exista el socio
        sql = "select count(*) from rsocios where codsocio = " & DBSet(RS!Codsocio, "N")
        If TotalRegistros(sql) = 0 Then
            Mens = "Socio no existe"
            sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Socio:" & RS!Codsocio, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute sql
        End If

        ' comprobamos que exista la variedad
        sql = "select count(*) from variedades where codvarie = " & DBSet(RS!codvarie, "N")
        If TotalRegistros(sql) = 0 Then
            Mens = "Variedad no existe"
            sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Variedad:" & RS!codvarie, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute sql
        End If

        ' comprobamos que exista el campo
        sql = "select count(*) from rcampos where codsocio = " & DBSet(RS!Codsocio, "N")
        sql = sql & " and nrocampo = " & DBSet(RS!codcampo, "N")
        sql = sql & " and codvarie = " & DBSet(RS!codvarie, "N")
        If TotalRegistros(sql) = 0 Then
            Mens = "Campo no existe"
            sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet(("Socio:" & RS!Codsocio & "-Campo:" & RS!codcampo) & "-Variedad:" & RS!codvarie, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute sql
        End If

        ' comprobamos que no exista mas de un campo con ese numero de orden campo (scampo.codcampo MB)
        sql = "select count(*) from rcampos where codsocio = " & DBSet(RS!Codsocio, "N")
        sql = sql & " and nrocampo = " & DBSet(RS!codcampo, "N")
        sql = sql & " and codvarie = " & DBSet(RS!codvarie, "N")
        If TotalRegistros(sql) > 1 Then
            Mens = "Campo con m�s de un registro"
            sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet(("Socio:" & RS!Codsocio & "-Campo:" & RS!codcampo) & "-Variedad:" & RS!codvarie, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute sql
        End If

' se activara cuando quite botones de lineas de la clasificacion

'        '20-05-2009: comprobamos que si no tiene clasificacion tenga en campo o en almacen
'        SQL = "select count(*) from tmpclasific where numalbar = " & DBSet(Rs!numalbar, "N")
'
'        If TotalRegistros(SQL) = 0 Then
'            SQL = "select tipoclasifica from variedades where codvarie = " & DBSet(Rs!CodVarie, "N")
'            Tipo = DevuelveValor(SQL)
'            If Tipo = 0 Then ' es por campo
'                SQL = "select count(*) from rcampos_clasif, rcampos where rcampos.nrocampo = " & DBSet(Rs!CodCampo, "N")
'                SQL = SQL & " and rcampos.codcampo= rcampos_clasif.codcampo and rcampos.codvarie = " & DBSet(Rs!CodVarie, "N")
'                SQL = SQL & " and rcampos.codsocio = " & DBSet(Rs!CodSocio, "N")
'
'                If TotalRegistros(SQL) = 0 Then
'                    Mens = "Campo sin clasificaci�n "
'                    SQL = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
'                          vUsu.Codigo & "," & DBSet(("Nro.Campo:" & Rs!CodCampo) & "-Variedad:" & Rs!CodVarie, "T") & "," & DBSet(Mens, "T") & ")"
'                    conn.Execute SQL
'                End If
'            Else ' es en almacen
'                SQL = "select count(*) from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")
'                If TotalRegistros(SQL) = 0 Then
'                    Mens = "Variedad sin calidades para clasificaci�n "
'                    SQL = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
'                          vUsu.Codigo & "," & DBSet(("Nro.Campo:" & Rs!CodCampo) & "-Variedad:" & Rs!CodVarie, "T") & "," & DBSet(Mens, "T") & ")"
'                    conn.Execute SQL
'                End If
'            End If
'        End If

        RS.MoveNext
    Wend
    Set RS = Nothing
    
    
    lblProgres(2).Caption = "Comprobando errores Tabla temporal clasifica "
    
    sql = "select count(*) from tmpclasific"
    longitud = TotalRegistros(sql)

    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0

    sql = "select * from tmpclasific"
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    B = True
    I = 0
    While Not RS.EOF And B
        I = I + 1

        Me.pb2.Value = Me.pb2.Value + 1
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh

        ' comprobamos que no exista el albaran en rclasifica
        sql = "select count(*) from rclasifica where numnotac = " & DBSet(RS!numalbar, "N")
        If TotalRegistros(sql) > 0 Then
            Mens = "Nro. de Nota ya existe"
            sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Albar�n:" & RS!numalbar, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute sql
        End If

        ' comprobamos que exista la variedad
        sql = "select count(*) from variedades where codvarie = " & DBSet(RS!codvarie, "N")
        If TotalRegistros(sql) = 0 Then
            Mens = "Variedad no existe"
            sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Variedad:" & RS!codvarie, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute sql
        End If

        ' comprobamos que exista la calidad
        sql = "select count(*) from rcalidad where codvarie = " & DBSet(RS!codvarie, "N")
        sql = sql & " and codcalid = " & DBSet(RS!codcalir, "N")
        If TotalRegistros(sql) = 0 Then
            Mens = "Calidad no existe"
            sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet(("Variedad:" & RS!codvarie & "-Calidad:" & RS!codcalir), "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute sql
        End If


        RS.MoveNext
    Wend
    Set RS = Nothing
    

    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    ComprobarErrores = B
    Exit Function

eComprobarErrores:
    ComprobarErrores = False
End Function



Private Function CargarTablasTemporales(nomFich1 As String, nomFich2 As String) As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim Variedad As String
Dim HoraEntrada As String

Dim Sql3 As String
Dim campo As String

    On Error GoTo eCargarTablasTemporales
    
    CargarTablasTemporales = False
    
    ' cargando tabla temporal primera
    NF = FreeFile
    Open nomFich2 For Input As #NF
    
    cad = ""
    
    Line Input #NF, cad
    I = 0
    
    lblProgres(2).Caption = "Cargando Tabla temporal: Entradas"
    longitud = FileLen(nomFich2)
    
    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0

    sql = "insert into tmpentrada(codsocio, codcampo, numalbar, codvarie, fecalbar, "
    sql = sql & "horalbar, kilosbru, kilosnet, numcajon) values  "
    Sql2 = ""

    While Not EOF(NF)
        I = I + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        Variedad = Format(RecuperaValor(cad, 4), "00") & Format(RecuperaValor(cad, 5), "00")
        HoraEntrada = DBSet(RecuperaValor(cad, 6) & " " & RecuperaValor(cad, 7), "FH")
        
'        Sql3 = "select codcampo from rcampos where codsocio = " & DBSet(RecuperaValor(cad, 1), "N") ' socio
'        Sql3 = Sql3 & " and codvarie = " & DBSet(Variedad, "N")     ' variedad
'        Campo = DevuelveValor(Sql3)
        
        Sql2 = Sql2 & "(" & DBSet(RecuperaValor(cad, 1), "N") & ","    ' socio
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 2), "N") & ","   ' campo codigo de campo MB
'        Sql2 = Sql2 & DBSet(Campo, "N") & "," ' campo
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 3), "N") & ","   ' albaran
        Sql2 = Sql2 & DBSet(Variedad, "N") & ","                ' variedad
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 6), "F") & ","   ' fecha entrada
        Sql2 = Sql2 & HoraEntrada & ","            ' hora de entrada
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 8), "N") & ","   ' kilos brutos
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 9), "N") & ","   ' kilos netos
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 10), "N") & ")," ' numero de cajones
        
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then ' falta la ultima linea
        Variedad = Format(RecuperaValor(cad, 4), "00") & Format(RecuperaValor(cad, 5), "00")
        HoraEntrada = DBSet(RecuperaValor(cad, 6) & " " & RecuperaValor(cad, 7), "FH")
        
'        Sql3 = "select codcampo from rcampos where codsocio = " & DBSet(RecuperaValor(cad, 1), "N") ' socio
'        Sql3 = Sql3 & " and codvarie = " & DBSet(Variedad, "N")     ' variedad
'        Campo = DevuelveValor(Sql3)
        
        Sql2 = Sql2 & "(" & DBSet(RecuperaValor(cad, 1), "N") & ","    ' socio
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 2), "N") & ","   ' campo codigo de campo MB
'        Sql2 = Sql2 & DBSet(Campo, "N") & "," ' campo
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 3), "N") & ","   ' albaran
        Sql2 = Sql2 & DBSet(Variedad, "N") & ","                ' variedad
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 6), "F") & ","   ' fecha entrada
        Sql2 = Sql2 & HoraEntrada & ","            ' hora de entrada
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 8), "N") & ","   ' kilos brutos
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 9), "N") & ","   ' kilos netos
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 10), "N") & ")," ' numero de cajones
    End If
    
    sql = sql & Mid(Sql2, 1, Len(Sql2) - 1)
    conn.Execute sql
    
    
    
    ' clasificacion
    
    NF = FreeFile
    Open nomFich1 For Input As #NF
    
    cad = ""
    
    Line Input #NF, cad
    I = 0
    
    lblProgres(2).Caption = "Cargando Tabla temporal: Clasificacion"
    longitud = FileLen(nomFich1)
    
    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0

    sql = "insert into tmpclasific(numalbar, codvarie, codcalir, porcenta) values  "
    Sql2 = ""
    
    While Not EOF(NF)
        I = I + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        Variedad = Format(RecuperaValor(cad, 2), "00") & Format(RecuperaValor(cad, 3), "00")
        
        
        Sql2 = Sql2 & "(" & DBSet(RecuperaValor(cad, 1), "N") & ","    ' numero de albaran
        Sql2 = Sql2 & DBSet(Variedad, "N") & ","                ' variedad
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 4), "N") & ","   ' calidad
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 5), "N") & "),"  ' porcentaje
        
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        I = I + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        Variedad = Format(RecuperaValor(cad, 2), "00") & Format(RecuperaValor(cad, 3), "00")
        
        
        Sql2 = Sql2 & "(" & DBSet(RecuperaValor(cad, 1), "N") & ","    ' numero de albaran
        Sql2 = Sql2 & DBSet(Variedad, "N") & ","                ' variedad
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 4), "N") & ","   ' calidad
        Sql2 = Sql2 & DBSet(RecuperaValor(cad, 5), "N") & "),"  ' porcentaje
    
    
    End If
    
    
    sql = sql & Mid(Sql2, 1, Len(Sql2) - 1)
    conn.Execute sql
    
    
    
    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    CargarTablasTemporales = True
    Exit Function

eCargarTablasTemporales:
    CargarTablasTemporales = False
End Function


Private Function CargarClasificacion() As Boolean
Dim sql As String
Dim Sql1 As String
Dim Sql3 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Precio As Currency
Dim Transporte As Currency
Dim Kilos As Long

Dim AlbarAnt As Long
Dim KilosAlbar As Long
Dim KilosNetAnt As Long
Dim VarieAnt As Long
Dim CalidAnt As Long
Dim longitud As Long

Dim campo As Variant
Dim cadMen As String


    On Error GoTo eCargarClasificacion
    
    CargarClasificacion = False
    
    
    lblProgres(2).Caption = "Cargando Entradas"
    
    sql = "select count(*) from tmpentrada order by numalbar"
    longitud = TotalRegistros(sql)
    
    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0
    
    
    sql = "select * from tmpentrada order by numalbar"
    
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Me.pb2.Value = Me.pb2.Value + 1
        lblProgres(3).Caption = "Albar�n " & DBLet(RS!numalbar, "N")
        Me.Refresh
        
        
        Transporte = 0
    
        sql = "insert into rclasifica (numnotac,fechaent,horaentr,codvarie,codsocio,codcampo,tipoentr,recolect,codtrans,codcapat,"
        sql = sql & "codtarif,kilosbru,numcajon,kilosnet,observac,"
        sql = sql & "imptrans,impacarr,imprecol,imppenal,tiporecol,horastra,numtraba,numalbar,fecalbar,impreso) values "
    
        campo = 0
        campo = DevuelveValor("select codcampo from rcampos where nrocampo = " & DBSet(RS!codcampo, "N") & " and codsocio=" & DBSet(RS!Codsocio, "N") & " and codvarie=" & DBSet(RS!codvarie, "N"))
    
        sql = sql & "(" & DBSet(RS!numalbar, "N") & ","
        sql = sql & DBSet(RS!Fecalbar, "F") & ","
        sql = sql & DBSet(RS!horalbar, "FH") & ","
        sql = sql & DBSet(RS!codvarie, "N") & ","
        sql = sql & DBSet(RS!Codsocio, "N") & ","
'        Sql = Sql & DBSet(Rs!codCampo, "N") & ","
        sql = sql & DBSet(campo, "N") & ","
        sql = sql & "0," ' tipoentr 0=normal
        sql = sql & "1," ' recolect 1=socio
        sql = sql & ValorNulo & "," 'transportista
        sql = sql & ValorNulo & "," 'capataz
        sql = sql & ValorNulo & "," 'tarifa
        sql = sql & DBSet(RS!KilosBru, "N") & ","
        sql = sql & DBSet(RS!Numcajon, "N") & ","
        sql = sql & DBSet(RS!KilosNet, "N") & ","
        sql = sql & ValorNulo & ","
        sql = sql & DBSet(Transporte, "N") & ","
        sql = sql & ValorNulo & ","
        sql = sql & ValorNulo & ","
        sql = sql & ValorNulo & ","
        sql = sql & "0," 'tiporecol 0=horas 1=destajo no admite valor nulo
        sql = sql & ValorNulo & ","
        sql = sql & ValorNulo & ","
        sql = sql & DBSet(RS!numalbar, "N") & ","
        sql = sql & DBSet(RS!Fecalbar, "F") & ",0)"
        
        conn.Execute sql
        
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing


    ' 21-05-2009: cargamos las clasificacion dependiendo de si es por campo o almacen de aquellas que
    ' no tengan clasificacion
    sql = "select * from tmpentrada order by numalbar "

    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        sql = "select count(*) from tmpclasific where numalbar = " & DBSet(RS!numalbar, "N")
        If TotalRegistros(sql) = 0 Then ' si no hay clasificacion en el fichero metemos la correspondiente
            Tipo = DevuelveValor("select tipoclasifica from variedades where codvarie = " & DBSet(RS!codvarie, "N"))
            If Tipo = 0 Then ' clasificacion en campo
                campo = 0
                campo = DevuelveValor("select codcampo from rcampos where nrocampo = " & DBSet(RS!codcampo, "N") & " and codsocio=" & DBSet(RS!Codsocio, "N") & " and codvarie=" & DBSet(RS!codvarie, "N"))

                sql = "insert into tmpclasific (numalbar, codvarie, codcalir, porcenta) "
                sql = sql & " select " & DBSet(RS!numalbar, "N") & ", codvarie, codcalid, muestra "
                sql = sql & " from rcampos_clasif where codcampo = " & DBSet(campo, "N")

                conn.Execute sql
            Else ' clasificacion en almacen
                sql = "insert into tmpclasific (numalbar, codvarie, codcalir, porcenta) "
                sql = sql & " select " & DBSet(RS!numalbar, "N") & ", codvarie, codcalid, 0 "
                sql = sql & " from rcalidad where codvarie = " & DBSet(RS!codvarie, "N")

                conn.Execute sql
            End If
        End If
        RS.MoveNext
    Wend
    Set RS = Nothing

    ' 21-05-2009
    
    lblProgres(2).Caption = "Cargando Clasificaci�n"
    
    sql = "select count(*) from tmpclasific, tmpentrada "
    sql = sql & " where tmpclasific.numalbar=tmpentrada.numalbar "
    longitud = TotalRegistros(sql)
    
    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0
    
    
    sql = "select *, tmpentrada.kilosnet as kilosent from tmpclasific, tmpentrada "
    sql = sql & " where tmpclasific.numalbar=tmpentrada.numalbar "
    sql = sql & " order by tmpclasific.numalbar, tmpclasific.codcalir"
    
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        AlbarAnt = DBLet(RS!numalbar, "N")
        KilosNetAnt = DBLet(RS!Kilosent, "N")
        VarieAnt = DBLet(RS!codvarie, "N")
        CalidAnt = DBLet(RS!codcalir, "N")
    End If
        
    KilosAlbar = 0
    While Not RS.EOF
        
        Me.pb2.Value = Me.pb2.Value + 1
        lblProgres(3).Caption = "Albar�n " & DBLet(RS!numalbar, "N") & " Variedad " & DBLet(RS!codvarie, "N") & " Calidad " & DBLet(RS!codcalir, "N")
        Me.Refresh
        
        Kilos = Round2(DBLet(RS!Kilosent, "N") * DBLet(RS!porcenta, "N") / 100, 0)
        
        If AlbarAnt <> DBLet(RS!numalbar, "N") Then
            If KilosNetAnt <> KilosAlbar Then
                Sql3 = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNetAnt - KilosAlbar, "N")
                Sql3 = Sql3 & " where numnotac = " & DBSet(AlbarAnt, "N")
                Sql3 = Sql3 & " and codvarie = " & DBSet(VarieAnt, "N")
                Sql3 = Sql3 & " and codcalid = " & DBSet(CalidAnt, "N")
            
                conn.Execute Sql3
            End If
        
            KilosAlbar = Kilos
            KilosNetAnt = DBLet(RS!Kilosent, "N")
            
            AlbarAnt = DBLet(RS!numalbar, "N")
        Else
            KilosAlbar = KilosAlbar + Kilos
        End If
    
        VarieAnt = DBLet(RS!codvarie, "N")
        CalidAnt = DBLet(RS!codcalir, "N")
        
        
        sql = "insert into rclasifica_clasif (numnotac,codvarie,codcalid,muestra,kilosnet) values"
        sql = sql & "(" & DBSet(RS!numalbar, "N") & ","
        sql = sql & DBSet(RS!codvarie, "N") & ","
        sql = sql & DBSet(RS!codcalir, "N") & ","
        sql = sql & DBSet(RS!porcenta, "N") & ","
        sql = sql & DBSet(Kilos, "N") & ")"
        
        conn.Execute sql
        
        RS.MoveNext
    Wend
    
    ' si la clasificacion es diferente actualizamos en la ultima calidad
    If KilosNetAnt <> KilosAlbar Then
        Sql3 = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNetAnt - KilosAlbar, "N")
        Sql3 = Sql3 & " where numnotac = " & DBSet(AlbarAnt, "N")
        Sql3 = Sql3 & " and codvarie = " & DBSet(VarieAnt, "N")
        Sql3 = Sql3 & " and codcalid = " & DBSet(CalidAnt, "N")
    
        conn.Execute Sql3
    End If
    
    Set RS = Nothing
    
    sql = "select rclasifica.* from rclasifica, tmpentrada where rclasifica.numnotac = tmpentrada.numalbar "
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not ActualizarTransporte(RS, cadMen) Then
            cadMen = "Actualizando gastos de transporte" & cadMen
            MsgBox cadMen, vbExclamation
            Set RS = Nothing
            
            pb2.visible = False
            lblProgres(2).Caption = ""
            lblProgres(3).Caption = ""
        
            CargarClasificacion = False
            Exit Function
        End If
    End If
    
    Set RS = Nothing
    
    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    CargarClasificacion = True
    Exit Function
    
eCargarClasificacion:
    MuestraError Err.Number, "Cargar clasificaci�n", Err.Description
End Function



Private Function ExistenFicheros() As Boolean
Dim b1 As Boolean
Dim b2 As Boolean
Dim cadMen As String

    On Error GoTo eExistenFicheros


    ExistenFicheros = False
    b1 = False
    b2 = False
    
    
    If Dir(vParamAplic.PathTraza, vbDirectory) = "" Or vParamAplic.PathTraza = "" Then
        cadMen = "La carpeta de los ficheros de traza " & vParamAplic.PathTraza & " de par�metros no existe. Revise."
        MsgBox cadMen, vbExclamation
        ExistenFicheros = False
        Exit Function
    End If
    
    cadMen = "Los Ficheros : " & vbCrLf
    
    If Dir(vParamAplic.PathTraza & "\clasific.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        clasific.txt"
        b1 = True
    End If
    If Dir(vParamAplic.PathTraza & "\entrada.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        entrada.txt"
        b2 = True
    End If
    
    If Not (b1 And b2) Then
        cadMen = cadMen & vbCrLf & vbCrLf & "no existen en el directorio de traza. Revise." & vbCrLf
        MsgBox cadMen, vbExclamation
    End If
    ExistenFicheros = (b1 And b2)
    Exit Function
    
eExistenFicheros:
    MuestraError Err.Number, "Error en Existen ficheros"
End Function


Private Function ActualizarTransporte(RS As ADODB.Recordset, cadErr As String) As Boolean
Dim Sql1 As String
Dim Rs2 As ADODB.Recordset
Dim KilosDestrio As Currency
Dim Precio As Currency
Dim Transporte As Currency
Dim Kilos As Currency


    On Error GoTo eActualizarTransporte

    If Not RS.EOF Then RS.MoveFirst
    While Not RS.EOF
        Sql1 = "select imptrans from rportespobla, rpartida, rcampos, variedades "
        Sql1 = Sql1 & " where rpartida.codparti = rcampos.codparti and "
        Sql1 = Sql1 & " variedades.codprodu = rportespobla.codprodu and "
        Sql1 = Sql1 & " rpartida.codpobla = rportespobla.codpobla and "
        Sql1 = Sql1 & " variedades.codvarie = " & DBSet(RS!codvarie, "N") & " and "
        Sql1 = Sql1 & " rcampos.codcampo = " & DBSet(RS!codcampo, "N") & " and "
        Sql1 = Sql1 & " rcampos.codvarie = variedades.codvarie "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Precio = 0
        If Not Rs2.EOF Then
            Precio = DBLet(Rs2.Fields(0).Value, "N")
        End If
        
        Set Rs2 = Nothing
        
        ' cogemos los kilos de la clasificacion que sean de destrio
        Sql1 = "select kilosnet from rclasifica_clasif, rcalidad where numnotac = " & DBSet(RS!numnotac, "N")
        Sql1 = Sql1 & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
        Sql1 = Sql1 & " and rclasifica_clasif.codcalid = rcalidad.codcalid  "
        Sql1 = Sql1 & " and rcalidad.tipcalid = 1 "
        KilosDestrio = DevuelveValor(Sql1)
        
        
        ' los gastos de transporte se calculan sobre los kilosnetos - los de destrio
        Kilos = DBLet(RS!KilosNet, "N") - KilosDestrio
        Transporte = Round2(Kilos * Precio, 2)
        
        Sql1 = "update rclasifica set imptrans = " & DBSet(Transporte, "N")
        Sql1 = Sql1 & " where numnotac = " & DBSet(RS!numnotac, "N")
        conn.Execute Sql1
        
        RS.MoveNext
    Wend
    
eActualizarTransporte:
    If Err.Number <> 0 Then
        ActualizarTransporte = False
        cadErr = Err.Description
    Else
        ActualizarTransporte = True
    End If
End Function

Private Function GeneraFicheroTraspasoCoop(pTabla As String, pWhere As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim vSocio As cSocio
Dim B As Boolean
Dim Nregs As Long
Dim Total As Variant

Dim cTabla As String
Dim vWhere As String


    On Error GoTo EGen
    GeneraFicheroTraspasoCoop = False
    
    cTabla = pTabla
    vWhere = pWhere
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = QuitarCaracterACadena(cTabla, "_1")
    If vWhere <> "" Then
        vWhere = QuitarCaracterACadena(vWhere, "{")
        vWhere = QuitarCaracterACadena(vWhere, "}")
        vWhere = QuitarCaracterACadena(vWhere, "_1")
    End If
    
    NFic = FreeFile
    
    Open App.Path & "\trascoop.txt" For Output As #NFic
    
    Set RS = Nothing
    
    'Imprimimos las lineas
    Aux = "select  rfactsoc.* "
    Aux = Aux & " from " & cTabla
    Aux = Aux & " where " & vWhere
    
    Set RS = New ADODB.Recordset
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then
        'No hayningun registro
    Else
        B = True
        Regs = 0
        While Not RS.EOF And B
            Regs = Regs + 1

            B = LineaTraspasoCoop(NFic, txtCodigo(45).Text, RS)
            
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
            
    Close (NFic)
    
    If Regs > 0 And B Then GeneraFicheroTraspasoCoop = True
    Exit Function
    
EGen:
    Set RS = Nothing
    Close (NFic)
    MuestraError Err.Number, Err.Description
End Function


Private Function GeneraFicheroTraspasoROPAS(pTabla As String, pWhere As String, pTabla1 As String, pWhere1 As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim vSocio As cSocio
Dim B As Boolean
Dim Nregs As Long
Dim Total As Variant

Dim cTabla As String
Dim vWhere As String

Dim Lin As Integer

Dim AntSocio As Long
Dim AntPoligono As Long
Dim AntParcela As Long

Dim FechaEnvio As String


    On Error GoTo EGen
    GeneraFicheroTraspasoROPAS = False
    
    cTabla = pTabla
    vWhere = pWhere
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = QuitarCaracterACadena(cTabla, "_1")
    If vWhere <> "" Then
        vWhere = QuitarCaracterACadena(vWhere, "{")
        vWhere = QuitarCaracterACadena(vWhere, "}")
        vWhere = QuitarCaracterACadena(vWhere, "_1")
    End If
    
    NFic = FreeFile
    
    FechaEnvio = Mid(txtCodigo(132).Text, 7, 4) & Mid(txtCodigo(132).Text, 4, 2) & Mid(txtCodigo(132).Text, 1, 2)
    Open App.Path & "\Socios_" & Format(txtCodigo(62).Text, "0000") & "_" & FechaEnvio & "_" & vParam.CifEmpresa & ".csv" For Output As #NFic
    
    Set RS = Nothing
    
    'Imprimimos las lineas
    Aux = "select  rsocios.*, rsocios_seccion.fecalta, rsocios_seccion.fecbaja" ', rsocios_seccion.* "
    Aux = Aux & " from " & cTabla
    If vWhere <> "" Then
        Aux = Aux & " where " & vWhere
    End If
    Aux = Aux & " order by rsocios.codsocio "
    
    pb7.Max = TotalRegistrosConsulta(Aux)
    pb7.visible = True
    Label2(187).visible = True
    Label2(187).Caption = "Cargando Socios"
    pb7.Value = 0
    
    Set RS = New ADODB.Recordset
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    If RS.EOF Then
        'No hayningun registro
    Else
            
        '[Monica]27/11/2012: Introducimos la cabecera
        cad = "Ejercicio; CifOpfh; Cif; Dni; NSocio; NombreSocio; Pais; TipoSocio; FAlta; FBaja"
        Print #NFic, cad
    
        B = True
        Regs = 0
        While Not RS.EOF And B
            IncrementarProgresNew pb7, 1
            DoEvents
            
            Regs = Regs + 1

            B = LineaTraspasoSocioROPAS(NFic, RS)
            
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
            
    Close (NFic)
    
    ' traspaso de campos de seccion horto
    If B Then
    
        cTabla = pTabla1
        vWhere = pWhere1
        
        cTabla = QuitarCaracterACadena(cTabla, "{")
        cTabla = QuitarCaracterACadena(cTabla, "}")
        cTabla = QuitarCaracterACadena(cTabla, "_1")
        If vWhere <> "" Then
            vWhere = QuitarCaracterACadena(vWhere, "{")
            vWhere = QuitarCaracterACadena(vWhere, "}")
            vWhere = QuitarCaracterACadena(vWhere, "_1")
        End If
        
        NFic = FreeFile
        
        
        Open App.Path & "\Parcela_" & Format(txtCodigo(62).Text, "0000") & "_" & FechaEnvio & "_" & vParam.CifEmpresa & ".csv" For Output As #NFic
        
        Set RS = Nothing
        
        '[Monica]14/02/2013: El fichero de campos se graba diferente para Picassent
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
            
            B = GeneracionFicheroCamposPicassent(NFic, cTabla, vWhere, Regs)
        
        Else
                
            B = GeneracionFicheroCampos(NFic, cTabla, vWhere, Regs)
        
        End If
        
        
    End If
    
    If Regs > 0 And B Then GeneraFicheroTraspasoROPAS = True
    Set RS = Nothing
    Close (NFic)
    pb7.visible = False
    Label2(187).visible = False
    DoEvents
    
    Exit Function
    
EGen:
    Set RS = Nothing
    Close (NFic)
    pb7.visible = False
    Label2(187).visible = False
    DoEvents
    MuestraError Err.Number, Err.Description
End Function


Private Function GeneracionFicheroCampos(NFic As Integer, cTabla As String, vWhere As String, Regs As Integer) As Boolean
Dim Aux As String
Dim RS As ADODB.Recordset
Dim B As Boolean
Dim Lin As Integer

Dim AntSocio As Long
Dim AntPoligono As Long
Dim AntParcela As Long

Dim cad As String


    On Error GoTo eGeneracionFicheroCampos

    GeneracionFicheroCampos = False

    Aux = "select rcampos.codsocio, rcampos.codvarie, rsocios.nifsocio, rcampos.poligono,  "
    Aux = Aux & " rcampos.parcela, rcampos.subparce, rcampos.codparti, rcampos.supsigpa, "
    Aux = Aux & " rcampos.recintos, rcampos.supcoope, rcampos.canaforo, rcampos.fecaltas, "
    Aux = Aux & " rcampos.fecbajas, rcampos.supcatas, rsocios_seccion.fecalta, rcampos.codcampo, rcampos.tipoparc, rcampos.refercatas  "
    Aux = Aux & " from " & cTabla
    If vWhere <> "" Then
        Aux = Aux & " where " & vWhere
    End If
    Aux = Aux & " order by rcampos.codsocio, rcampos.poligono, rcampos.parcela, "
    Aux = Aux & " rcampos.subparce, rcampos.recintos, rcampos.codvarie"
    
    Set RS = New ADODB.Recordset
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then
        'No hayningun registro
    Else
        B = True
        Regs = 0
        Lin = 0
        
        '[Monica]27/11/2012: Introducimos la cabecera
        cad = "Ejercicio; CifOpfh; Dni; Pais; TipoParcela; CodParcela; Provincia; Municipio; Agregado; Zona; Poligono; Parcela; Recinto; SubRecinto; SupParcela; SupRecinto; SupSubRecinto; FAlta; FBaja; Cosecha; Producto;SupCultivo;Produccion"
        Print #NFic, cad
        
        
        If Not RS.EOF Then
            AntSocio = DBLet(RS!Codsocio, "N")
            AntPoligono = DBLet(RS!Poligono, "N")
            AntParcela = DBLet(RS!Parcela, "N")
        End If
        
        pb7.Max = TotalRegistrosConsulta(Aux)
        pb7.visible = True
        Label2(187).visible = True
        Label2(187).Caption = "Cargando Campos"
        pb7.Value = 0
        
        
        While Not RS.EOF And B
            IncrementarProgresNew pb7, 1
            DoEvents
            
            Regs = Regs + 1

            If AntSocio <> RS!Codsocio Or AntPoligono <> RS!Poligono Or AntParcela <> RS!Parcela Then
                Lin = 0
            End If
            Lin = Lin + 1

            B = LineaTraspasoCampoROPAS(NFic, RS, Lin)
            
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
            
    Close (NFic)
    
    GeneracionFicheroCampos = True
    Exit Function


eGeneracionFicheroCampos:
    MuestraError Err.Number, "Error en la Generacion de fichero de Campos.", Err.Description
End Function



Private Function LineaTraspasoCoop(NFich As Integer, Coop As String, ByRef RS As ADODB.Recordset) As Boolean
Dim cad As String
Dim Areas As Long
Dim Tipo As Integer
Dim sql As String
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim CodiIVA As String
Dim TipoIRPF As Byte
Dim PorcIva As String
Dim vPorcIva As Currency
Dim CoopSoc As Currency

Dim Producto As String
Dim Variedad As String
Dim NomVar As String
Dim codVar As Long

Dim nifSocio As String
Dim Kilos As Long
Dim vPorcGasto As String
Dim vImporte As Currency
Dim Gastos As Currency



    On Error GoTo eLineaTraspasoCoop

    LineaTraspasoCoop = False

    cad = ""
    
    sql = "select count(*) from rfactsoc_variedad where codtipom = " & DBSet(RS!CodTipom, "T")
    sql = sql & " and numfactu = " & DBSet(RS!numfactu, "N") & " and fecfactu = " & DBSet(RS!fecfactu, "F")
    
    If TotalRegistros(sql) > 1 Then
        Producto = "00"
        Variedad = "00"
        NomVar = "Varias Var."
    Else
        sql = "select rfactsoc_variedad.codvarie  from rfactsoc_variedad where codtipom = " & DBSet(RS!CodTipom, "T")
        sql = sql & " and numfactu = " & DBSet(RS!numfactu, "N") & " and fecfactu = " & DBSet(RS!fecfactu, "F")
        
        codVar = DevuelveValor(sql)
        
        Producto = Mid(Format(codVar, "0000"), 1, 2)
        Variedad = Mid(Format(codVar, "0000"), 3, 2)
        
        NomVar = DevuelveValor("select nomvarie from variedades where codvarie = " & DBSet(codVar, "N"))
    End If
    
    
    If CInt(Coop) = 1 Or CInt(Coop) = 3 Or CInt(Coop) = 4 Then
        Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(RS!CodTipom, "T"))
        Select Case Tipo
            Case 1, 3, 7, 9 'anticipo normales, almazara y bodega
                cad = "0|"
            Case 2, 4, 8, 10 'liquidacion normales, almazara y bodega
                cad = "1|"
            
        End Select
'        Producto = DevuelveValor("select codprodu from variedades where codvarie = " & DBSet(Rs!CodVarie, "N"))
        nifSocio = DevuelveValor("select nifsocio from rsocios where codsocio =" & DBSet(RS!Codsocio, "N"))
        
        sql = "select sum(kilosnet) from rfactsoc_variedad where codtipom = " & DBSet(RS!CodTipom, "T")
        sql = sql & " and numfactu = " & DBSet(RS!numfactu, "N") & " and fecfactu = " & DBSet(RS!fecfactu, "F")
        Kilos = DevuelveValor(sql)
        
        If CInt(Coop) = 3 Or CInt(Coop) = 4 Then
            cad = cad & Format(DBLet(RS!numfactu, "N"), "000000") & "|"
            cad = cad & Format(DBLet(RS!fecfactu, "F"), "yymmdd") & "|"
            cad = cad & Format(DBLet(RS!Codsocio, "N"), "0000") & "|"
            cad = cad & Format(DBLet(Producto, "N"), "00") & "|"
            cad = cad & Format(DBLet(Variedad, "N"), "00") & "|"
            cad = cad & RellenaABlancos(NomVar, True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(RS!baseimpo, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(RS!ImporIva, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(RS!TotalFac, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(RS!ImpReten, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(RS!impapor, "N"), "#######0.00"), True, 11) & "|"
            
        Else
            cad = cad & Format(DBLet(RS!numfactu, "N"), "000000") & "|"
            cad = cad & Format(DBLet(RS!fecfactu, "F"), "yymmdd") & "|"
            cad = cad & Format(DBLet(RS!Codsocio, "N"), "000000") & "|"
            cad = cad & Format(DBLet(Producto, "N"), "00") & "|"
            cad = cad & Format(DBLet(Variedad, "N"), "00") & "|"
            cad = cad & RellenaABlancos(NomVar, True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(RS!baseimpo, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(RS!ImporIva, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(RS!TotalFac, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(RS!ImpReten, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(RS!impapor, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(nifSocio, True, 9) & "|"
            cad = cad & Format(Kilos, "00000000") & "|"
            
        End If
    Else
        cad = cad & Format(DBLet(RS!numfactu, "N"), "0000000")
        cad = cad & Format(DBLet(RS!Codsocio, "N"), "0000000")
        cad = cad & Format(DBLet(RS!fecfactu, "F"), "yymmdd")
        cad = cad & RellenaABlancos(NomVar, True, 11)
        cad = cad & RellenaABlancos(Format(Abs(DBLet(RS!baseimpo, "N")), "00000.00"), True, 8)
        
        If DBLet(RS!baseimpo, "N") < 0 Then
            cad = cad & "-"
        Else
            cad = cad & "+"
        End If
        
        vPorcIva = Round2(DBLet(RS!porc_iva, "N") * 100, 0)
        
        cad = cad & Format(vPorcIva, "0000")
        cad = cad & "0000"
        cad = cad & Format(Abs(DBLet(RS!ImporIva, "N")), "000.00")
        
        If DBLet(RS!ImporIva, "N") < 0 Then
            cad = cad & "-"
        Else
            cad = cad & "+"
        End If
        
        ' total factura
        cad = cad & Format(Abs(DBLet(RS!TotalFac, "N")), "00000.00")
        
        If DBLet(RS!TotalFac, "N") < 0 Then
            cad = cad & "-"
        Else
            cad = cad & "+"
        End If
        
        cad = cad & "00000000"
        
        ' base de retencion
        If DBLet(RS!BaseReten, "N") = 0 Then
            cad = cad & "00000000+"
        Else
            If DBLet(RS!BaseReten, "N") < 0 Then
                cad = cad & Format(Abs(DBLet(RS!BaseReten, "N")), "00000.00") & "-"
            Else
                cad = cad & Format(Abs(DBLet(RS!BaseReten, "N")), "00000.00") & "+"
            End If
        End If
        
        ' porcentaje de retencion
        cad = cad & Format(Round2(DBLet(RS!porc_ret, "N") * 100, 0), "0000")
        If DBLet(RS!ImpReten, "N") >= 0 Then
            cad = cad & Format(DBLet(RS!ImpReten, "N"), "000.00") & "+"
        Else
            cad = cad & Format(Abs(DBLet(RS!ImpReten, "N")), "000.00") & "-"
        End If
        
        ' gastos de la cooperativa
        CoopSoc = DevuelveValor("select codcoope from rsocios where codsocio = " & DBLet(RS!Codsocio, "N"))
        
        vPorcGasto = ""
        vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", CStr(CoopSoc), "N")
        If vPorcGasto = "" Then vPorcGasto = "0"
        
        sql = "select sum(imporvar) from rfacsoc_variedad where codtipom = " & DBSet(RS!CodTipom, "T")
        sql = sql & " and numfactu = " & DBSet(RS!numfactu, "N")
        sql = sql & " and fecfactu = " & DBSet(RS!fecfactu, "F")
        
        vImporte = DevuelveValor(sql)
        Gastos = Round2(vImporte * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
        
        cad = cad & Format(Round2(CCur(vPorcGasto) * 100, 0), "0000")
        If Gastos >= 0 Then
            cad = cad & Format(Abs(Gastos), "000.00") & "+"
        Else
            cad = cad & Format(Abs(Gastos), "000.00") & "-"
        End If
        
    End If
    
    Print #NFich, cad
    
    LineaTraspasoCoop = True
    Exit Function
    
eLineaTraspasoCoop:
    MuestraError Err.Number, "Carga Linea de Traspaso Cooperativas", Err.Description
End Function



Private Function CopiarFicheroCoop(Coop As String) As Boolean
Dim nomFich As String
On Error GoTo ecopiarfichero

    CopiarFicheroCoop = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.cd1.DefaultExt = "txt"
    
    cd1.Filter = "Archivos txt|txt|"
    cd1.FilterIndex = 1
    
    ' copiamos el primer fichero
    Select Case CInt(Coop)
        Case 1, 3, 4
            cd1.FileName = "tex.irp"
        Case 5, 6
            cd1.FileName = "liquid"
        Case Else
             cd1.FileName = "liquid"
    End Select
    
    Me.cd1.ShowSave
    
    If cd1.FileName <> "" Then
        FileCopy App.Path & "\trascoop.txt", cd1.FileName
    End If
    
    CopiarFicheroCoop = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function


' carga tabla tmpclasifica para el listado de kilos por socio cooperativa
Private Function CargarTemporal6(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim sql As String
Dim Sql2 As String
Dim Kilos1 As Long
Dim Kilos1VC As Long
Dim Kilos2 As Long
Dim Kilos2VC As Long
Dim Kilos3 As Long
Dim Kilos4 As Long
Dim Kilos5 As Long
Dim Kilos5VC As Long
Dim Kilos6 As Long
Dim Kilos7 As Long
Dim Kilos8 As Long
Dim Kilos9 As Long
Dim vCond As String
Dim vCond2 As String
Dim vResult As String
Dim NumRegElim As Integer
Dim cadena As String


    On Error GoTo eCargarTemporal
    
    Screen.MousePointer = vbHourglass
    
    CargarTemporal6 = False

    Sql2 = "delete from tmpclasifica where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    sql = "Select variedades.codvarie FROM variedades "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " where " & cWhere
    End If

    vCond = ""
    vCond2 = ""
    
    If txtCodigo(54).Text <> "" Then vCond = vCond & " and rhisfruta.codsocio >= " & DBSet(txtCodigo(54).Text, "N")
    If txtCodigo(55).Text <> "" Then vCond = vCond & " and rhisfruta.codsocio <= " & DBSet(txtCodigo(55).Text, "N")
    
    If txtCodigo(52).Text <> "" Then vCond = vCond & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(52).Text, "F")
    If txtCodigo(53).Text <> "" Then vCond = vCond & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(53).Text, "F")
    
    If Check7.Value = 1 Then
        If txtCodigo(54).Text <> "" Then vCond2 = vCond2 & " and rclasifica.codsocio >= " & DBSet(txtCodigo(54).Text, "N")
        If txtCodigo(55).Text <> "" Then vCond2 = vCond2 & " and rclasifica.codsocio <= " & DBSet(txtCodigo(55).Text, "N")
        
        If txtCodigo(52).Text <> "" Then vCond2 = vCond2 & " and rclasifica.fechaent >= " & DBSet(txtCodigo(52).Text, "F")
        If txtCodigo(53).Text <> "" Then vCond2 = vCond2 & " and rclasifica.fechaent <= " & DBSet(txtCodigo(53).Text, "F")
    End If
    
    cadena = ""
    For NumRegElim = 1 To ListView1(0).ListItems.Count
        If ListView1(0).ListItems(NumRegElim).Checked Then
            cadena = cadena & (NumRegElim - 1) & ","
        End If
    Next NumRegElim
    ' quitamos la ultima coma
    If cadena <> "" Then
        cadena = Mid(cadena, 1, Len(cadena) - 1)
    End If
    
    If cadena <> "" Then
        vCond = vCond & " and rhisfruta.tipoentr in (" & cadena & ")"
        vCond2 = vCond2 & " and rclasifica.tipoentr in (" & cadena & ")"
    Else
        vCond = vCond & " and rhisfruta.tipoentr = -1"
        vCond2 = vCond2 & " and rclasifica.tipoentr = -1"
    End If
    
    
    vResult = ""
    
    
    ' obtenemos los kilos de cada variedad con las condiciones
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        
        'KILOS PRODUCCION NORMAL COOPERATIVA --> KILOS1
        Sql2 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif, rhisfruta "
        Sql2 = Sql2 & " where rhisfruta.codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " and rhisfruta.tipoentr <> 2 " ' produccion normal
        Sql2 = Sql2 & " and rhisfruta.recolect = 0 " ' recolectado cooperativa
        Sql2 = Sql2 & " and rhisfruta.numalbar = rhisfruta_clasif.numalbar "
        If vCond <> "" Then
            Sql2 = Sql2 & vCond
        End If
        
        Kilos1 = DevuelveValor(Sql2)
        
        'KILOS PRODUCCION VENTACAMPO COOPERATIVA --> KILOS1VC
        Sql2 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif, rhisfruta "
        Sql2 = Sql2 & " where rhisfruta.codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " and rhisfruta.tipoentr = 1 " ' produccion VENTACAMPO
        Sql2 = Sql2 & " and rhisfruta.recolect = 0 " ' recolectado cooperativa
        Sql2 = Sql2 & " and rhisfruta.numalbar = rhisfruta_clasif.numalbar "
        If vCond <> "" Then
            Sql2 = Sql2 & vCond
        End If
        
        Kilos1VC = DevuelveValor(Sql2)
        
        
        
        If Check7.Value Then
            Sql2 = "select sum(rclasifica.kilosnet) from rclasifica "
            Sql2 = Sql2 & " where rclasifica.codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and rclasifica.tipoentr <> 2 " ' produccion normal
            Sql2 = Sql2 & " and rclasifica.recolect = 0 "  ' recolectado cooperativa
            If vCond2 <> "" Then
                Sql2 = Sql2 & vCond2
            End If
        
            Kilos1 = Kilos1 + DevuelveValor(Sql2)
        
            'VENTACAMPO
            Sql2 = "select sum(rclasifica.kilosnet) from rclasifica "
            Sql2 = Sql2 & " where rclasifica.codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and rclasifica.tipoentr = 1 " ' venta campo
            Sql2 = Sql2 & " and rclasifica.recolect = 0 "  ' recolectado cooperativa
            If vCond2 <> "" Then
                Sql2 = Sql2 & vCond2
            End If
        
            Kilos1VC = Kilos1VC + DevuelveValor(Sql2)
        
        
        End If
        
        
        'KILOS PRODUCCION NORMAL SOCIO --> KILOS2
        Sql2 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif, rhisfruta "
        Sql2 = Sql2 & " where rhisfruta.codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " and rhisfruta.tipoentr <> 2 " ' produccion normal
        Sql2 = Sql2 & " and rhisfruta.recolect = 1 " ' recolectado socio
        Sql2 = Sql2 & " and rhisfruta.numalbar = rhisfruta_clasif.numalbar "
        If vCond <> "" Then
            Sql2 = Sql2 & vCond
        End If
        
        Kilos2 = DevuelveValor(Sql2)
    
        'KILOS PRODUCCION VENTA CAMPO SOCIO --> KILOS2VC
        Sql2 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif, rhisfruta "
        Sql2 = Sql2 & " where rhisfruta.codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " and rhisfruta.tipoentr = 1 " ' venta campo
        Sql2 = Sql2 & " and rhisfruta.recolect = 1 " ' recolectado socio
        Sql2 = Sql2 & " and rhisfruta.numalbar = rhisfruta_clasif.numalbar "
        If vCond <> "" Then
            Sql2 = Sql2 & vCond
        End If
        
        Kilos2VC = DevuelveValor(Sql2)
    
    
    
        If Check7.Value Then
            Sql2 = "select sum(rclasifica.kilosnet) from rclasifica "
            Sql2 = Sql2 & " where rclasifica.codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and rclasifica.tipoentr <> 2 " ' produccion normal
            Sql2 = Sql2 & " and rclasifica.recolect = 1 "  ' recolectado socio
            If vCond2 <> "" Then
                Sql2 = Sql2 & vCond2
            End If
        
            Kilos2 = Kilos2 + DevuelveValor(Sql2)
        
            ' VENTA CAMPO
            Sql2 = "select sum(rclasifica.kilosnet) from rclasifica "
            Sql2 = Sql2 & " where rclasifica.codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and rclasifica.tipoentr = 1 " ' venta campo
            Sql2 = Sql2 & " and rclasifica.recolect = 1 "  ' recolectado socio
            If vCond2 <> "" Then
                Sql2 = Sql2 & vCond2
            End If
        
            Kilos2VC = Kilos2VC + DevuelveValor(Sql2)
        
        
        End If
    
    
        ' KILOS PRODUCCION INTEGRADA COOPERATIVA --> KILOS3
        Sql2 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif, rhisfruta "
        Sql2 = Sql2 & " where rhisfruta.codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " and rhisfruta.tipoentr = 2 " ' produccion integrada
        Sql2 = Sql2 & " and rhisfruta.recolect = 0 " ' recolectado cooperativa
        Sql2 = Sql2 & " and rhisfruta.numalbar = rhisfruta_clasif.numalbar "
        If vCond <> "" Then
            Sql2 = Sql2 & vCond
        End If
        
        Kilos3 = DevuelveValor(Sql2)
        
        If Check7.Value Then
            Sql2 = "select sum(rclasifica.kilosnet) from rclasifica "
            Sql2 = Sql2 & " where rclasifica.codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and rclasifica.tipoentr = 2 " ' produccion integrada
            Sql2 = Sql2 & " and rclasifica.recolect = 0 "  ' recolectado cooperativa
            If vCond2 <> "" Then
                Sql2 = Sql2 & vCond2
            End If
        
            Kilos3 = Kilos3 + DevuelveValor(Sql2)
        End If
        
        ' KILOS PRODUCCION INTEGRADA SOCIO --> KILOS4
        Sql2 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif, rhisfruta "
        Sql2 = Sql2 & " where rhisfruta.codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " and rhisfruta.tipoentr = 2 " ' produccion integrada
        Sql2 = Sql2 & " and rhisfruta.recolect = 1 " ' recolectado socio
        Sql2 = Sql2 & " and rhisfruta.numalbar = rhisfruta_clasif.numalbar "
        If vCond <> "" Then
            Sql2 = Sql2 & vCond
        End If
        
        Kilos4 = DevuelveValor(Sql2)
        
        If Check7.Value Then
            Sql2 = "select sum(rclasifica.kilosnet) from rclasifica "
            Sql2 = Sql2 & " where rclasifica.codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and rclasifica.tipoentr = 2 " ' produccion integrada
            Sql2 = Sql2 & " and rclasifica.recolect = 1 " ' recolectado socio
            If vCond2 <> "" Then
                Sql2 = Sql2 & vCond2
            End If
        
            Kilos4 = Kilos4 + DevuelveValor(Sql2)
        End If
        
        'TOTAL PRODUCCION NORMAL POR VARIEDAD --> KILOS5
        Kilos5 = Kilos2 + Kilos1
        
        'TOTAL PRODUCCION INTEGRADA POR VARIEDAD --> KILOS6
        Kilos6 = Kilos4 + Kilos3
        
        'TOTAL PRODUCCION POR SOCIO --> KILOS7
        Kilos7 = Kilos2 + Kilos4
        
        'TOTAL PRODUCCION COOPERATIVA --> KILOS8
        Kilos8 = Kilos1 + Kilos3
        
        'TOTAL KILOS VARIEDAD --> KILOS9
        Kilos9 = Kilos7 + Kilos8
    
    
        'TOTAL PRODUCCION VENTA CAMPO POR VARIEDAD --> KILOS5
        Kilos5VC = Kilos2VC + Kilos1VC
    
        vResult = vResult & "(" & vUsu.Codigo & "," & DBSet(RS!codvarie, "N") & ","
        vResult = vResult & DBSet(Kilos2, "N", "S") & "," & DBSet(Kilos1, "N", "S") & ","
        vResult = vResult & DBSet(Kilos5, "N", "S") & "," & DBSet(Kilos4, "N", "S") & ","
        vResult = vResult & DBSet(Kilos3, "N", "S") & "," & DBSet(Kilos6, "N", "S") & ","
        vResult = vResult & DBSet(Kilos7, "N", "S") & "," & DBSet(Kilos8, "N", "S") & ","
        vResult = vResult & DBSet(Kilos9, "N", "S") & ","
        vResult = vResult & DBSet(Kilos2VC, "N", "S") & "," & DBSet(Kilos1VC, "N", "S") & ","
        vResult = vResult & DBSet(Kilos5VC, "N", "S") & "),"
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If vResult <> "" Then
        Sql2 = "insert into tmpclasifica (codusu,codvarie,cal1,cal2,cal3,cal4,cal5,"
        Sql2 = Sql2 & "cal6,cal7,cal8,cal9,cal10,cal11,cal12) values "
        
        Sql2 = Sql2 & Mid(vResult, 1, Len(vResult) - 1)  ' quitamos la ultima coma
    End If

    conn.Execute Sql2
    
    ' borramos aquellos registros que no tienen kilos de ningun tipo
    Sql2 = "delete from tmpclasifica where cal1 is null and cal2 is null and cal3 is null and "
    Sql2 = Sql2 & " cal4 is null and cal5 is null and cal6 is null and cal7 is null and "
    Sql2 = Sql2 & " cal8 is null and cal9 is null and cal10 is null and cal11 is null and cal12 is null and codusu = " & DBSet(vUsu.Codigo, "N")
    
    conn.Execute Sql2
    
    CargarTemporal6 = True
    Screen.MousePointer = vbDefault

    Exit Function
    
eCargarTemporal:
    Screen.MousePointer = vbDefault
    MuestraError "Cargando temporal", Err.Description
End Function


'Lineas traspaso ropas


Private Function LineaTraspasoSocioROPAS(NFich As Integer, ByRef RS As ADODB.Recordset) As Boolean
Dim cad As String
Dim Areas As Long
Dim Tipo As Integer
Dim sql As String
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim CodiIVA As String
Dim TipoIRPF As Byte
Dim PorcIva As String
Dim vPorcIva As Currency
Dim CoopSoc As Currency

Dim Producto As String
Dim Variedad As String
Dim NomVar As String
Dim codVar As Long

Dim nifSocio As String
Dim Kilos As Long
Dim vPorcGasto As String
Dim vImporte As Currency
Dim Gastos As Currency
Dim vCaracter As String


    On Error GoTo eLineaTraspasoSocioROPAS

    LineaTraspasoSocioROPAS = False

    cad = ""
    cad = cad & Format(txtCodigo(62).Text, "0000") & ";"
    cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 12) & ";"
    cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 12) & ";"
    cad = cad & RellenaABlancos(RS!nifSocio, True, 12) & ";"
    cad = cad & Format(RS!Codsocio, "######") & ";"
    cad = cad & RellenaABlancos(RS!nomsocio, True, 60) & ";ES;"
    
    '[Monica]08/04/2014: para el caso de picassent depende de que el socio tenga CIF
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        'si nos encontramos una letra al principio, entonces se trata de un cif
        vCaracter = Asc(Mid(Trim(DBLet(RS!nifSocio, "T")), 1, 1))
        If (vCaracter >= 65 And vCaracter <= 90) Or (vCaracter >= 97 And vCaracter <= 122) Then
            cad = cad & "X;"
        Else
            cad = cad & "P;"
        End If
    Else
        'como estaba
        If DBLet(RS!TipoIRPF, "N") <> 2 Then
            cad = cad & "P;"
        Else
            cad = cad & "J;"
        End If
    End If
    
    cad = cad & Format(DBLet(RS!FecAlta, "F"), "dd/mm/yyyy") & ";"
    
    If Not IsNull(RS!fecbaja) And DBLet(RS!fecbaja) <> "" Then
        cad = cad & ";" & Format(DBLet(RS!fecbaja, "F"), "dd/mm/yyyy")
    End If

    Print #NFich, cad
    
    LineaTraspasoSocioROPAS = True
    Exit Function
    
eLineaTraspasoSocioROPAS:
    MuestraError Err.Number, "Carga Linea de Traspaso Socios ROPAS", Err.Description
End Function



Private Function LineaTraspasoCampoROPAS(NFich As Integer, ByRef RS As ADODB.Recordset, Lin As Integer) As Boolean
Dim cad As String
Dim Areas As Long
Dim Tipo As Integer
Dim sql As String
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim CodiIVA As String
Dim TipoIRPF As Byte
Dim PorcIva As String
Dim vPorcIva As Currency
Dim CoopSoc As Currency

Dim Producto As String
Dim Variedad As String
Dim NomVar As String
Dim codVar As Long

Dim nifSocio As String
Dim Kilos As Long
Dim vPorcGasto As String
Dim vImporte As Currency

Dim Pobla As String
Dim CodSigPa As String
Dim HectaSig As Currency
Dim FecAlta As Date
Dim CodConse As Long
Dim CanAfo As Currency
Dim Super As Currency
Dim I As Integer
Dim SubParce As String

Dim HectaSigParcela As Currency
Dim HectaSigRecinto As Currency
Dim HectaSigSubRecinto As Currency
Dim SuperLinea As Currency


Dim Rs2 As ADODB.Recordset


    On Error GoTo eLineaTraspasoCampoROPAS

    LineaTraspasoCampoROPAS = False


    sql = "select * from rcampos_cooprop where codcampo = " & DBSet(RS!codcampo, "N")
    
    
    If TotalRegistrosConsulta(sql) <> 0 Then
    
        Set Rs2 = New ADODB.Recordset
        Rs2.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        Pobla = ""
        Pobla = DevuelveValor("select codpobla from rpartida where codparti = " & DBSet(RS!codparti, "N"))
        
        HectaSig = 0 '  SUPERFICIE TOTAL PARCELA
        
        sql = "select sum(supsigpa) from rcampos, rpartida where poligono = " & DBSet(RS!Poligono, "N")
        sql = sql & " and parcela = " & DBSet(RS!Parcela, "N")
        sql = sql & " and rcampos.fecbajas is null "
        sql = sql & " and rpartida.codpobla = " & DBSet(Pobla, "T")
        sql = sql & " and rcampos.codparti = rpartida.codparti "
        
        HectaSig = DevuelveValor(sql)
        
        HectaSigRecinto = 0 '  SUPERFICIE TOTAL RECINTO
        
        sql = "select sum(supsigpa) from rcampos, rpartida where poligono = " & DBSet(RS!Poligono, "N")
        sql = sql & " and parcela = " & DBSet(RS!Parcela, "N")
        sql = sql & " and recintos = " & DBSet(RS!recintos, "N")
        sql = sql & " and rcampos.fecbajas is null "
        sql = sql & " and rpartida.codpobla = " & DBSet(Pobla, "T")
        sql = sql & " and rcampos.codparti = rpartida.codparti "
        
        HectaSigRecinto = DBLet(RS!supsigpa, "N") 'DevuelveValor(Sql)
        
        Super = DBLet(RS!supcoope, "N")
        If DBLet(RS!supcoope, "N") > DBLet(RS!supsigpa, "N") Then
            Super = DBLet(RS!supsigpa, "N")
        End If
        
        I = 1
        
        While Not Rs2.EOF
            Set vSocio = New cSocio
        
            If vSocio.LeerDatos(Rs2!Codsocio) Then
                
                cad = ""
                cad = cad & Format(txtCodigo(62).Text, "0000") & ";"
                cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 12) & ";"
                cad = cad & RellenaABlancos(vSocio.nif, True, 12) & ";ES;"
                
                If RS!tipoparc = 0 Then
                    cad = cad & "R;"
                    cad = cad & Space(27) & ";"
                Else
                    cad = cad & "U;"
                    cad = cad & RellenaABlancos(DBLet(RS!refercatas, "T"), True, 27) & ";"
                End If
                
            
                Pobla = ""
                Pobla = DevuelveValor("select codpobla from rpartida where codparti = " & DBSet(RS!codparti, "N"))
                
                cad = cad & Mid(Pobla, 1, 2) & ";"
            
                CodSigPa = ""
                CodSigPa = DevuelveValor("select codsigpa from rpueblos where codpobla = " & DBSet(Pobla, "T"))
        
                cad = cad & Format(CodSigPa, "###") & ";"
                
                If RS!tipoparc = 0 Then
                    cad = cad & "000;"
                    cad = cad & "00;"
                    
                    
                    cad = cad & Format(DBLet(RS!Poligono, "N"), "###") & ";"
                    cad = cad & Format(DBLet(RS!Parcela, "N"), "#####") & ";"
                    cad = cad & Format(DBLet(RS!recintos, "N"), "#####") & ";"
                
                    SubParce = Trim(DBLet(RS!SubParce)) & I
                    
                    cad = cad & RellenaABlancos(SubParce, True, 2) & ";"
                Else
                    cad = cad & ";;;;;;"
                
                End If
                
                
        
                cad = cad & Format(HectaSig, "##0.0000") & ";"
        
                cad = cad & Format(HectaSigRecinto, "##0.0000") & ";" ' antes estaba rs!supsigpa
        
                ' este seria el prorrateo
                HectaSigSubRecinto = Round2(HectaSigRecinto * DBLet(Rs2!Porcentaje, "N") / 100, 4)
                cad = cad & Format(HectaSigSubRecinto, "##0.0000") & ";"
            
                FecAlta = DBLet(RS!FecAltas, "F")
                
                '[Monica]23/01/2013: si la fecha de alta del campo es anterior a la fecha de alta de socio
                '                    que ponga la fecha de alta del socio
                If FecAlta < vSocio.FechaAlta Then FecAlta = vSocio.FechaAlta
                
        
                cad = cad & Format(FecAlta, "dd/mm/yyyy") & ";"
                If DBLet(RS!fecbajas) <> "" Then
                    cad = cad & Format(RS!fecbajas, "dd/mm/yyyy") & ";"
                Else
                    cad = cad & ";"
                End If
                
                cad = cad & Format(I, "#") & ";"  ' contador de subparcelas
                
                CodConse = 0
                CodConse = DevuelveValor("select codconse from variedades where codvarie = " & DBSet(RS!codvarie, "N"))
                
                cad = cad & RellenaABlancos(CStr(CodConse), True, 6) & ";"
                
                SuperLinea = Round2(Super * DBLet(Rs2!Porcentaje, "N") / 100, 4)
                
                cad = cad & Format(SuperLinea, "##0.0000") & ";"
                
                '[Monica]26/04/2012: a�ado esta instruccion
                CanAfo = Round2(DBLet(RS!Canaforo, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 0)
                
                If CanAfo = 0 Then Let CanAfo = 10
                '[Monica]26/04/2012: sustituyo esta instruccion por la de abajo
            '    CanAfo = Round2(Rs!canaforo / 1000, 2) 'En toneladas
                CanAfo = Round2(CanAfo / 1000, 2) 'En toneladas
                
                cad = cad & Format(CanAfo, "###0.00")
                
                Print #NFich, cad
            
                I = I + 1
            
            
            
            End If
            
            Set vSocio = Nothing
            
            Rs2.MoveNext
        Wend
        
    Else

        cad = ""
        cad = cad & Format(txtCodigo(62).Text, "0000") & ";"
        cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 12) & ";"
        cad = cad & RellenaABlancos(RS!nifSocio, True, 12) & ";ES;"
        
        If RS!tipoparc = 0 Then
            cad = cad & "R;"
            cad = cad & Space(27) & ";"
        Else
            cad = cad & "U;"
            cad = cad & RellenaABlancos(DBLet(RS!refercatas, "T"), True, 27) & ";"
        End If
        
        
        Pobla = ""
        Pobla = DevuelveValor("select codpobla from rpartida where codparti = " & DBSet(RS!codparti, "N"))
        
        cad = cad & Mid(Pobla, 1, 2) & ";"
        
        CodSigPa = ""
        CodSigPa = DevuelveValor("select codsigpa from rpueblos where codpobla = " & DBSet(Pobla, "T"))
        
        cad = cad & Format(CodSigPa, "###") & ";"
        
        If DBLet(RS!tipoparc, "N") = 0 Then
            cad = cad & "000;"
            cad = cad & "00;"
            cad = cad & Format(DBLet(RS!Poligono, "N"), "###") & ";"
            cad = cad & Format(DBLet(RS!Parcela, "N"), "#####") & ";"
            cad = cad & Format(DBLet(RS!recintos, "N"), "#####") & ";"
            
            cad = cad & RellenaABlancos(DBLet(RS!SubParce, "T"), True, 2) & ";"
        Else
            cad = cad & ";;;;;;"
        End If
            
        
        HectaSig = 0 '  SUPERFICIE TOTAL PARCELA
        
        sql = "select sum(supsigpa) from rcampos, rpartida where poligono = " & DBSet(RS!Poligono, "N")
        sql = sql & " and parcela = " & DBSet(RS!Parcela, "N")
        sql = sql & " and rcampos.fecbajas is null "
        sql = sql & " and rpartida.codpobla = " & DBSet(Pobla, "T")
        sql = sql & " and rcampos.codparti = rpartida.codparti "
        
        HectaSig = DevuelveValor(sql)
        
        cad = cad & Format(HectaSig, "##0.0000") & ";"
        
        HectaSigRecinto = 0 '  SUPERFICIE TOTAL RECINTO
        
        sql = "select sum(supsigpa) from rcampos, rpartida where poligono = " & DBSet(RS!Poligono, "N")
        sql = sql & " and parcela = " & DBSet(RS!Parcela, "N")
        sql = sql & " and recintos = " & DBSet(RS!recintos, "N")
        sql = sql & " and rcampos.fecbajas is null "
        sql = sql & " and rpartida.codpobla = " & DBSet(Pobla, "T")
        sql = sql & " and rcampos.codparti = rpartida.codparti "
        
        HectaSigRecinto = DBLet(RS!supsigpa, "N") 'DevuelveValor(Sql)
        
        cad = cad & Format(HectaSigRecinto, "##0.0000") & ";" ' antes estaba rs!supsigpa
        cad = cad & Format(HectaSigRecinto, "##0.0000") & ";" ' antes estaba RS!supcatas
        
        FecAlta = DBLet(RS!FecAltas, "F")
        '[Monica]23/01/2013: si la fecha de alta del campo es anterior a la fecha de alta de socio
        '                    que ponga la fecha de alta del socio
        If RS!FecAlta > RS!FecAltas Then ' fecha alta socio > fecha alta campo
            FecAlta = RS!FecAlta
        End If
        
        cad = cad & Format(FecAlta, "dd/mm/yyyy") & ";"
        If DBLet(RS!fecbajas) <> "" Then
            cad = cad & Format(RS!fecbajas, "dd/mm/yyyy") & ";"
        Else
            cad = cad & ";"
        End If
        Lin = 1
        cad = cad & Format(Lin, "#") & ";"  ' contador de subparcelas
        
        
        CodConse = 0
        CodConse = DevuelveValor("select codconse from variedades where codvarie = " & DBSet(RS!codvarie, "N"))
        
        cad = cad & RellenaABlancos(CStr(CodConse), True, 6) & ";"
        
        Super = DBLet(RS!supcoope, "N")
        If DBLet(RS!supcoope, "N") > DBLet(RS!supsigpa, "N") Then
            Super = DBLet(RS!supsigpa, "N")
        End If
        
        cad = cad & Format(Super, "##0.0000") & ";"
        
        '[Monica]26/04/2012: a�ado esta instruccion
        CanAfo = DBLet(RS!Canaforo, "N")
        
        If CanAfo = 0 Then Let CanAfo = 10
        '[Monica]26/04/2012: sustituyo esta instruccion por la de abajo
    '    CanAfo = Round2(Rs!canaforo / 1000, 2) 'En toneladas
        CanAfo = Round2(CanAfo / 1000, 2) 'En toneladas
        
        cad = cad & Format(CanAfo, "###0.00")
        
        Print #NFich, cad
    
    End If
    
    LineaTraspasoCampoROPAS = True
    Exit Function
    
eLineaTraspasoCampoROPAS:
    MuestraError Err.Number, "Carga Linea de Traspaso Campos ROPAS", Err.Description
End Function


Private Function CopiarFicheroROPAS() As Boolean
Dim nomFich As String
Dim FechaEnvio As String


On Error GoTo ecopiarfichero

    CopiarFicheroROPAS = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.cd1.DefaultExt = "csv"
    
    FechaEnvio = Mid(txtCodigo(132).Text, 7, 4) & Mid(txtCodigo(132).Text, 4, 2) & Mid(txtCodigo(132).Text, 1, 2)
    
    cd1.Filter = "Archivos csv|csv|"
    cd1.FilterIndex = 1
    'cd1.FileName = "socios.csv"
    cd1.FileName = "Socios_" & Format(txtCodigo(62).Text, "0000") & "_" & FechaEnvio & "_" & vParam.CifEmpresa & ".csv"
    Me.cd1.ShowSave
    
    If cd1.FileName <> "" Then
        'FileCopy App.Path & "\socios.csv", cd1.FileName
        FileCopy App.Path & "\Socios_" & Format(txtCodigo(62).Text, "0000") & "_" & FechaEnvio & "_" & vParam.CifEmpresa & ".csv", cd1.FileName
        
        'cd1.FileName = "parcelas.csv"
        'FileCopy App.Path & "\parcelas.csv", cd1.FileName
        cd1.FileName = "Parcela_" & Format(txtCodigo(62).Text, "0000") & "_" & FechaEnvio & "_" & vParam.CifEmpresa & ".csv"
        FileCopy App.Path & "\Parcela_" & Format(txtCodigo(62).Text, "0000") & "_" & FechaEnvio & "_" & vParam.CifEmpresa & ".csv", cd1.FileName

    End If
    
    CopiarFicheroROPAS = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim sql As String
Dim Sql2 As String
Dim vClien As cSocio
' a�adido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String
Dim Nregs As Long

    B = True
    Select Case OpcionListado
        Case 19 ' fichero de agriweb
            If B Then
                If txtCodigo(27).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente el a�o del ejercicios.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(26)
                End If
            End If
            If B Then
                If txtCodigo(28).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente el CIF de la industria transformadora.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(28)
                End If
            End If
            If B Then
                If txtCodigo(29).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente los kilos contratados.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(29)
                End If
            End If
            If B Then
                If txtCodigo(30).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de formalizaci�n.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(30)
                End If
            End If
            If B Then
                If txtCodigo(31).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Superficie Total de Contrato.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(31)
                End If
            End If
            If B Then
                If txtCodigo(32).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente el Precio Estipulado de Compra.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(32)
                End If
            End If
            
        Case 21 ' traspaso desde el calibrador
            ' en el caso del calibrador grande de Catadau hemos de introducir
            ' obligatoriamente la fecha
            If (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 9) And Combo1(6).ListIndex = 0 Then
                If txtCodigo(63).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la fecha de calibrado.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(63)
                End If
            End If
             '[Monica]21/04/2016: a�adidas las notas para Castellduc
            If vParamAplic.Cooperativa = 5 Then
                If txtCodigo(170).Text = "" Or txtCodigo(179).Text = "" Then
                    MsgBox "Debe introducir desde hasta notas.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(170)
                End If
            End If
            
        Case 28 ' alta masiva de bonificaciones
            If txtCodigo(75).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la variedad.", vbExclamation
                B = False
                PonerFoco txtCodigo(75)
            End If
            
            If B And txtCodigo(74).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la fecha de inicio.", vbExclamation
                B = False
                PonerFoco txtCodigo(74)
            End If
            
            If B And ExistenBonificaciones Then
                MsgBox "Existen bonificaciones para esa variedad en el rango de fechas. Revise.", vbExclamation
                B = False
                PonerFoco txtCodigo(75)
            End If
            
        Case 29 ' baja masiva de bonificaciones
            If txtCodigo(75).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la variedad.", vbExclamation
                B = False
                PonerFoco txtCodigo(75)
            End If
        
        Case 30 ' Generacion de clasificacion (Picassent)
            If txtCodigo(83).Text = "" Then
                MsgBox "Debe introducir obligatoriamente un socio.", vbExclamation
                B = False
                PonerFoco txtCodigo(83)
            Else
                If EstaSocioDeAlta(txtCodigo(83)) Then
                    Dim vSocio As cSocio
                    Set vSocio = New cSocio
                    If Not vSocio.LeerDatosSeccion(txtCodigo(83).Text, vParamAplic.Seccionhorto) Then
                         MsgBox "El socio no est� dado de alta en la secci�n Hortofrut�cola. Revise.", vbExclamation
                         B = False
                         PonerFoco txtCodigo(83)
                    End If
                End If
            End If
            If B Then
                If txtCodigo(80).Text <> "" Then
                    sql = "select count(*) from rcampos where codsocio = " & DBSet(txtCodigo(83).Text, "N")
                    sql = sql & " and nrocampo = " & DBSet(txtCodigo(80).Text, "N")
                    sql = sql & " and codvarie = " & DBSet(RecuperaValor(CadTag, 1), "N")
                    Nregs = TotalRegistros(sql)
                    If Nregs = 0 Then
                        MsgBox "No existe el campo de ese socio variedad. Revise.", vbExclamation
                        B = False
                        PonerFoco txtCodigo(80)
                    Else
                        If Nregs > 1 Then
                            MsgBox "Hay m�s de un campo. Revise.", vbExclamation
                            B = False
                            PonerFoco txtCodigo(80)
                        End If
                    End If
                End If
            End If
            If B Then
                If txtCodigo(79).Text = "" Then
                    MsgBox "Debe de introducir un valor en Porcentaje de Destrio.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(79)
                End If
            End If
    
        Case 40 ' Impresion de Ordenes de Recoleccion
            If txtNombre(147).Text = "" And Not EsReimpresion Then
                MsgBox "Debe introducir obligatoriamente un responsable.", vbExclamation
                B = False
                PonerFoco txtCodigo(147)
            End If
            If B Then
                If txtNombre(149).Text = "" And Not EsReimpresion Then
                    MsgBox "Debe introducir obligatoriamente una variedad.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(149)
                End If
            End If
'[Monica]30/09/2013: dejo que no metan la partida
'            If b Then
'                If txtNombre(142).Text = "" And Not EsReimpresion Then
'                    MsgBox "Debe de introducir una Partida.", vbExclamation
'                    b = False
'                    PonerFoco txtcodigo(142)
'                End If
'            End If
            If B Then
                If txtCodigo(138).Text = "" And Not EsReimpresion Then
                    MsgBox "Debe de introducir una Fecha.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(138)
                End If
            End If
            
            If B And EsReimpresion Then
                If txtCodigo(141).Text = "" Then
                    MsgBox "Si es reimpresi�n, debe de introducir el nro de orden.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(141)
                End If
            End If
            
        Case 48 ' traspaso de albaranes de retirada
            If txtCodigo(169).Text = "" Then
                MsgBox "Debe seleccionar una cooperativa. Reintroduzca.", vbExclamation
                B = False
                PonerFoco txtCodigo(169)
            End If
    End Select
    DatosOK = B

End Function

'********************************************************************
'***************** TRASPASO DE DATOS DE ALMAZARA ********************
'********************************************************************

'********************* solo para Mogente ****************************


Private Sub CmdAcepTrasDatosAlmz_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim B As Boolean
Dim vSQL As String

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

     '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtCodigo(64).Text)
    cHasta = Trim(txtCodigo(65).Text)
    nDesde = txtNombre(64).Text
    nHasta = txtNombre(65).Text
    If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
     
    tabla = "rcampos INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio and rcampos.fecbajas is null  "
    tabla = "(" & tabla & ") INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie "
    tabla = "(" & tabla & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
    tabla = "(" & tabla & ") INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
    tabla = tabla & " and grupopro.codgrupo = 5 " 'grupo de oliva
    
      
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadSelect) Then
        B = GeneraFicheroTraspasoAlmazara(tabla, cadSelect)
        If B Then
            If CopiarFicheroDatosAlmz() Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                cmdCancel_Click (4)
            End If
        End If
     End If


    
End Sub


Private Function GeneraFicheroTraspasoAlmazara(pTabla As String, pWhere As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim vSocio As cSocio
Dim B As Boolean
Dim Nregs As Long
Dim Total As Variant
Dim nomparti As String
Dim nomvarie As String

Dim cTabla As String
Dim vWhere As String


    On Error GoTo EGen
    GeneraFicheroTraspasoAlmazara = False
    
    cTabla = pTabla
    vWhere = pWhere
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = QuitarCaracterACadena(cTabla, "_1")
    If vWhere <> "" Then
        vWhere = QuitarCaracterACadena(vWhere, "{")
        vWhere = QuitarCaracterACadena(vWhere, "}")
        vWhere = QuitarCaracterACadena(vWhere, "_1")
    End If
    
    NFic = FreeFile
    
    Open App.Path & "\trasalmz.txt" For Output As #NFic
    
    Set RS = Nothing
    
    'Imprimimos las lineas
    Aux = "select  rcampos.*, rsocios.* "
    Aux = Aux & " from " & cTabla
    If vWhere <> "" Then Aux = Aux & " where " & vWhere
    Aux = Aux & " order by rcampos.codsocio, rcampos.codcampo "
    
    Set RS = New ADODB.Recordset
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then
        'No hayningun registro
    Else
        B = True
        Regs = 0
        While Not RS.EOF And B
            Regs = Regs + 1
            
            nomparti = ""
            nomparti = DevuelveDesdeBDNew(cAgro, "rpartida", "nomparti", "codparti", RS!codparti, "N")
            
            nomvarie = ""
            nomvarie = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", RS!codvarie, "N")
            
            cad = ""
            cad = cad & RellenaABlancos(Format(RS!Codsocio, "000000"), True, 13)
            cad = cad & RellenaABlancos(RS!nifSocio, True, 14)
            cad = cad & RellenaABlancos(RS!nomsocio, True, 51)
            cad = cad & RellenaABlancos(RS!prosocio, True, 15)
            cad = cad & RellenaABlancos(RS!dirsocio, True, 44)
            cad = cad & RellenaABlancos(RS!codpostal, True, 12)
            cad = cad & RellenaABlancos(RS!pobsocio, True, 25)
            cad = cad & RellenaABlancos(Format(RS!codcampo, "00000000"), True, 9)
            cad = cad & RellenaABlancos(Format(RS!codparti, "0000"), True, 5)
            cad = cad & RellenaABlancos(nomparti, True, 35)
            cad = cad & RellenaABlancos(Format(RS!codvarie, "000000"), True, 10)
            cad = cad & RellenaABlancos(nomvarie, True, 25)
            cad = cad & RellenaABlancos(Format(RS!Poligono, "000"), True, 5)
            cad = cad & RellenaABlancos(Format(RS!Parcela, "000000"), True, 10)
            
        
            Print #NFic, cad
            
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
            
    Close (NFic)
    
    If Regs > 0 And B Then GeneraFicheroTraspasoAlmazara = True
    Exit Function
    
EGen:
    Set RS = Nothing
    Close (NFic)
    MuestraError Err.Number, Err.Description
End Function

Private Function CopiarFicheroDatosAlmz() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFicheroDatosAlmz = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.cd1.DefaultExt = "txt"
    
    cd1.Filter = "Archivos txt|txt|"
    cd1.FilterIndex = 1
    
    cd1.FileName = "DatosAlmazara.txt"
    
    Me.cd1.ShowSave
    
    If cd1.FileName <> "" Then
        FileCopy App.Path & "\trasalmz.txt", cd1.FileName
    End If
    
    CopiarFicheroDatosAlmz = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function

Private Sub VisualizarFecha(Indice As Integer)
    '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
    If Indice = 6 And (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 9) Then
        If Combo1(6).ListIndex = 0 Then
            FrameFecha.Enabled = True
            FrameFecha.visible = True
            PonerFoco txtCodigo(63)
        Else
            FrameFecha.Enabled = False
            FrameFecha.visible = False
            cmdAcepTras.SetFocus
        End If
    End If
End Sub

Private Function ExistenBonificaciones() As Boolean
Dim sql As String
Dim Dias As Long
Dim UltimaFecha As Date

    ExistenBonificaciones = False

    Dias = CCur(txtCodigo(76).Text)
    
    UltimaFecha = DateAdd("d", Dias, CDate(txtCodigo(74).Text))
    
    sql = "select count(*) from rbonifentradas where codvarie = " & DBSet(txtCodigo(75).Text, "N")
    sql = sql & " and fechaent >= " & DBSet(txtCodigo(74).Text, "F") & " and fechaent < " & DBSet(UltimaFecha, "F")

    ExistenBonificaciones = (TotalRegistros(sql) <> 0)

End Function



Private Function InsertarBonificaciones() As Boolean
Dim sql As String
Dim Sql2 As String
Dim Porcentaje As Currency
Dim I As Long
Dim Fecha As Date

    On Error GoTo eInsertarBonificaciones
        
    InsertarBonificaciones = False
        
    sql = "insert into rbonifentradas (codvarie, fechaent, porcbonif) values "
    
    Sql2 = ""
    Fecha = CDate(txtCodigo(74).Text)
    Porcentaje = CCur(ImporteSinFormato(txtCodigo(77).Text))
    For I = 1 To CCur(txtCodigo(76).Text)
    
        Sql2 = Sql2 & "(" & DBSet(txtCodigo(75).Text, "N") & "," & DBSet(Fecha, "F") & ","
        Sql2 = Sql2 & DBSet(Porcentaje, "N") & "),"
        
        ' le sumamos el indice de correccion al porcentaje
        Porcentaje = Porcentaje + CCur(ImporteSinFormato(txtCodigo(78).Text))
        Fecha = DateAdd("d", 1, Fecha)
        
        
    Next I
    
    If Sql2 <> "" Then
        'quitamos la ultima coma
        Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
        
        sql = sql & Sql2
        conn.Execute sql
    End If

    InsertarBonificaciones = True
    Exit Function
    

eInsertarBonificaciones:
    MuestraError Err.Number, "Insertando Bonificaciones", Err.Description
End Function


Private Function EliminarBonificaciones() As Boolean
Dim sql As String
Dim Sql2 As String
Dim Porcentaje As Currency

    On Error GoTo eEliminarBonificaciones
        
    EliminarBonificaciones = False
        
    sql = "delete from rbonifentradas where codvarie = " & DBSet(txtCodigo(75).Text, "N")
    
    conn.Execute sql

    EliminarBonificaciones = True
    Exit Function
    

eEliminarBonificaciones:
    MuestraError Err.Number, "Eliminando Bonificaciones", Err.Description
End Function


Private Function CargarTemporalDatosDestrio(vtabla As String, vWhere As String) As Boolean
Dim sql As String
Dim KilosTot As Currency
Dim KilosMan As Currency
Dim KilosPixat As Currency
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

    On Error GoTo eCargarTemporalDatosDestrio
    
    CargarTemporalDatosDestrio = True


    sql = "delete from tmpexcel where codusu = " & vUsu.Codigo
    conn.Execute sql
    
    sql = "select rcontrol.codvarie, codsocio, codcampo, fechacla from " & QuitarCaracterACadena(vtabla, "_1")
    If vWhere <> "" Then
        vWhere = QuitarCaracterACadena(vWhere, "{")
        vWhere = QuitarCaracterACadena(vWhere, "}")
        vWhere = QuitarCaracterACadena(vWhere, "_1")
        sql = sql & " WHERE " & vWhere
    End If
    sql = sql & " group by 1,2,3,4"
    sql = sql & " order by 1,2,3,4"
    
    
    Pb4.visible = True
    Pb4.Max = TotalRegistrosConsulta(sql)
    'Me.Refresh
    DoEvents
    Pb4.Value = 0
    Label2(117).visible = True
    
    Set Rs2 = New ADODB.Recordset
    Rs2.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs2.EOF
    
        IncrementarProgresNew Me.Pb4, 1
        DoEvents
        Me.Refresh

        sql = "select idplaga, (sum(kilosplaga1) + sum(kilosplaga2) + sum(kilosplaga3) + sum(kilosplaga4) + "
        sql = sql & " sum(kilosplaga5) + sum(kilosplaga6) + sum(kilosplaga7) + sum(kilosplaga8) + "
        sql = sql & " sum(kilosplaga9) + sum(kilosplaga10) + sum(kilosplaga11)) as Total  "
        sql = sql & " from rcontrol_plagas "
        sql = sql & " where codvarie = " & DBSet(Rs2!codvarie, "N")
        sql = sql & " and codsocio = " & DBSet(Rs2!Codsocio, "N")
        sql = sql & " and codcampo = " & DBSet(Rs2!codcampo, "N")
        sql = sql & " and fechacla = " & DBSet(Rs2!fechacla, "F")
        sql = sql & " group by 1 "
        sql = sql & " order by 1 "
    '    Sql = Sql & " and rcontrol_plagas.idplaga <> 2"
       
        Set RS = New ADODB.Recordset
        RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        sql = "select sum(kilosman) from rcontrol "
        sql = sql & " where codvarie = " & DBSet(Rs2!codvarie, "N")
        sql = sql & " and codsocio = " & DBSet(Rs2!Codsocio, "N")
        sql = sql & " and codcampo = " & DBSet(Rs2!codcampo, "N")
        sql = sql & " and fechacla = " & DBSet(Rs2!fechacla, "F")
        
        KilosMan = DevuelveValor(sql)
        
        KilosTot = 0
        While Not RS.EOF
            sql = "insert into tmpexcel (codusu,numalbar,fecalbar,codvarie,codsocio,codcampo,calidad1,calidad2) values ( "
            sql = sql & vUsu.Codigo & ","
            sql = sql & DBLet(RS!idplaga, "N") & ","
            sql = sql & DBSet(Rs2!fechacla, "F") & ","
            sql = sql & DBSet(Rs2!codvarie, "N") & ","
            sql = sql & DBSet(Rs2!Codsocio, "N") & ","
            sql = sql & DBSet(Rs2!codcampo, "N") & ","
            sql = sql & DBSet(RS!Total, "N") & ","
            sql = sql & DBSet(KilosMan, "N") & ")"
            
            conn.Execute sql
            If DBLet(RS!idplaga, "N") <> 2 Then KilosTot = KilosTot + DBLet(RS!Total, "N")
            RS.MoveNext
        Wend
        Set RS = Nothing
        
        sql = "update tmpexcel set kilosnet = " & DBSet(KilosTot, "N") & " where codusu = " & vUsu.Codigo
        sql = sql & " and codvarie = " & DBSet(Rs2!codvarie, "N")
        sql = sql & " and codsocio = " & DBSet(Rs2!Codsocio, "N")
        sql = sql & " and codcampo = " & DBSet(Rs2!codcampo, "N")
        sql = sql & " and fecalbar = " & DBSet(Rs2!fechacla, "F")
        
        conn.Execute sql
        
        '[Monica]26/01/2017: tema de pixat
        If SeAplicaPixat(DBLet(Rs2!codvarie), DBLet(Rs2!fechacla)) Then
            sql = "select sum(calidad1) from tmpexcel where codusu= " & vUsu.Codigo
            sql = sql & " and codvarie = " & DBSet(Rs2!codvarie, "N")
            sql = sql & " and codsocio = " & DBSet(Rs2!Codsocio, "N")
            sql = sql & " and codcampo = " & DBSet(Rs2!codcampo, "N")
            sql = sql & " and fecalbar = " & DBSet(Rs2!fechacla, "F")
            sql = sql & " and numalbar = 15 "
            
            KilosPixat = DevuelveValor(sql)
            
            sql = "update tmpexcel set calidad3 = " & DBSet(KilosPixat, "N") & " where codusu = " & vUsu.Codigo
            sql = sql & " and codvarie = " & DBSet(Rs2!codvarie, "N")
            sql = sql & " and codsocio = " & DBSet(Rs2!Codsocio, "N")
            sql = sql & " and codcampo = " & DBSet(Rs2!codcampo, "N")
            sql = sql & " and fecalbar = " & DBSet(Rs2!fechacla, "F")
            
            conn.Execute sql
        
        End If
        
    
        Rs2.MoveNext
    Wend
    
    Set Rs2 = Nothing
    
    CargarTemporalDatosDestrio = True

    Exit Function

eCargarTemporalDatosDestrio:
    MuestraError Err.Number, "Cargar Datos Temporal Destrio", Err.Description

    Pb4.visible = False
    Label2(117).visible = False
    DoEvents
End Function



Private Function CargarTemporalCampos(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim vCampAnt As CCampAnt
Dim BdAntAnterior As String
Dim KilosAntAnterior As Long
Dim KilosAnt As Long

    On Error GoTo eCargarTemporalCampos
    
    CargarTemporalCampos = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    
    sql = "select distinct rcampos.codcampo  from " & cTabla & " where " & cWhere
        
    Sql1 = "select " & vUsu.Codigo & ", rcampos.codcampo, 0,0,sum(kilosnet) from rhisfruta right join (rcampos inner join rpartida on rcampos.codparti = rpartida.codparti)  on rhisfruta.codcampo = rcampos.codcampo where rcampos.codcampo in (" & sql & ")"
    Sql1 = Sql1 & " group by 1,2,3,4"

    Sql2 = "insert into tmpinformes (codusu, importe1, importe2, importe3, importe4) " & Sql1
    conn.Execute Sql2
    
    ' Cargo los valores de las campa�as anteriores en importe1 e importe2
    Set vCampAnt = New CCampAnt
    If vCampAnt.Leer = 0 Then
        BdAntAnterior = vCampAnt.LeerAnterior(True)
    End If
        
    sql = "select * from tmpinformes where codusu = " & vUsu.Codigo
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    pb5.Max = TotalRegistrosConsulta(sql)
    pb5.visible = True
    Label2(136).visible = True
    pb5.Value = 0
    
    While Not RS.EOF
        IncrementarProgresNew pb5, 1
        DoEvents
        
        KilosAnt = 0
        KilosAntAnterior = 0
        If vCampAnt.BaseDatos <> "" Then
            sql = "select sum(kilosnet) from " & vCampAnt.BaseDatos & ".rhisfruta where codcampo = " & DBSet(RS!importe1, "N")
        
            KilosAnt = DevuelveValor(sql)
        End If

        If BdAntAnterior <> "" Then
            sql = "select sum(kilosnet) from " & BdAntAnterior & ".rhisfruta where codcampo = " & DBSet(RS!importe1, "N")
        
            KilosAntAnterior = DevuelveValor(sql)
        End If

        ' actualizamos el registro del campo actual
        sql = "update tmpinformes set importe2 = " & DBSet(KilosAntAnterior, "N")
        sql = sql & " , importe3 = " & DBSet(KilosAnt, "N")
        sql = sql & " where codusu = " & DBSet(vUsu.Codigo, "N")
        sql = sql & " and importe1 = " & DBSet(RS!importe1, "N")

        conn.Execute sql

        RS.MoveNext
    Wend
    Set RS = Nothing
    
    CargarTemporalCampos = True
    
    pb5.visible = False
    Label2(136).visible = False
    Exit Function
    
eCargarTemporalCampos:
    pb5.visible = False
    MuestraError "Cargando temporal campos", Err.Description
End Function

Private Sub ConexionConta()
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub


Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim cad As String
Dim RS As ADODB.Recordset
    
Dim vSeccion As CSeccion
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    
    If txtCodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set RS = New ADODB.Recordset
        RS.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not RS.EOF Then
            FechaIni = DBLet(RS!FechaIni, "F")
            FechaFin = DateAdd("yyyy", 1, DBLet(RS!FechaFin, "F"))
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(ind).Text, FechaFin) Then
                 cad = "El per�odo de contabilizaci�n debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 cad = cad & "    Desde: " & FechaIni & vbCrLf
                 cad = cad & "    Hasta: " & FechaFin
                 MsgBox cad, vbExclamation
                 txtCodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        RS.Close
        Set RS = Nothing
    Else
        ComprobarFechasConta = True
    End If
            
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function



Private Function DatosOkGastos(cadWHERE As String) As Boolean
Dim B As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim cta As String
Dim sql As String

   B = True

   If txtCodigo(108).Text = "" Then
        MsgBox "No se puede contabilizar, el gasto no tiene fecha. Revise.", vbExclamation
        B = False
   Else
        ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
         Orden1 = ""
         Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")

         Orden2 = ""
         Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
         FIni = CDate(Orden1)
         FFin = CDate(Orden2)
         If Not (CDate(Orden1) <= CDate(txtCodigo(108).Text) And CDate(txtCodigo(108).Text) < CDate(Day(FIni) & "/" & Month(FIni) & "/" & Year(FIni) + 2)) Then
            MsgBox "La Fecha del gasto no es del ejercicio actual ni del siguiente. Revise.", vbExclamation
            B = False
         End If
   End If


   
   'cta contable de contrapartida
   If B Then
        If txtCodigo(128).Text = "" Then
             MsgBox "Introduzca la Cta.Contable Contrapartida para contabilizar.", vbExclamation
             B = False
             PonerFoco txtCodigo(128)
        Else
             cta = ""
             cta = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", txtCodigo(128).Text, "T")
             If cta = "" Then
                 MsgBox "La cuenta contable de Contrapartida no existe. Reintroduzca.", vbExclamation
                 B = False
                 PonerFoco txtCodigo(128)
             End If
        End If
    End If
   
   'cta contable del concepto de gasto
   If B Then
        cta = DevuelveValor("select codmacgto from rconcepgasto where codgasto in ( select codgasto from rcampos_gastos where " & cadWHERE & ")")
                
        If cta = "0" Then
             MsgBox "El Concepto de Gasto no tiene una cuenta contable de gasto. Revise.", vbExclamation
             B = False
        Else
             cta = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", cta, "T")
             If cta = "" Then
                 MsgBox "La cuenta contable del concepto de Gasto no existe. Revise.", vbExclamation
                 B = False
             End If
        End If
    End If
   
   DatosOkGastos = B

End Function


Private Sub CmdAcepCambioNro_Click(Index As Integer)
Dim sql As String

    If txtCodigo(129).Text = "" Or txtCodigo(130).Text = "" Then
        MsgBox "Debe introducir un valor en el N�mero de Factura", vbExclamation
        PonerFoco txtCodigo(129)
    Else
        If txtCodigo(129).Text <> txtCodigo(130).Text Then
            MsgBox "El N�mero de Factura no coincide con la confirmaci�n. Reintroduzca.", vbExclamation
            PonerFoco txtCodigo(129)
        Else
            If CambioNroFactura(txtCodigo(130).Text, txtCodigo(131).Text) Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                cmdCancel_Click (15)
            End If
        End If
    End If
End Sub

Private Function CambioNroFactura(NuevoNro As String, NuevaFecha As String) As Boolean
Dim sql As String
Dim Mens As String
Dim Concepto As String
Dim Rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim Sql2 As String
Dim Sql3 As String

    On Error GoTo eCambioNroFactura
    
    CambioNroFactura = False
    
    conn.BeginTrans
    
    ' por si estuviera en una factura rectificativa se cambia primero en la rectificativa
    sql = "update rfactsoc aaaa, rfactsoc bbbb set "
'    Sql = Sql & " aaaa.rectif_numfactu = " & DBSet(NuevoNro, "N")
    sql = sql & " aaaa.rectif_fecfactu = " & DBSet(NuevaFecha, "F")
    sql = sql & " where aaaa.rectif_codtipom = bbbb.codtipom and aaaa.rectif_numfactu = bbbb.numfactu and aaaa.rectif_fecfactu = bbbb.fecfactu "
    conn.Execute sql
    
    
    '[Monica]02/12/2014: en el caso de picassent, preguntamos si quiere insertar en ringresos
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        Sql3 = "select codtipom,numfactu,fecfactu,codsocio from rfactsoc "
        Sql3 = Sql3 & " where " & NumCod
        Sql3 = Sql3 & " order by 1,2,3"
        
        Set rs3 = New ADODB.Recordset
        rs3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If DBLet(rs3!CodTipom, "T") = "FAT" Then
    
            Mens = "� Desea insertar en ingresos de liquidaci�n a terceros ?"
            
            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                
                Sql2 = "select aa.codtipom,aa.numfactu,aa.fecfactu,aa.codsocio,bb.codvarie,bb.imporvar from rfactsoc aa inner join rfactsoc_variedad bb on "
                Sql2 = Sql2 & " aa.codtipom = bb.codtipom and aa.numfactu = bb.numfactu and aa.fecfactu = bb.fecfactu "
                Sql2 = Sql2 & " where aa.codtipom = " & DBSet(rs3!CodTipom, "T")
                Sql2 = Sql2 & " and aa.numfactu = " & DBSet(rs3!numfactu, "N")
                Sql2 = Sql2 & " and aa.fecfactu = " & DBSet(rs3!fecfactu, "F")
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                
                Concepto = "FRA." & Trim(NuevoNro) & " DE " & NuevaFecha
                
                While Not Rs2.EOF
                    sql = "insert into ringresos (codsocio,codvarie,concepto,importe) values "
                    sql = sql & "(" & DBSet(Rs2!Codsocio, "N") & "," & DBSet(Rs2!codvarie, "N") & "," & DBSet(Concepto, "T") & ","
                    sql = sql & DBSet(Rs2!imporvar, "N") & ")"
                    
                    conn.Execute sql
                    
                    Rs2.MoveNext
                Wend
                Set Rs2 = Nothing
                
            End If
        End If
    End If
    Set rs3 = Nothing
    
    ' cabecera rfactsoc
    sql = "update rfactsoc set fecfactu = " & DBSet(NuevaFecha, "F")
    sql = sql & " ,contabilizado = 0 "
    sql = sql & " ,pdtenrofact = 2 "
    sql = sql & " ,numfacrec = " & DBSet(NuevoNro, "T")
    sql = sql & " where " & NumCod
    
    conn.Execute sql
    
    
    
    CambioNroFactura = True
    
    conn.CommitTrans
    Exit Function
    
eCambioNroFactura:
    conn.RollbackTrans
    MuestraError Err.Number, "Cambio N�mero de Factura", Err.Description
End Function

Private Function CargarVotos(vtabla As String, vSelect As String) As Boolean
Dim sql As String
Dim RS As ADODB.Recordset
Dim SqlValues As String
Dim Sql2 As String
Dim Votos As Long


    On Error GoTo eCargarVotos
    
    CargarVotos = False
    
    sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute sql
    
    '[Monica]15/01/2013: el c�lculo de campos se hace con rcampos en lugar de con rpozos
                                            'codsocio,hanegadas,votos
    Sql2 = "insert into tmpinformes (codusu, codigo1, precio1, importe2) values "
    
    '[Monica]13/03/2014: enlazamos con el propietario
    sql = "select rcampos.codpropiet, sum(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ") hanegadas from (" & vtabla & ") INNER JOIN  rcampos ON rsocios.codsocio = rcampos.codpropiet "
    If vSelect <> "" Then sql = sql & " where " & vSelect
    If sql <> "" Then
        sql = sql & " and rsocios_seccion.fecbaja is null "
    Else
        sql = sql & " where rsocios_seccion.fecbaja is null "
    End If
    sql = sql & " and rcampos.fecbajas is null "
    sql = sql & " group by 1 order by 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SqlValues = ""
    
    While Not RS.EOF
        Votos = CalculoVotos(DBLet(RS!Hanegadas, "N"))
    
        SqlValues = SqlValues & "(" & vUsu.Codigo & "," & DBSet(RS!Codpropiet, "N") & "," & DBSet(RS!Hanegadas, "N") & "," & DBSet(Votos, "N") & "),"
    
        RS.MoveNext
    Wend
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        conn.Execute Sql2 & SqlValues
    End If
    Set RS = Nothing
    
    CargarVotos = True
    Exit Function
    
eCargarVotos:
    MuestraError Err.Number, "Cargar Votos", Err.Description
End Function

Private Function CalculoVotos(Hanegadas As Currency) As Long
Dim sql As String
Dim Votos As Currency
    
    On Error Resume Next

    Select Case Hanegadas
        Case Is < 25
            Votos = Int(Hanegadas)
        Case 25 To 100
            Votos = 25 + Int((Hanegadas - 25) / 25)
            ' parte entera o fraccion
            If ((Hanegadas - 25) / 25) <> Int((Hanegadas - 25) / 25) Then Votos = Votos + 1
        Case Is > 100
            Votos = 28 + Int((Hanegadas - 100) / 100)
            ' parte entera o fraccion
            If ((Hanegadas - 100) / 100) <> Int((Hanegadas - 100) / 100) Then Votos = Votos + 1
        Case Else
            Votos = 0
    End Select
    CalculoVotos = Votos

End Function


'*************************************************************************************
'*****************************  ROPAS PARA PICASSENT  ********************************
'*************************************************************************************

Private Function GeneracionFicheroCamposPicassent(NFic As Integer, cTabla As String, vWhere As String, Regs As Integer) As Boolean
Dim Aux As String
Dim RS As ADODB.Recordset
Dim B As Boolean
Dim Lin As Integer

Dim AntSocio As Long
Dim AntPoligono As Long
Dim AntParcela As Long
Dim AntRecintos As Long

Dim AntCodconse As Long
Dim sql As String


Dim cad As String

Dim Campos As String
Dim rs3 As ADODB.Recordset
Dim Sql3 As String


    On Error GoTo eGeneracionFicheroCamposPicassent

    GeneracionFicheroCamposPicassent = False

    
    '[Monica]02/04/2014: trabajamos con la tabla intermedia para sumar supcultcatas y canaforo posteriormente
    sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute sql
    '                                       campo,    codconse, poligono, parcela,  recinto
    sql = "insert into tmpinformes (codusu, importe1, importe2, importe3, importe4, importe5) "
    sql = sql & "select " & vUsu.Codigo & ",rcampos.codcampo, variedades.codconse, rcampos_parcelas.poligono, rcampos_parcelas.parcela, rcampos_parcelas.recintos "
    sql = sql & " from (" & cTabla & ") INNER JOIN rcampos_parcelas ON rcampos.codcampo = rcampos_parcelas.codcampo "
    
    If vWhere <> "" Then
        sql = sql & " where " & vWhere
    End If
    conn.Execute sql



    Aux = "select rcampos.codsocio, rcampos_parcelas.poligono, rcampos_parcelas.parcela, rcampos_parcelas.recintos, rcampos_parcelas.subparce,"
    '[Monica]02/04/2014: antes rcampos.codvarie ahora variedades.codsigpa
    Aux = Aux & " rcampos.codparti, variedades.codconse, rsocios.nifsocio, "
    Aux = Aux & " rcampos.fecaltas, rcampos.fecbajas, rsocios_seccion.fecalta, "
    Aux = Aux & " rcampos.codcampo, rcampos.tipoparc, rcampos.refercatas, rcampos_parcelas.supsigpa, rcampos_parcelas.supcultsigpa, "
    Aux = Aux & " rcampos_parcelas.supcultcatas, rcampos.canaforo "
    Aux = Aux & " from (" & cTabla & ") INNER JOIN rcampos_parcelas ON rcampos.codcampo = rcampos_parcelas.codcampo "
    
    If vWhere <> "" Then
        Aux = Aux & " where " & vWhere
    End If
    
    Aux = Aux & " order by rcampos.codsocio, rcampos_parcelas.poligono, rcampos_parcelas.parcela, rcampos_parcelas.recintos, "
    Aux = Aux & " rcampos_parcelas.subparce, variedades.codconse" ' antes rcampos.codvarie
    
    Set RS = New ADODB.Recordset
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then
        'No hayningun registro
    Else
        B = True
        Regs = 0
        Lin = 0
        
        '[Monica]27/11/2012: Introducimos la cabecera
        cad = "Ejercicio; CifOpfh; Dni; Pais; TipoParcela; CodParcela; Provincia; Municipio; Agregado; Zona; Poligono; Parcela; Recinto; SubRecinto; SupParcela; SupRecinto; SupSubRecinto; FAlta; FBaja; Cosecha; Producto;SupCultivo;Produccion"
        Print #NFic, cad
        
        
        If Not RS.EOF Then
            AntSocio = DBLet(RS!Codsocio, "N")
            AntPoligono = DBLet(RS!Poligono, "N")
            AntParcela = DBLet(RS!Parcela, "N")
            AntRecintos = DBLet(RS!recintos, "N")
            AntCodconse = DBLet(RS!CodConse, "N")
        End If
        
        pb7.Max = TotalRegistrosConsulta(Aux)
        pb7.visible = True
        Label2(187).visible = True
        Label2(187).Caption = "Cargando Campos"
        pb7.Value = 0
        
        
        While Not RS.EOF And B
            IncrementarProgresNew pb7, 1
            DoEvents
            
            Regs = Regs + 1

            If AntSocio <> RS!Codsocio Or AntPoligono <> RS!Poligono Or AntParcela <> RS!Parcela Or AntRecintos <> RS!recintos Or AntCodconse <> RS!CodConse Then
                Lin = 0
                '[Monica]02/04/2014: cuando rompemos metemos la linea, antes era abajo

                Campos = ""
                
                Sql3 = "select importe1 from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
                Sql3 = Sql3 & " and importe2 = " & DBSet(RS!CodConse, "N") & " and importe3 = " & DBSet(RS!Poligono, "N")
                Sql3 = Sql3 & " and importe4 = " & DBSet(RS!Parcela, "N") & " and importe5 = " & DBSet(RS!recintos, "N")
                
                Set rs3 = New ADODB.Recordset
                rs3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not rs3.EOF
                    Campos = Campos & DBSet(rs3!importe1, "N") & ","
                    rs3.MoveNext
                Wend
                Set rs3 = Nothing
                If Campos <> "" Then
                    Campos = Mid(Campos, 1, Len(Campos) - 1)
                Else
                    Campos = "-1"
                End If

                B = LineaTraspasoCampoROPASPicassent(NFic, RS, Lin, Campos)
            
                AntSocio = DBLet(RS!Codsocio, "N")
                AntPoligono = DBLet(RS!Poligono, "N")
                AntParcela = DBLet(RS!Parcela, "N")
                AntRecintos = DBLet(RS!recintos, "N")
                AntCodconse = DBLet(RS!CodConse, "N")
            
            End If
            Lin = Lin + 1

    '        b = LineaTraspasoCampoROPASPicassent(NFic, RS, Lin)
            
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
            
    Close (NFic)
    
    GeneracionFicheroCamposPicassent = True
    Exit Function


eGeneracionFicheroCamposPicassent:
    MuestraError Err.Number, "Error en la Generacion de fichero de Campos.", Err.Description
End Function


Private Function LineaTraspasoCampoROPASPicassent(NFich As Integer, ByRef RS As ADODB.Recordset, Lin As Integer, Campos As String) As Boolean
Dim cad As String
Dim Areas As Long
Dim Tipo As Integer
Dim sql As String
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim CodiIVA As String
Dim TipoIRPF As Byte
Dim PorcIva As String
Dim vPorcIva As Currency
Dim CoopSoc As Currency

Dim Producto As String
Dim Variedad As String
Dim NomVar As String
Dim codVar As Long

Dim nifSocio As String
Dim Kilos As Long
Dim vPorcGasto As String
Dim vImporte As Currency

Dim Pobla As String
Dim CodSigPa As String
Dim HectaSig As Currency
Dim FecAlta As Date
Dim CodConse As Long
Dim CanAfo As Currency
Dim Super As Currency
Dim I As Integer
Dim SubParce As String

Dim HectaSigParcela As Currency
Dim HectaSigRecinto As Currency
Dim HectaSigSubRecinto As Currency
Dim HectaSubRecinto As Currency
Dim SuperLinea As Currency
Dim Produccion As Currency
Dim SupParcelas As Currency

Dim Rs2 As ADODB.Recordset
Dim Sql4 As String
Dim Sql3 As String
Dim Total As Currency


    On Error GoTo eLineaTraspasoCampoROPASPicassent
    
    LineaTraspasoCampoROPASPicassent = False

    cad = ""
    cad = cad & Format(txtCodigo(62).Text, "0000") & ";"
    cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 12) & ";"
    cad = cad & RellenaABlancos(RS!nifSocio, True, 12) & ";ES;"
    
    If RS!tipoparc = 0 Then
        cad = cad & "R;"
        cad = cad & Space(27) & ";"
    Else
        cad = cad & "U;"
        cad = cad & RellenaABlancos(DBLet(RS!refercatas, "T"), True, 27) & ";"
    End If
    
    Pobla = ""
    Pobla = DevuelveValor("select codpobla from rpartida where codparti = " & DBSet(RS!codparti, "N"))
    
    cad = cad & Mid(Pobla, 1, 2) & ";"
    
    CodSigPa = ""
    CodSigPa = DevuelveValor("select codsigpa from rpueblos where codpobla = " & DBSet(Pobla, "T"))
    
    cad = cad & Format(CodSigPa, "###") & ";"
    
    If DBLet(RS!tipoparc, "N") = 0 Then
        cad = cad & "000;"
        cad = cad & "00;"
        cad = cad & Format(DBLet(RS!Poligono, "N"), "###") & ";"
        cad = cad & Format(DBLet(RS!Parcela, "N"), "#####") & ";"
        cad = cad & Format(DBLet(RS!recintos, "N"), "#####") & ";"
        
        '[Monica]02/04/2014: cambiamos el subrecinto (antes grababamos Lin ahora Lin1)
        Dim Lin1 As Integer
        
        Select Case RS!CodConse
            Case "80130"
                Lin1 = 2
            Case "80140"
                Lin1 = 1
            Case "80110"
                Lin1 = 3
            Case Else
                Lin1 = Lin
        End Select
        cad = cad & RellenaABlancos(DBLet(Lin1, "T"), True, 2) & ";" ' antes rs!subparce
'antes
'        Cad = Cad & RellenaABlancos(DBLet(Lin, "T"), True, 2) & ";" ' antes rs!subparce


    Else
        cad = cad & ";;;;;;"
    End If
        
    HectaSig = DBLet(RS!supsigpa, "N") '  SUPERFICIE TOTAL PARCELA
    
    cad = cad & Format(HectaSig, "##0.0000") & ";"
    
    HectaSigRecinto = DBLet(RS!supcultsigpa, "N") '  SUPERFICIE TOTAL RECINTO
    
    cad = cad & Format(HectaSigRecinto, "##0.0000") & ";"
    
    '[Monica]02/04/2014: sumamos las superficies antes era la rs!supcultcatas
    Sql3 = "select sum(supcultcatas) from rcampos_parcelas where poligono = " & DBSet(RS!Poligono, "N") & " and parcela = " & DBSet(RS!Parcela, "N") & " and recintos = " & DBSet(RS!recintos, "N")
    Sql3 = Sql3 & " and codcampo in (" & Campos & ")"
    Total = DevuelveValor(Sql3)
    HectaSubRecinto = Total
'    HectaSubRecinto = DBLet(RS!supcultcatas, "N")
    cad = cad & Format(HectaSubRecinto, "##0.0000") & ";"
    
    FecAlta = DBLet(RS!FecAltas, "F")
    '[Monica]23/01/2013: si la fecha de alta del campo es anterior a la fecha de alta de socio
    '                    que ponga la fecha de alta del socio
    If RS!FecAlta > RS!FecAltas Then ' fecha alta socio > fecha alta campo
        FecAlta = RS!FecAlta
    End If
    
    cad = cad & Format(FecAlta, "dd/mm/yyyy") & ";"
    If DBLet(RS!fecbajas) <> "" Then
        cad = cad & Format(RS!fecbajas, "dd/mm/yyyy") & ";"
    Else
        cad = cad & ";"
    End If
        
    cad = cad & Format(1, "#") & ";"  ' contador de subparcelas
    
'[Monica]02/04/2014: el codigo de conselleria lo tenemos en el select antes codconse ahora rs!codconse
'    CodConse = 0
'    CodConse = DevuelveValor("select codconse from variedades where codvarie = " & DBSet(RS!codvarie, "N"))
    
    cad = cad & RellenaABlancos(CStr(DBLet(RS!CodConse, "N")), True, 6) & ";"
    
    cad = cad & Format(HectaSubRecinto, "##0.0000") & ";"
    
'    '[Monica]14/02/2013: la produccion vamos a poner que es la real
'    Sql4 = "select sum(kilosnet) from rhisfruta where codcampo = " & DBSet(RS!CodCampo, "N") & " and codvarie = " & DBSet(RS!CodVarie, "N")
'    Sql4 = Sql4 & " and codsocio = " & DBSet(RS!CodSocio, "N")
'
'    Produccion = DevuelveValor(Sql4)

    '[Monica]15/02/2013: no es la produccion real es la estimada (canaforo)
'    Produccion = DBLet(RS!Canaforo, "N")

    '[Monica]02/04/2014: sumamos los canaforo de los campos que intervienen
    Sql3 = "select sum(canaforo) from rcampos where codcampo in (" & Campos & ")"
    Produccion = DevuelveValor(Sql3)

    SupParcelas = DevuelveValor("select sum(supcultsigpa) from rcampos_parcelas where codcampo = " & DBSet(RS!codcampo, "N"))
    CanAfo = 0
    If SupParcelas <> 0 Then
        CanAfo = Round2((DBLet(RS!supcultsigpa, "N") * Produccion / SupParcelas) / 1000, 2)  'En toneladas
    End If
    
    cad = cad & Format(CanAfo, "###0.00")
    
    Print #NFich, cad
    
    LineaTraspasoCampoROPASPicassent = True
    Exit Function
    
eLineaTraspasoCampoROPASPicassent:
    MuestraError Err.Number, "Carga Linea de Traspaso Campos ROPAS", Err.Description
End Function





Private Function GenerarEntradasSIN(pTabla As String, pWhere As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim vSocio As cSocio
Dim B As Boolean
Dim Nregs As Long
Dim Total As Variant

Dim cTabla As String
Dim vWhere As String
Dim sql As String

    On Error GoTo EGen
    
    GenerarEntradasSIN = False
    
    conn.BeginTrans
    
    
    cTabla = pTabla
    vWhere = pWhere
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = QuitarCaracterACadena(cTabla, "_1")
    If vWhere <> "" Then
        vWhere = QuitarCaracterACadena(vWhere, "{")
        vWhere = QuitarCaracterACadena(vWhere, "}")
        vWhere = QuitarCaracterACadena(vWhere, "_1")
    End If
    
        
    sql = "insert into rhisfrutasin (numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,kilosbru,numcajon,kilosnet,impreso,impentrada,cobradosn, "
    sql = sql & " transportadopor,kilostra,esbonifespecial,estarepcooprop) "
    sql = sql & " select aaa.numfactu + 9000000, " & DBSet(txtCodigo(135).Text, "F") & ", bbb.codvarie, aaa.codsocio,  bbb.codcampo, 0,0, sum(bbb.kilosnet) bruto,round(sum(bbb.kilosnet) / 20,0) cajas,"
    sql = sql & " sum(bbb.kilosnet),0,0,0,0,sum(bbb.kilosnet),0,0"
    sql = sql & " from (rfactsoc aaa inner join rfactsoc_variedad bbb on aaa.CodTipom = bbb.CodTipom And aaa.numfactu = bbb.numfactu "
    sql = sql & " and aaa.fecfactu = bbb.fecfactu) inner join rfactsoc_calidad ccc"
    sql = sql & " on aaa.codtipom = ccc.codtipom and aaa.numfactu = ccc.numfactu and aaa.fecfactu = ccc.fecfactu and "
    sql = sql & " bbb.CodVarie = ccc.CodVarie And bbb.CodCampo = ccc.CodCampo"
    sql = sql & " where " & vWhere
    sql = sql & " group by 1,2,3,4,5,6,7 "

    conn.Execute sql


    sql = "insert into rhisfrutasin_entradas (numalbar,numnotac,fechaent,horaentr,kilosbru,numcajon,kilosnet,kilostra,tiporecol) "
    sql = sql & " select aaa.numfactu + 9000000, aaa.numfactu + 9000000, DATE('" & Format(txtCodigo(135).Text, "yyyy-mm-dd") & " 09:02:00' + INTERVAL aaa.numfactu / 3 * RAND() HOUR) fecha, "
    sql = sql & " ('" & Format(txtCodigo(135).Text, "yyyy-mm-dd") & " 06:02:00' + INTERVAL aaa.numfactu * RAND() MINUTE) hora, bbb.kilosnet, round(bbb.kilosnet / 20,0),bbb.KilosNet , bbb.KilosNet, 0"
    sql = sql & " from rfactsoc aaa inner join rfactsoc_variedad bbb on   aaa.codtipom = bbb.codtipom and aaa.numfactu = bbb.numfactu "
    sql = sql & " and aaa.fecfactu = bbb.fecfactu where " & vWhere
    
    conn.Execute sql


    sql = "Update rhisfrutasin_entradas"
    sql = sql & " set rhisfrutasin_entradas.horaentr = concat(rhisfrutasin_entradas.fechaent, ' ', time(rhisfrutasin_entradas.horaentr))"
    sql = sql & " where rhisfrutasin_entradas.numalbar >= 9000000;"

    conn.Execute sql

    sql = "insert into rhisfrutasin_clasif (numalbar,codvarie,codcalid,kilosnet) "
    sql = sql & " select aaa.numfactu + 9000000, bbb.codvarie, ccc.codcalid, ccc.kilosnet "
    sql = sql & " from (rfactsoc aaa inner join rfactsoc_variedad bbb on aaa.CodTipom = bbb.CodTipom And aaa.numfactu = bbb.numfactu "
    sql = sql & " and aaa.fecfactu = bbb.fecfactu) inner join rfactsoc_calidad ccc on aaa.codtipom = ccc.codtipom and aaa.numfactu = ccc.numfactu and aaa.fecfactu = ccc.fecfactu and "
    sql = sql & " bbb.CodVarie = ccc.CodVarie And bbb.CodCampo = ccc.CodCampo "
    sql = sql & " where " & vWhere

    conn.Execute sql


    sql = "update rhisfrutasin_entradas aaa, rhisfrutasin bbb "
    sql = sql & " set aaa.horaentr = concat(bbb.fecalbar, ' ', time(aaa.horaentr)), "
    sql = sql & " aaa.FechaEnt = bbb.fecalbar"
    sql = sql & " Where aaa.Numalbar >= 9000000 And aaa.Numalbar = bbb.Numalbar And aaa.FechaEnt > bbb.fecalbar"

    conn.Execute sql
    
    conn.CommitTrans
    
    GenerarEntradasSIN = True
    Exit Function
    
EGen:
    conn.RollbackTrans
    MuestraError Err.Number, "Generaci�n Entradas", Err.Description
End Function

Private Function CargarTemporalMiembros(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql3 As String
    
    On Error GoTo eCargarTemporal
    
    CargarTemporalMiembros = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "Select rsocios.codsocio FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    
    Sql3 = "select " & vUsu.Codigo & ",2, 1, codsocio, nifmiembro, nommiembro, votos, capital from rsocios_miembros where codsocio in (" & sql & ")"
    
    ' miembros                              'miembro,productor,socio, nif,     nombre,  votos,    capital
    Sql2 = "insert into tmpinformes (codusu, campo1, campo2, codigo1, nombre1, nombre2, importe1, importe2) "
    conn.Execute Sql2 & Sql3
    ' socios
    Sql3 = "select distinct " & vUsu.Codigo & ",1,if(tipoprod = 4,0,1), rsocios.codsocio, rsocios.nifsocio, rsocios.nomsocio, rsocios.votos, rsocios.capital from " & cTabla
    If cWhere <> "" Then Sql3 = Sql3 & " where " & cWhere
    conn.Execute Sql2 & Sql3
            
    CargarTemporalMiembros = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function




Private Function CargarTemporal7(cTabla As String, cWhere As String, cTabla2 As String) As Boolean
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql3 As String
Dim SocioAnt As Long
Dim CampoAnt As Long
Dim SocioAct As Long
Dim CampoAct As Long
Dim SQLaux As String
Dim SqlAux2 As String
Dim Ha As Currency
Dim Producto As String
    
'[Monica]13/11/2013: prorrateamos segun los coopropietarios
Dim Porcen As Currency
Dim Canaforo As String
Dim KilosAse As String

Dim Hanegadas As Currency
Dim Hectareas As Currency
Dim Arboles As Long
                    
Dim DCanaforo As Long
Dim DkilosAse As Long
Dim KilosTot As Currency
Dim Sql4 As String

                    
    
    On Error GoTo eCargarTemporal
    
    CargarTemporal7 = False

    Sql2 = "delete from tmpinfkilos where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
    End If
    
    sql = "select distinct rcampos.codsocio, rcampos.codcampo "
    sql = sql & " from " & cTabla
    sql = sql & " where rcampos.fecbajas is null "
    If cWhere <> "" Then
        sql = sql & " and " & cWhere
    End If
    sql = sql & " union "
    sql = sql & " select distinct rhisfruta.codsocio, rhisfruta.codcampo "
    sql = sql & " from (" & cTabla & ") inner join rhisfruta on rcampos.codcampo = rhisfruta.codcampo and rcampos.codsocio = rhisfruta.codsocio "
    If cWhere <> "" Then
        sql = sql & " where " & cWhere
    End If
    If txtCodigo(184).Text <> "" Then sql = sql & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(184).Text, "F")
    If txtCodigo(185).Text <> "" Then sql = sql & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(185).Text, "F")
    
    '[Monica]13/11/2013: faltan los medieros para sacar los kilos de las entradas
    sql = sql & " union "
    sql = sql & " select distinct rhisfruta.codsocio, rhisfruta.codcampo "
    sql = sql & " from (" & cTabla2 & ") inner join rhisfruta on rcampos_cooprop.codcampo = rhisfruta.codcampo and rcampos_cooprop.codsocio = rhisfruta.codsocio "
    If cWhere <> "" Then
        sql = sql & " where " & cWhere
    End If
    If txtCodigo(184).Text <> "" Then sql = sql & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(184).Text, "F")
    If txtCodigo(185).Text <> "" Then sql = sql & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(185).Text, "F")
    
    
    sql = sql & " order by 1, 2"
    
    
    pb11.visible = True
    Label2(267).visible = True
    
    Sql4 = "select count(*) from (" & sql & ") aaaa"
    CargarProgres pb11, DevuelveValor(Sql4)
    
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql2 = "insert into tmpinfkilos (codusu, codsocio, codcampo, kilosnet, "
    Sql2 = Sql2 & "canaforo, nroarbol) values "
    
    While Not RS.EOF
        IncrementarProgres pb11, 1
    
        SocioAct = DBLet(RS.Fields(0).Value, "N")
        CampoAct = DBLet(RS.Fields(1).Value, "N")
            
        Sql3 = "(" & vUsu.Codigo & "," & DBSet(RS.Fields(0).Value, "N") & "," & DBSet(RS.Fields(1).Value, "N") & ","
        
        SQLaux = "select sum(kilosnet) from rhisfruta where codsocio = " & DBSet(RS.Fields(0).Value, "N")
        SQLaux = SQLaux & " and codcampo = " & DBSet(RS.Fields(1).Value, "N")
        If txtCodigo(39).Text <> "" Then SQLaux = SQLaux & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(184).Text, "F")
        If txtCodigo(40).Text <> "" Then SQLaux = SQLaux & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(185).Text, "F")
        
        KilosTot = DevuelveValor(SQLaux)
        
        SQLaux = "select sum(kilosnet) from rclasifica where codsocio = " & DBSet(RS.Fields(0).Value, "N")
        SQLaux = SQLaux & " and codcampo = " & DBSet(RS.Fields(1).Value, "N")
        If txtCodigo(39).Text <> "" Then SQLaux = SQLaux & " and rclasifica.fechaent >= " & DBSet(txtCodigo(184).Text, "F")
        If txtCodigo(40).Text <> "" Then SQLaux = SQLaux & " and rclasifica.fechaent <= " & DBSet(txtCodigo(185).Text, "F")
        
        KilosTot = KilosTot + DevuelveValor(SQLaux)
        
        SQLaux = "select sum(kilosnet) from rentradas where codsocio = " & DBSet(RS.Fields(0).Value, "N")
        SQLaux = SQLaux & " and codcampo = " & DBSet(RS.Fields(1).Value, "N")
        If txtCodigo(39).Text <> "" Then SQLaux = SQLaux & " and rentradas.fechaent >= " & DBSet(txtCodigo(184).Text, "F")
        If txtCodigo(40).Text <> "" Then SQLaux = SQLaux & " and rentradas.fechaent <= " & DBSet(txtCodigo(185).Text, "F")
        
        KilosTot = KilosTot + DevuelveValor(SQLaux)
        
       
        Sql3 = Sql3 & DBSet(KilosTot, "N") & "," 'kilosnet
        
        SqlAux2 = "select canaforo, kilosase "
        SqlAux2 = SqlAux2 & " from rcampos where codcampo = " & DBSet(RS.Fields(1).Value, "N")
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open SqlAux2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs1.EOF Then
            '[Monica]13/11/2013: a�adimos el porcentaje de coopropiedad
            Porcen = PorCoopropiedadCampo(RS.Fields(1).Value, RS.Fields(0).Value) / 100
            If Porcen <> 0 Then
        
                Canaforo = Round2(DBLet(Rs1.Fields(0).Value, "N") * Porcen, 0)
                KilosAse = Round2(DBLet(Rs1.Fields(1).Value, "N") * Porcen, 0)
                
                Sql3 = Sql3 & DBSet(Canaforo, "N") & ","
                Sql3 = Sql3 & DBSet(KilosAse, "N") & "),"
                
        
            Else
                ' si no hay coopropietarios es todo suyo
            
                Sql3 = Sql3 & DBSet(Rs1.Fields(0).Value, "N") & "," 'canaforo
                Sql3 = Sql3 & DBSet(Rs1.Fields(1).Value, "N") & ")," 'kilosase
        
            End If
            
        
        Else
            Sql3 = Sql3 & "0,0),"
        End If
        
        Sql2 = Sql2 & Sql3
        
        Set Rs1 = Nothing
        
        
        RS.MoveNext
    Wend

    'quitamos la ultima coma
    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
    conn.Execute Sql2
    
    '[Monica]13/11/2013: puede que hayan errores en el prorrateo de arboles y canaforo, se lo daremos al
    sql = "select codcampo, sum(canaforo) canaforo, sum(nroarbol) kilosase from tmpinfkilos where codusu = " & vUsu.Codigo & " group by codcampo order by codcampo "
    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        sql = "select codsocio, canaforo, kilosase "
        sql = sql & " from rcampos where codcampo = " & DBSet(RS!codcampo, "N")
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs1.EOF Then
            DCanaforo = DBLet(RS!Canaforo, "N") - DBLet(Rs1!Canaforo, "N")
            DkilosAse = DBLet(RS!KilosAse, "N") - DBLet(Rs1.Fields(2).Value, "N")
        
            sql = "update tmpinfkilos set "
            sql = sql & " canaforo = canaforo + " & DBSet(DCanaforo, "N")
            sql = sql & " ,nroarbol = nroarbol + " & DBSet(DkilosAse, "N")
            sql = sql & " where codusu = " & vUsu.Codigo
            sql = sql & " and codcampo = " & DBSet(RS!codcampo, "N")
            sql = sql & " and codsocio = " & DBSet(Rs1!Codsocio, "N")
        
            conn.Execute sql
        End If
        
        RS.MoveNext
    Wend
    
    CargarTemporal7 = True
    
    pb11.visible = False
    Label2(267).visible = False
    
    Exit Function
    
eCargarTemporal:
    pb11.visible = False
    Label2(267).visible = False
    MuestraError "Cargando temporal", Err.Description
End Function


