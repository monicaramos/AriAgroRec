VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmListAnticipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7500
   Icon            =   "frmListAnticipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAnticipos 
      Height          =   5640
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Frame FrameRecolectado 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   3210
         TabIndex        =   164
         Top             =   2880
         Width           =   2865
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   165
            Tag             =   "Recolectado|N|N|0|1|rhisfruta|recolect||N|"
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label Label11 
            Caption         =   "Recolectado"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   166
            Top             =   150
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdAceptarAntGastos 
         Caption         =   "&AcepGast"
         Height          =   375
         Left            =   2610
         TabIndex        =   8
         Top             =   4950
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame FrameOpciones 
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   3840
         TabIndex        =   102
         Top             =   3690
         Width           =   2115
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   104
            Top             =   30
            Width           =   1965
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Factura"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   103
            Top             =   420
            Width           =   1995
         End
      End
      Begin VB.Frame FrameFechaAnt 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   645
         Left            =   390
         TabIndex        =   26
         Top             =   3750
         Width           =   3045
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   930
            Picture         =   "frmListAnticipos.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Anticipo"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   25
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListAnticipos.frx":0097
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListAnticipos.frx":03A1
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1575
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   1215
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1575
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1215
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptarAnt 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   7
         Top             =   4950
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelAnt 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5190
         TabIndex        =   9
         Top             =   4935
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3405
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3000
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   4560
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1320
         MouseIcon       =   "frmListAnticipos.frx":06AB
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2550
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1320
         MouseIcon       =   "frmListAnticipos.frx":07FD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   735
         TabIndex        =   25
         Top             =   2595
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   735
         TabIndex        =   24
         Top             =   2205
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   23
         Top             =   1950
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListAnticipos.frx":094F
         ToolTipText     =   "Buscar fecha"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListAnticipos.frx":09DA
         ToolTipText     =   "Buscar fecha"
         Top             =   3405
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1320
         MouseIcon       =   "frmListAnticipos.frx":0A65
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1605
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1320
         MouseIcon       =   "frmListAnticipos.frx":0BB7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   360
         TabIndex        =   20
         Top             =   1020
         Width           =   405
      End
      Begin VB.Label Label3 
         Caption         =   "Listado de Anticipos"
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
         TabIndex        =   19
         Top             =   465
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   690
         TabIndex        =   18
         Top             =   1620
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   690
         TabIndex        =   17
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   705
         TabIndex        =   16
         Top             =   3405
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   705
         TabIndex        =   15
         Top             =   3060
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   375
         TabIndex        =   14
         Top             =   2820
         Width           =   450
      End
   End
   Begin VB.Frame FrameGeneraFactura 
      Height          =   5790
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   6585
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   4110
         TabIndex        =   99
         Top             =   3900
         Width           =   1965
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Factura"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   101
            Top             =   360
            Width           =   1545
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   100
            Top             =   0
            Width           =   1635
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   79
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   78
         Top             =   3810
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   77
         Top             =   3270
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   75
         Top             =   2385
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "Text5"
         Top             =   2025
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "Text5"
         Top             =   2385
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   74
         Top             =   2010
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   76
         Top             =   2880
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Text5"
         Top             =   2880
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Text5"
         Top             =   3270
         Width           =   3285
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "frmListAnticipos.frx":0D09
         Left            =   1920
         List            =   "frmListAnticipos.frx":0D0B
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Tag             =   "Recolección|N|N|0|3|rhisfruta|recolect|||"
         Top             =   960
         Width           =   1425
      End
      Begin VB.CommandButton CmdAcepGenFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3930
         TabIndex        =   80
         Top             =   5145
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelGenFac 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5100
         TabIndex        =   81
         Top             =   5145
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   73
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1560
         Width           =   1050
      End
      Begin MSComctlLib.ProgressBar Pb3 
         Height          =   255
         Left            =   420
         TabIndex        =   68
         Top             =   4710
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Entrada"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   9
         Left            =   420
         TabIndex        =   94
         Top             =   3630
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   8
         Left            =   975
         TabIndex        =   93
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   975
         TabIndex        =   92
         Top             =   4215
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   930
         TabIndex        =   91
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   930
         TabIndex        =   90
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   4
         Left            =   420
         TabIndex        =   89
         Top             =   1830
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1590
         MouseIcon       =   "frmListAnticipos.frx":0D0D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1590
         MouseIcon       =   "frmListAnticipos.frx":0E5F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2010
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1590
         Picture         =   "frmListAnticipos.frx":0FB1
         ToolTipText     =   "Buscar fecha"
         Top             =   4215
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1590
         Picture         =   "frmListAnticipos.frx":103C
         ToolTipText     =   "Buscar fecha"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   88
         Top             =   2700
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   975
         TabIndex        =   87
         Top             =   2955
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   975
         TabIndex        =   86
         Top             =   3330
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1590
         MouseIcon       =   "frmListAnticipos.frx":10C7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1590
         MouseIcon       =   "frmListAnticipos.frx":1219
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2910
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   3
         Left            =   420
         TabIndex        =   71
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Generación de Factura Venta Campo"
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
         TabIndex        =   70
         Top             =   360
         Width           =   5940
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   1590
         Picture         =   "frmListAnticipos.frx":136B
         ToolTipText     =   "Buscar fecha"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   6
         Left            =   420
         TabIndex        =   69
         Top             =   1380
         Width           =   1815
      End
   End
   Begin VB.Frame FrameReimpresion 
      Height          =   5220
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   6675
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   3780
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text5"
         Top             =   3405
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   35
         Top             =   3780
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   34
         Top             =   3405
         Width           =   830
      End
      Begin VB.CommandButton cmdAceptarReimp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   36
         Top             =   4275
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelReimp 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5340
         TabIndex        =   38
         Top             =   4275
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   33
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2775
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1740
         MaxLength       =   7
         TabIndex        =   30
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1365
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   31
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1755
         Width           =   830
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1110
         Index           =   0
         Left            =   3180
         TabIndex        =   97
         Top             =   1350
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1958
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   5
         Left            =   3180
         TabIndex        =   98
         Top             =   1110
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   6060
         Picture         =   "frmListAnticipos.frx":13F6
         ToolTipText     =   "Desmarcar todos"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   5820
         Picture         =   "frmListAnticipos.frx":1DF8
         ToolTipText     =   "Marcar todos"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1470
         MouseIcon       =   "frmListAnticipos.frx":864A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1470
         MouseIcon       =   "frmListAnticipos.frx":879C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3405
         Width           =   240
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
         Index           =   11
         Left            =   510
         TabIndex        =   49
         Top             =   3165
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   870
         TabIndex        =   48
         Top             =   3780
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   855
         TabIndex        =   47
         Top             =   3405
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1485
         Picture         =   "frmListAnticipos.frx":88EE
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1485
         Picture         =   "frmListAnticipos.frx":8979
         ToolTipText     =   "Buscar fecha"
         Top             =   2415
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   825
         TabIndex        =   46
         Top             =   2775
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   825
         TabIndex        =   45
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   465
         TabIndex        =   44
         Top             =   2115
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Factura"
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
         Left            =   495
         TabIndex        =   43
         Top             =   1125
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   42
         Top             =   1395
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   41
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Reimpresión de Facturas Socios"
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
         TabIndex        =   40
         Top             =   315
         Width           =   5160
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
   Begin VB.Frame FrameResultados 
      Height          =   5490
      Left            =   0
      TabIndex        =   105
      Top             =   0
      Width           =   6675
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen"
         Height          =   255
         Index           =   4
         Left            =   810
         TabIndex        =   114
         Top             =   4170
         Width           =   1965
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   130
         Text            =   "Text5"
         Top             =   2460
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   28
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   129
         Text            =   "Text5"
         Top             =   2070
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1725
         MaxLength       =   7
         TabIndex        =   111
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2445
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1725
         MaxLength       =   7
         TabIndex        =   110
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2055
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   113
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3525
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   112
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3165
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancelResul 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5340
         TabIndex        =   117
         Top             =   4635
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepResul 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   116
         Top             =   4635
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   109
         Top             =   1485
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   108
         Top             =   1110
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   107
         Text            =   "Text5"
         Top             =   1485
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   106
         Text            =   "Text5"
         Top             =   1110
         Width           =   3675
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1110
         Index           =   1
         Left            =   3150
         TabIndex        =   115
         Top             =   3120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1958
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   29
         Left            =   1440
         MouseIcon       =   "frmListAnticipos.frx":8A04
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2460
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1440
         MouseIcon       =   "frmListAnticipos.frx":8B56
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Informe de Resultados"
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
         TabIndex        =   128
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   795
         TabIndex        =   127
         Top             =   2445
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   795
         TabIndex        =   126
         Top             =   2085
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
         TabIndex        =   125
         Top             =   1830
         Width           =   390
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   23
         Left            =   435
         TabIndex        =   124
         Top             =   2865
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   795
         TabIndex        =   123
         Top             =   3165
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   795
         TabIndex        =   122
         Top             =   3540
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmListAnticipos.frx":8CA8
         ToolTipText     =   "Buscar fecha"
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   9
         Left            =   1440
         Picture         =   "frmListAnticipos.frx":8D33
         ToolTipText     =   "Buscar fecha"
         Top             =   3180
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   795
         TabIndex        =   121
         Top             =   1155
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   19
         Left            =   795
         TabIndex        =   120
         Top             =   1530
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
         Index           =   18
         Left            =   435
         TabIndex        =   119
         Top             =   915
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1440
         MouseIcon       =   "frmListAnticipos.frx":8DBE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1440
         MouseIcon       =   "frmListAnticipos.frx":8F10
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   5790
         Picture         =   "frmListAnticipos.frx":9062
         ToolTipText     =   "Marcar todos"
         Top             =   2850
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   6030
         Picture         =   "frmListAnticipos.frx":F8B4
         ToolTipText     =   "Desmarcar todos"
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   7
         Left            =   3150
         TabIndex        =   118
         Top             =   2880
         Width           =   1815
      End
   End
   Begin VB.Frame FrameDesFacturacion 
      Height          =   4740
      Left            =   30
      TabIndex        =   51
      Top             =   0
      Width           =   6555
      Begin VB.Frame FrameTipoFactura 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   390
         TabIndex        =   95
         Top             =   1410
         Width           =   3615
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            ItemData        =   "frmListAnticipos.frx":102B6
            Left            =   1380
            List            =   "frmListAnticipos.frx":102B8
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Tag             =   "Recolección|N|N|0|3|rhisfruta|recolect|||"
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Factura"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   96
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   2475
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   62
         Tag             =   "admon"
         Top             =   1170
         Width           =   1545
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   54
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2685
         Width           =   945
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   53
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   55
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3360
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelDesF 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4860
         TabIndex        =   57
         Top             =   4125
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepDesF 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   56
         Top             =   4125
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   255
         Left            =   420
         TabIndex        =   66
         Top             =   3780
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Este proceso borra facturas correlativas "
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
         Left            =   210
         TabIndex        =   65
         Top             =   450
         Width           =   5820
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Actualiza contadores"
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
         Left            =   105
         TabIndex        =   64
         Top             =   780
         Width           =   5595
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
         Left            =   1440
         TabIndex        =   63
         Top             =   1170
         Width           =   2235
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   17
         Left            =   900
         TabIndex        =   61
         Top             =   2685
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   10
         Left            =   900
         TabIndex        =   60
         Top             =   2325
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Factura"
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
         Left            =   495
         TabIndex        =   59
         Top             =   2055
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   8
         Left            =   465
         TabIndex        =   58
         Top             =   3045
         Width           =   1815
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1470
         Picture         =   "frmListAnticipos.frx":102BA
         ToolTipText     =   "Buscar fecha"
         Top             =   3360
         Width           =   240
      End
   End
   Begin VB.Frame FrameGrabacionModelos 
      Height          =   6225
      Left            =   0
      TabIndex        =   131
      Top             =   0
      Width           =   6675
      Begin VB.Frame FrameContacto 
         Caption         =   "Persona de Contacto"
         ForeColor       =   &H00972E0B&
         Height          =   915
         Left            =   420
         TabIndex        =   157
         Top             =   3510
         Width           =   5865
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   37
            Left            =   4470
            MaxLength       =   9
            TabIndex        =   139
            Tag             =   "Campol|N|S|||clientes|codposta|000000000||"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   36
            Left            =   150
            MaxLength       =   40
            TabIndex        =   138
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   510
            Width           =   4260
         End
         Begin VB.Label Label4 
            Caption         =   "Teléfono"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   4530
            TabIndex        =   159
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label4 
            Caption         =   "Apellidos y Nombre"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   210
            TabIndex        =   158
            Top             =   300
            Width           =   2595
         End
      End
      Begin VB.Frame FrameDomicilio 
         Caption         =   "Domicilio Presentador"
         ForeColor       =   &H00972E0B&
         Height          =   855
         Left            =   420
         TabIndex        =   156
         Top             =   4560
         Width           =   5895
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   40
            Left            =   150
            MaxLength       =   2
            TabIndex        =   140
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   480
            Width           =   450
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   39
            Left            =   4710
            MaxLength       =   5
            TabIndex        =   142
            Tag             =   "Campol|N|S|||clientes|codposta|00000||"
            Top             =   480
            Width           =   780
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   38
            Left            =   780
            MaxLength       =   20
            TabIndex        =   141
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   480
            Width           =   3840
         End
         Begin VB.Label Label4 
            Caption         =   "Siglas"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   39
            Left            =   150
            TabIndex        =   162
            Top             =   270
            Width           =   435
         End
         Begin VB.Label Label4 
            Caption         =   "Número"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   4740
            TabIndex        =   161
            Top             =   270
            Width           =   705
         End
         Begin VB.Label Label4 
            Caption         =   "Nombre"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   780
            TabIndex        =   160
            Top             =   270
            Width           =   2595
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   1710
         MaxLength       =   13
         TabIndex        =   137
         Tag             =   "Campol|N|S|||clientes|codposta|0000000000000||"
         Top             =   3090
         Width           =   1380
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   30
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   136
         Tag             =   "Campol|N|S|||clientes|codposta|0000||"
         Top             =   2730
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   35
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   146
         Text            =   "Text5"
         Top             =   1350
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   34
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   145
         Text            =   "Text5"
         Top             =   975
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   133
         Top             =   1350
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   132
         Top             =   975
         Width           =   830
      End
      Begin VB.CommandButton CmdAcepModelo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   143
         Top             =   5640
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelModelo 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5310
         TabIndex        =   144
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   135
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2325
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   32
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   134
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1935
         Width           =   1050
      End
      Begin ComctlLib.StatusBar BarraEst 
         Height          =   285
         Left            =   0
         TabIndex        =   163
         Top             =   6150
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   503
         Style           =   1
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   1
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Nro.Justific."
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   28
         Left            =   420
         TabIndex        =   155
         Top             =   3120
         Width           =   945
      End
      Begin VB.Label Label4 
         Caption         =   "Ejercicio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   420
         TabIndex        =   154
         Top             =   2760
         Width           =   705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   35
         Left            =   1440
         MouseIcon       =   "frmListAnticipos.frx":10345
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   34
         Left            =   1440
         MouseIcon       =   "frmListAnticipos.frx":10497
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   975
         Width           =   240
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
         Left            =   435
         TabIndex        =   153
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   34
         Left            =   795
         TabIndex        =   152
         Top             =   1380
         Width           =   420
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
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   12
         Left            =   1440
         Picture         =   "frmListAnticipos.frx":105E9
         ToolTipText     =   "Buscar fecha"
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   11
         Left            =   1440
         Picture         =   "frmListAnticipos.frx":10674
         ToolTipText     =   "Buscar fecha"
         Top             =   1950
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   32
         Left            =   795
         TabIndex        =   150
         Top             =   2370
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   795
         TabIndex        =   149
         Top             =   1995
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   30
         Left            =   435
         TabIndex        =   148
         Top             =   1695
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Grabación Modelo"
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
         Left            =   420
         TabIndex        =   147
         Top             =   270
         Width           =   5160
      End
   End
End
Attribute VB_Name = "frmListAnticipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados / Procesos ANTICIPOS ====
    '=============================
    ' 1 .- Informe de Anticipos
    ' 2 .- Prevision de Pagos de Anticipos
    ' 3 .- Facturación de Anticipos
    ' 5 .- Deshacer proceso de Facturación Anticipos
    
    
    '==== Listados / Procesos FACTURAS SOCIOS ====
    '==================================
    ' 4 .- Reimpresion de Facturas
    ' 8 .- Informe de Resultados
    ' 9 .- Informe de Retenciones
    
    ' 10.- Grabacion Modelo 190
    ' 11.- Grabación Modelo 346
    
    '==== Listados / Procesos VENTA CAMPO ====
    '=============================
    ' 6 .- Facturación de Venta Campo (Anticipo o Liquidación)
    ' 7 .- Deshacer proceso de Facturación de Venta Campo (Anticipo o Liquidación)
    
    '==== Listados / Procesos LIQUIDACIONES ====
    '================================
    ' 12 .- Informe de Liquidaciones
    ' 13 .- Prevision de Pagos de Liquidacion
    ' 14 .- Facturación de Liquidacion
    ' 15 .- Deshacer proceso de Facturación Anticipos
    
Public AnticipoGastos As Boolean ' si true entonces es que se trata de anticipos de gastos de recoleccion

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


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

Private Sub CmdAcepDesF_Click()
Dim tipo As Byte
    If DatosOk Then
        Pb2.visible = True
        Select Case OpcionListado
            Case 5 ' anticipo
                tipo = 0
            Case 7
                ' venta campo
                Select Case Combo1(1).ListIndex
                    Case 0 ' anticipo
                        tipo = 1
                    Case 1 ' liquidacion
                        tipo = 2
                End Select
            Case 15 ' liquidacion
                tipo = 3
        End Select
        If DeshacerFacturacion(tipo, txtcodigo(9).Text, txtcodigo(10).Text, txtcodigo(11).Text, Pb2) Then
            MsgBox "Proceso realizado correctamente", vbExclamation
            cmdCancelDesF_Click
        End If
    End If
End Sub

Private Sub CmdAcepGenFac_Click()
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

Dim NRegs As Long
Dim FecFac As Date
Dim tipoMov As String

Dim vSQL As String

    vSQL = ""
    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(16).Text)
        cHasta = Trim(txtcodigo(17).Text)
        nDesde = txtNombre(16).Text
        nHasta = txtNombre(17).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(18).Text)
        cHasta = Trim(txtcodigo(19).Text)
        nDesde = txtNombre(18).Text
        nHasta = txtNombre(19).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        If txtcodigo(18).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(18).Text, "N")
        If txtcodigo(19).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(19).Text, "N")
        
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(22).Text)
        cHasta = Trim(txtcodigo(23).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.SeccionHorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.SeccionHorto) Then Exit Sub
        
        'sólo entradas de VENTA CAMPO
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} = 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} = 1") Then Exit Sub
        
        'sólo entradas que tengan importe (rhisfruta.impentrada)
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.impentrada} <> 0") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.impentrada} <> 0") Then Exit Sub
        
        
        nTabla = "(rhisfruta INNER JOIN rsocios_seccion ON rhisfruta.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie "
        nTabla = "(" & nTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
        
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = vSQL
        
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        If HayRegParaInforme(nTabla, cadSelect) Then
            NRegs = TotalFacturas(nTabla, cadSelect)
            If NRegs <> 0 Then
                'combo1(0).listindex = 0 ---> anticipo venta campo
                '                    = 1 ---> liquidación venta campo
                Select Case Combo1(0).ListIndex
                    Case 0 ' anticipo
                        If Not ComprobarTiposMovimiento(2, nTabla, cadSelect) Then Exit Sub
                    Case 1 ' liquidacion venta campo
                        If Not ComprobarTiposMovimiento(3, nTabla, cadSelect) Then Exit Sub
                End Select
                
                Me.Pb3.visible = True
                Me.Pb3.Max = NRegs
                Me.Pb3.Value = 0
                Me.Refresh
                        
                If FacturacionVentaCampo(Combo1(0).ListIndex, nTabla, cadSelect, txtcodigo(14).Text, Me.Pb3) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                   
                    ' si imprimimos resumen
                    If Me.Check1(0).Value Then
                        cadFormula = ""
                        cadParam = cadParam & "pFecFac= """ & txtcodigo(14).Text & """|"
                        numParam = numParam + 1
                        
                        FecFac = CDate(txtcodigo(14).Text)
                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                
                        cadNombreRPT = "rResumFacturas.rpt"
                        
                        Select Case Combo1(0).ListIndex
                            Case 0 ' anticipos
                                cadTitulo = "Resumen Facturas Anticipos Venta Campo"
                                cadParam = cadParam & "pTitulo= ""Resumen Fact.Anticipos V.Campo""|"
                                numParam = numParam + 1
                            Case 1 ' liquidaciones
                                cadTitulo = "Resumen Facturas Liquidación Venta Campo"
                                cadParam = cadParam & "pTitulo= ""Resumen Fact.Liquidación V.Campo""|"
                                numParam = numParam + 1
                        End Select
                        ConSubInforme = False
                        
                        LlamarImprimir
                    End If
                    
                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE VENTA CAMPO
                    If Me.Check1(1).Value Then
                        cadFormula = ""
                        cadSelect = ""
                        'Tipo de Factura: Anticipo
                        Select Case Combo1(0).ListIndex
                            Case 0 ' anticipos
                                tipoMov = "FAC"
                                cadAux = "({stipom.tipodocu} = 3)"
                                cadTitulo = "Reimpresión Facturas Anticipos V.Campo"
                            Case 1
                                tipoMov = "FLC"
                                cadAux = "({stipom.tipodocu} = 4)"
                                cadTitulo = "Reimpresión Facturas Liquidación V.Campo"
                        End Select
                        
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                        'Nº Factura
                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(Combo1(0).ListIndex) & "])"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                        'Fecha de Factura
                        FecFac = CDate(txtcodigo(14).Text)
                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = "{rfactsoc.fecfactu}= '" & Format(FecFac, FormatoFecha) & "'"
                        
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                       
                        indRPT = 23 'Impresion de facturas de socios
                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        ConSubInforme = True
                        
                        LlamarImprimir
                        
                        If frmVisReport.EstaImpreso Then
                            ActualizarRegistros "rfactsoc", cadSelect
                        End If
                    End If
                                   
                End If
            Else
                MsgBox "No hay entradas a facturar.", vbExclamation
            End If
            
            Me.Pb3.visible = False
            CmdCancelGenFac_Click
        End If
    End If
End Sub

Private Sub CmdAcepModelo_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim I As Byte
Dim nTabla As String

Dim vWhere As String
Dim b As Boolean
Dim tipo As Byte
    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub

    'D/H Socios
    cDesde = Trim(txtcodigo(34).Text)
    cHasta = Trim(txtcodigo(35).Text)
    nDesde = txtNombre(34).Text
    nHasta = txtNombre(35).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(32).Text)
    cHasta = Trim(txtcodigo(33).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
   
    nTabla = "rfactsoc INNER JOIN usuarios.stipom stipom ON rfactsoc.codtipom = stipom.codtipom "
    
    Select Case OpcionListado
        Case 10 'modelo 190
            If Not AnyadirAFormula(cadFormula, "{rfactsoc.impreten} <> 0") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{rfactsoc.impreten} <> 0") Then Exit Sub
        
            If Not AnyadirAFormula(cadFormula, "{stipom.tipodocu} in [1,2,3,4]") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{stipom.tipodocu} in (1,2,3,4)") Then Exit Sub
        
        Case 11 'modelo 346
            ' seleccionamos tipodocu: 5 = subvencion
            '                         6 = siniestro
            If Not AnyadirAFormula(cadFormula, "{stipom.tipodocu} in [5,6]") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{stipom.tipodocu} in (5,6)") Then Exit Sub
    
            If Not AnyadirAFormula(cadFormula, "{rfactsoc_variedad.imporvar} <> 0") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{rfactsoc_variedad.imporvar} <> 0") Then Exit Sub
            
            nTabla = "(" & nTabla & ") INNER JOIN rfactsoc_variedad ON rfactsoc.codtipom = rfactsoc_variedad.codtipom "
            nTabla = nTabla & " and rfactsoc.numfactu = rfactsoc_variedad.numfactu "
            nTabla = nTabla & " and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
            
    End Select
    

    If HayRegParaInforme(nTabla, cadSelect) Then
        b = GeneraFicheroModelo(OpcionListado - 10, nTabla, cadSelect)
        If b Then
            If CopiarFichero Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                CmdCancelModelo_Click
            End If
        End If
    End If
End Sub

Private Sub CmdAcepResul_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim nTabla As String

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
    
    'Tipo de movimiento:
    Tipos = ""
    For I = 1 To ListView1(1).ListItems.Count
        If ListView1(1).ListItems(I).Checked Then
            Tipos = Tipos & DBSet(ListView1(1).ListItems(I).Key, "T") & ","
        End If
    Next I
    
    If Tipos = "" Then
        MsgBox "Debe seleccionar al menos un tipo de factura.", vbExclamation
        Exit Sub
    Else
        ' quitamos la ultima coma
        Tipos = "{rfactsoc.codtipom} in (" & Mid(Tipos, 1, Len(Tipos) - 1) & ")"
        If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
        Tipos = Replace(Replace(Tipos, "(", "["), ")", "]")
        If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
    End If
    
    'D/H Socios
    cDesde = Trim(txtcodigo(24).Text)
    cHasta = Trim(txtcodigo(25).Text)
    nDesde = txtNombre(24).Text
    nHasta = txtNombre(25).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Clase
    cDesde = Trim(txtcodigo(28).Text)
    cHasta = Trim(txtcodigo(29).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(26).Text)
    cHasta = Trim(txtcodigo(27).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
        
    nTabla = "(rfactsoc INNER JOIN rfactsoc_variedad ON rfactsoc.codtipom = rfactsoc_variedad.codtipom "
    nTabla = nTabla & " and rfactsoc.numfactu = rfactsoc_variedad.numfactu and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu) "
    nTabla = nTabla & " INNER JOIN variedades ON rfactsoc_variedad.codvarie = variedades.codvarie "
    
    Set frmMens = New frmMensajes
    
    frmMens.OpcionMensaje = 16
    frmMens.Show vbModal
    
    Set frmMens = Nothing
    
    
    
    If HayRegistros(nTabla, cadSelect) Then
        cadParam = cadParam & "pResumen=" & Me.Check1(4).Value & "|"
        numParam = numParam + 1
        'Nombre fichero .rpt a Imprimir
        Select Case OpcionListado
            Case 8
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = "rInfResultados.rpt"
                cadTitulo = "Informe de Resultados"
            Case 9
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = "rInfRetenciones.rpt"
                cadTitulo = "Informe de Retenciones"
        End Select
          
        ConSubInforme = False
        
        LlamarImprimir
    End If

End Sub

Private Sub cmdAceptarAnt_Click()
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

Dim NRegs As Long
Dim FecFac As Date
Dim TipoPrec As Byte ' 0 anticipos
                     ' 1 liquidaciones
Dim b As Boolean
Dim Sql2 As String


    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
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
        
        Sql2 = ""
        If txtcodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(20).Text, "N")
        If txtcodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(21).Text, "N")
        
        
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
            
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.SeccionHorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.SeccionHorto) Then Exit Sub
        
        'Socio que no sea tercero
        If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        
        
        'sólo entradas distintas de VENTA CAMPO
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1") Then Exit Sub
        
        
        '++monica: 28/07/2009 dependiendo del tipo de recoleccion (0=coop 1=socio 2=todos)
        Select Case Combo1(2).ListIndex
            Case 0      ' recolectado cooperativa
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 0") Then Exit Sub
            Case 1      ' recolectado socio
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 1") Then Exit Sub
            Case 2      ' ambos
            
        End Select
        
        
        nTabla = "(((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN rhisfruta_clasif ON rhisfruta.numalbar = rhisfruta_clasif.numalbar) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie "
        
        Select Case OpcionListado
            Case 1 ' Listado de anticipos
                'Nombre fichero .rpt a Imprimir
                indRPT = 24 ' informe de anticipos
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatAnticipos.rpt"
                cadTitulo = "Informe de Anticipos"
            Case 2 ' Prevision de pago de anticipos
                cadNombreRPT = "rPrevPagosAnt.rpt"
                cadTitulo = "Previsión de Pago de Anticipos"
            
            Case 3 ' Facturación de Anticipos
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Anticipos"
            
            Case 12 ' Listado de Liquidacion
                'Nombre fichero .rpt a Imprimir
                indRPT = 26 ' informe de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatLiquidacion.rpt"
                cadTitulo = "Informe de Liquidación"
                
            Case 13 ' Prevision de pago de liquidacion
                cadNombreRPT = "rPrevPagosLiq.rpt"
                cadTitulo = "Previsión de Pago de Liquidación"
            
            Case 14 ' Facturación de Liquidacion
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación"
                
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
        
            Select Case OpcionListado
                Case 1, 2, 3
                    TipoPrec = 0 ' ANTICIPOS
                Case 12, 13, 14
                    TipoPrec = 1 ' LIQUIDACIONES
            End Select
            
            If HayPreciosVariedades(TipoPrec, nTabla, cadSelect, Combo1(2).ListIndex) Then
                'D/H fecha
                cDesde = Trim(txtcodigo(6).Text)
                cHasta = Trim(txtcodigo(7).Text)
                cadDesde = CDate(cDesde)
                cadhasta = CDate(cHasta)
                cadAux = "{rprecios.fechaini}= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechaini}=" & DBSet(txtcodigo(6).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                cadAux = "{rprecios.fechafin}= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechafin}=" & DBSet(txtcodigo(7).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                ' si se trata de anticipos--> seleccionamos los precios de anticipos
                ' sino los de liquidaciones
                If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = " & TipoPrec) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = " & TipoPrec) Then Exit Sub
                
                Select Case OpcionListado
                    Case 1, 12 '1 - informe de anticipos
                               '12- informe de liquidaciones
                        'pasamos como parametro la fecha de anticipo
                        cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
                        numParam = numParam + 1
                        ConSubInforme = False
                        
                        LlamarImprimir
                    
                    Case 2  '2 - listado de prevision de pagos de anticipos
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        If CargarTemporalAnticipos(nTabla, cadSelect) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                            ConSubInforme = False
                            
                            LlamarImprimir
                        End If
                        
                    Case 13 '13- listado de prevision de pagos de liquidaciones
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        If CargarTemporalLiquidacion(nTabla, cadSelect) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                            ConSubInforme = False
                            
                            LlamarImprimir
                        End If
                        
                    Case 3, 14 '3 .- factura de anticipos
                               '14.- factura de liquidaciones
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        NRegs = TotalFacturas(nTabla, cadSelect)
                        If NRegs <> 0 Then
                            If Not ComprobarTiposMovimiento(TipoPrec, nTabla, cadSelect) Then
                                Exit Sub
                            End If
                            
                            Me.Pb1.visible = True
                            Me.Pb1.Max = NRegs
                            Me.Pb1.Value = 0
                            Me.Refresh
                            b = False
                            If TipoPrec = 0 Then
                                b = FacturacionAnticipos(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1)
                            Else
                                b = FacturacionLiquidaciones(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1)
                            End If
                            
                            If b Then
                                MsgBox "Proceso realizado correctamente.", vbExclamation
                                               
                                'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                                If Me.Check1(2).Value Then
                                    cadFormula = ""
                                    cadParam = cadParam & "pFecFac= """ & txtcodigo(15).Text & """|"
                                    numParam = numParam + 1
                                    If TipoPrec = 0 Then
                                        cadParam = cadParam & "pTitulo= ""Resumen Facturación de Anticipos""|"
                                    Else
                                        cadParam = cadParam & "pTitulo= ""Resumen Facturación de Liquidaciones""|"
                                    End If
                                    numParam = numParam + 1
                                    
                                    FecFac = CDate(txtcodigo(15).Text)
                                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                    ConSubInforme = False
                                    
                                    LlamarImprimir
                                End If
                                'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS/LIQUIDACION
                                If Me.Check1(3).Value Then
                                    cadFormula = ""
                                    cadSelect = ""
                                    If TipoPrec = 0 Then 'Tipo de Factura: Anticipo
                                        cadAux = "({stipom.tipodocu} = 1)"
                                    Else  'Tipo de Factura: Liquidación
                                        cadAux = "({stipom.tipodocu} = 2)"
                                    End If
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                    'Nº Factura
                                    If TipoPrec = 0 Then
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(2) & "])"
                                    Else
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(3) & "])"
                                    End If
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                     
                                    'Fecha de Factura
                                    FecFac = CDate(txtcodigo(15).Text)
                                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                                   
                                    indRPT = 23 'Impresion de facturas de socios
                                    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                                    'Nombre fichero .rpt a Imprimir
                                    cadNombreRPT = nomDocu
                                    'Nombre fichero .rpt a Imprimir
                                    If TipoPrec = 0 Then
                                        cadTitulo = "Reimpresión de Facturas Anticipos"
                                    Else
                                        cadTitulo = "Reimpresión de Facturas Liquidaciones"
                                    End If
                                    ConSubInforme = True
                                    
                                    LlamarImprimir
                                    
                                    If frmVisReport.EstaImpreso Then
                                        ActualizarRegistrosFac "rfactsoc", cadSelect
                                    End If
                                End If
                                'SALIR DE LA FACTURACION DE ANTICIPOS / LIQUIDACIONES
                                cmdCancelAnt_Click
                            End If
                        Else
                            MsgBox "No hay entradas a facturar.", vbExclamation
                        End If
                End Select
'            '++monica:27/07/2009
'            Else
'                MsgBox "No hay precios para las calidades en este rango. Revise.", vbExclamation
            End If
        End If
    End If
End Sub

Private Sub cmdAceptarAntGastos_Click()
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

Dim NRegs As Long
Dim FecFac As Date
Dim TipoPrec As Byte ' 0 anticipos
                     ' 1 liquidaciones
Dim b As Boolean
Dim Sql2 As String

Dim cadSelect1 As String



    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
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
        
        Sql2 = ""
        If txtcodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(20).Text, "N")
        If txtcodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(21).Text, "N")
        
        
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
            
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.SeccionHorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.SeccionHorto) Then Exit Sub
        
        'Socio que no sea tercero
        If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        
        
        'sólo entradas distintas de VENTA CAMPO
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1") Then Exit Sub
        
        
        'sólo entradas recolectadas por socio
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 1") Then Exit Sub
        
        
        
        nTabla = "((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN rhisfruta_clasif ON rhisfruta.numalbar = rhisfruta_clasif.numalbar) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
        
        Select Case OpcionListado
            Case 2 ' Prevision de pago de anticipos gastos recoleccion
                cadNombreRPT = "rPrevPagosAntGastos.rpt"
                cadTitulo = "Previsión Pago de Anticipos Gastos"
            
            Case 3 ' Facturación de Anticipos de Gastos de recoleccion
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Anticipos Gastos"
            
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
        
            Select Case OpcionListado
                Case 2  '2 - listado de prevision de pagos de anticipos
                    cadSelect1 = " rhisfruta.tipoentr <> 1 and rhisfruta.recolect = 1 "
                    If txtcodigo(6).Text <> "" Then cadSelect1 = cadSelect1 & " and rhisfruta.fecalbar >=" & DBSet(txtcodigo(6).Text, "F")
                    If txtcodigo(7).Text <> "" Then cadSelect1 = cadSelect1 & " and rhisfruta.fecalbar <=" & DBSet(txtcodigo(7).Text, "F")
                    
                    If CargarTemporalAnticiposGastos(nTabla, cadSelect, cadSelect1) Then
                        cadFormula = ""
                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                        ConSubInforme = False
                        
                        LlamarImprimir
                    End If
                    
                Case 3  '3 .- factura de anticipos de gastos
                    TipoPrec = 0 ' son anticipos
                    
                    NRegs = TotalFacturas(nTabla, cadSelect)
                    If NRegs <> 0 Then
                        If Not ComprobarTiposMovimiento(TipoPrec, nTabla, cadSelect) Then
                            Exit Sub
                        End If
                        
                        Me.Pb1.visible = True
                        Me.Pb1.Max = NRegs
                        Me.Pb1.Value = 0
                        Me.Refresh
                        
                        cadSelect1 = " rhisfruta.tipoentr <> 1 and rhisfruta.recolect = 1 "
                        If txtcodigo(6).Text <> "" Then cadSelect1 = cadSelect1 & " and rhisfruta.fecalbar >=" & DBSet(txtcodigo(6).Text, "F")
                        If txtcodigo(7).Text <> "" Then cadSelect1 = cadSelect1 & " and rhisfruta.fecalbar <=" & DBSet(txtcodigo(7).Text, "F")
                        
                        
                        b = FacturacionAnticiposGastos(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1, cadSelect1)
                        
                        If b Then
                            MsgBox "Proceso realizado correctamente.", vbExclamation
                                           
                            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS GASTOS
                            If Me.Check1(2).Value Then
                                cadFormula = ""
                                cadParam = cadParam & "pFecFac= """ & txtcodigo(15).Text & """|"
                                numParam = numParam + 1
                                If TipoPrec = 0 Then
                                    cadParam = cadParam & "pTitulo= ""Resumen Facturación de Anticipos Gastos""|"
                                End If
                                numParam = numParam + 1
                                
                                FecFac = CDate(txtcodigo(15).Text)
                                cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                ConSubInforme = False
                                
                                LlamarImprimir
                            End If
                            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS GASTOS
                            If Me.Check1(3).Value Then
                                cadFormula = ""
                                cadSelect = ""
                                If TipoPrec = 0 Then 'Tipo de Factura: Anticipo
                                    cadAux = "({stipom.tipodocu} = 1)"
                                End If
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                'Nº Factura
                                If TipoPrec = 0 Then
                                    cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(2) & "])"
                                Else
                                    cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(3) & "])"
                                End If
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                 
                                'Fecha de Factura
                                FecFac = CDate(txtcodigo(15).Text)
                                cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                               
                                indRPT = 23 'Impresion de facturas de socios
                                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                                'Nombre fichero .rpt a Imprimir
                                cadNombreRPT = nomDocu
                                'Nombre fichero .rpt a Imprimir
                                If TipoPrec = 0 Then
                                    cadTitulo = "Reimpresión de Facturas Anticipos"
                                End If
                                ConSubInforme = True
                                
                                LlamarImprimir
                                
                                If frmVisReport.EstaImpreso Then
                                    ActualizarRegistrosFac "rfactsoc", cadSelect
                                End If
                            End If
                            'SALIR DE LA FACTURACION DE ANTICIPOS / LIQUIDACIONES
                            cmdCancelAnt_Click
                        End If
                    Else
                        MsgBox "No hay Gastos Recolección a facturar.", vbExclamation
                    End If
            End Select
        End If
    End If
End Sub


Private Sub cmdAceptarReimp_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String

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
    
    'Tipo de movimiento:
    Tipos = ""
    For I = 1 To ListView1(0).ListItems.Count
        If ListView1(0).ListItems(I).Checked Then
            Tipos = Tipos & DBSet(ListView1(0).ListItems(I).Key, "T") & ","
        End If
    Next I
    
    If Tipos = "" Then
        MsgBox "Debe seleccionar al menos un tipo de factura.", vbExclamation
        Exit Sub
    Else
        ' quitamos la ultima coma
        Tipos = "{rfactsoc.codtipom} in (" & Mid(Tipos, 1, Len(Tipos) - 1) & ")"
        If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
        Tipos = Replace(Replace(Tipos, "(", "["), ")", "]")
        If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
    End If
    
    'D/H Cliente
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Nro de Factura
    cDesde = Trim(txtcodigo(4).Text)
    cHasta = Trim(txtcodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    
    
    If HayRegistros(Tabla, cadSelect) Then
        indRPT = 23 'Impresion de facturas de anticipos
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
          
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
          
          
        'Nombre fichero .rpt a Imprimir
        cadTitulo = "Reimpresión de Facturas Socios"
        ConSubInforme = True
        
        LlamarImprimir
        
        If frmVisReport.EstaImpreso Then
            ActualizarRegistros "rfactsoc", cadSelect
        End If
    End If


End Sub

Private Sub cmdCancelAnt_Click()
    Unload Me
End Sub

Private Sub CmdCancelGenFac_Click()
    Unload Me
End Sub

Private Sub CmdCancelModelo_Click()
    Unload Me
End Sub

Private Sub cmdCancelReimp_Click()
    Unload Me
End Sub

Private Sub cmdCancelDesF_Click()
    Unload Me
End Sub

Private Sub CmdCancelResul_Click()
    Unload Me
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Index = 1 Then
        Select Case Combo1(Index).ListIndex
            Case 0 ' anticipo venta campo
                ' si solo hay un tipo de movimiento de anticipo venta campo
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(3) = 1 Then
                    txtcodigo(9).Text = vParamAplic.PrimFactAntVC
                    txtcodigo(10).Text = vParamAplic.UltFactAntVC
                End If
            Case 1 ' liquidacion venta campo
                ' si solo hay un tipo de movimiento de liquidacion venta campo
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(4) = 1 Then
                    txtcodigo(9).Text = vParamAplic.PrimFactLiqVC
                    txtcodigo(10).Text = vParamAplic.UltFactLiqVC
                End If
        End Select
    End If
End Sub

Private Sub combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 1, 2, 3 ' 1-Inf.Anticipos
                         ' 2-Listado de Previsión de pago
                         ' 3-Facturas de Anticipos
                PonerFoco txtcodigo(12)
                
            Case 4    ' reimpresion de facturas de SOCIOS
                PonerFoco txtcodigo(4)
                
            Case 5    ' deshacer proceso de facturacion de anticipos
                PonerFoco txtcodigo(8)
                Me.Pb2.visible = False
                ' si solo hay un tipo de movimiento de anticipo
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(1) = 1 Then
                    txtcodigo(9).Text = vParamAplic.PrimFactAnt
                    txtcodigo(10).Text = vParamAplic.UltFactAnt
                End If
                
            Case 6    ' generacion de factura de venta campo (anticipo o liquidacion)
                Combo1(0).ListIndex = 0 ' por defecto anticipo
                Pb3.visible = False
                txtcodigo(14).Text = Format(Now, "dd/mm/yyyy")
                Check1(0).Value = 1
                Check1(1).Value = 1
                PonerFocoCmb Combo1(0)
                
            Case 7    ' deshacer proceso de facturacion de venta campo
                Me.Pb2.visible = False
                Combo1(1).ListIndex = 0 ' por defecto anticipo
'                txtCodigo(9).Text = vParamAplic.PrimFactAntVC
'                txtCodigo(10).Text = vParamAplic.UltFactAntVC
                PonerFoco txtcodigo(8)
            
            Case 8, 9   ' 8 - informe de resultados
                        ' 9 - informe de retenciones
                PonerFoco txtcodigo(24)
                
            Case 10, 11  ' 10 - grabacion modelo 190
                         ' 11 - grabacion modelo 346
                PonerFoco txtcodigo(34)
                Me.FrameDomicilio.visible = (OpcionListado = 10)
                Me.FrameDomicilio.Enabled = (OpcionListado = 10)
                Me.BarraEst.Enabled = (OpcionListado = 10)
                Me.BarraEst.visible = (OpcionListado = 10)
                txtcodigo(30).Text = Format(Year(Now), "0000")
                txtcodigo(36).Text = vParam.PerContacto
                txtcodigo(37).Text = vParam.Telefono
            
            Case 12, 13, 14 ' 12-Inf.Liquidacion
                            ' 13-Listado de Previsión de pago
                            ' 14-Facturas de Liquidacion
                PonerFoco txtcodigo(12)
            
            Case 15    ' deshacer proceso de facturacion de liquidacion
                PonerFoco txtcodigo(8)
                Me.Pb2.visible = False
                ' si solo hay un tipo de movimiento de liquidacion
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(1) = 1 Then
                    txtcodigo(9).Text = vParamAplic.PrimFactLiq
                    txtcodigo(10).Text = vParamAplic.UltFactLiq
                End If
                
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    Set List = New Collection
    For h = 24 To 27
        List.Add h
    Next h
    For h = 1 To 10
        List.Add h
    Next h
    List.Add 12
    List.Add 13
    List.Add 14
    List.Add 15
    List.Add 18
    List.Add 19
    
    For h = 0 To 1
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 12 To 13
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 16 To 21
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 24 To 25
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 28 To 29
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 34 To 35
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    FrameAnticipos.visible = False
    FrameReimpresion.visible = False
    FrameDesFacturacion.visible = False
    FrameGeneraFactura.visible = False
    FrameResultados.visible = False
    FrameGrabacionModelos.visible = False
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 1, 12   '1- Informe de Anticipos
                 '12- Informe de Liquidacion
        FrameAnticiposVisible True, h, w
        Tabla = "rhisfruta"
        Me.FrameFechaAnt.visible = True
        Me.FrameFechaAnt.Enabled = True
        If OpcionListado = 1 Then
            Me.Label3.Caption = "Informe de Anticipos"
        Else
            Me.Label3.Caption = "Informe de Liquidación"
        End If
        Me.Pb1.visible = False
        Me.FrameOpciones.visible = False
        Me.FrameOpciones.Enabled = False
        
        CargaCombo
        Combo1(2).ListIndex = 2
    Case 2, 13   '2 - Listado de prevision de pagos de anticipos
                 '13- Listado de prevision de pagos de liquidacion
        FrameAnticiposVisible True, h, w
        Tabla = "rhisfruta"
        Me.FrameFechaAnt.visible = False
        Me.FrameFechaAnt.Enabled = False
        If OpcionListado = 2 Then
            Me.Label3.Caption = "Previsión de Pagos Anticipos"
            If AnticipoGastos Then
                Me.Label3.Caption = "Previsión Pagos Anticipos Gastos"
            End If
        Else
            Me.Label3.Caption = "Previsión de Pagos Liquidación"
        End If
        
        Me.Pb1.visible = False
        Me.FrameOpciones.visible = False
        Me.FrameOpciones.Enabled = False
        
        CargaCombo
        Combo1(2).ListIndex = 2
    Case 3, 14   '3 - Factura de Anticipos
                 '14- Factura de Liquidacion
        FrameAnticiposVisible True, h, w
        Tabla = "rhisfruta"
        Me.FrameFechaAnt.visible = True
        Me.FrameFechaAnt.Enabled = True
        Me.Caption = "Facturación"
        If OpcionListado = 3 Then
            Me.Label3.Caption = "Factura de Anticipos"
            If AnticipoGastos Then
                Me.Label3.Caption = "Factura de Anticipos Gastos"
            End If
        Else
            Me.Label3.Caption = "Factura de Liquidación"
        End If
        Me.Pb1.visible = False
        Me.FrameOpciones.visible = True
        Me.FrameOpciones.Enabled = True
        Me.Check1(2).Value = 1
        Me.Check1(3).Value = 1
        
        CargaCombo
        Combo1(2).ListIndex = 2
    Case 4   ' Reimpresion de facturas de SOCIOS
        FrameReimpresionVisible True, h, w
        Tabla = "rfactsoc"
        Me.FrameFechaAnt.visible = True
        Me.FrameFechaAnt.Enabled = True
        Me.Label3.Caption = "Factura de Socios"
        CargarListView (0)
        
    Case 5   ' Deshacer Proceso de facturación de Anticipos
        ActivarCLAVE
        FrameTipoFactura.visible = False
        FrameDesFacturacionVisible True, h, w
        Tabla = "rfactsoc"
        Me.Caption = "Deshacer Proceso Facturación de Anticipos"
        
    Case 6   ' Generacion de factura de venta campo (anticipo o liquidacion)
        FrameGeneraFacturaVisible True, h, w
        CargaCombo
        Tabla = "rhisfruta"
        Me.Caption = "Facturación"
    
    Case 7   ' Deshacer Proceso de facturación de venta campo
        ActivarCLAVE
        FrameTipoFactura.visible = True
        CargaCombo
        FrameDesFacturacionVisible True, h, w
        Tabla = "rfactsoc"
        Me.Caption = "Deshacer Proceso Facturación Venta Campo"
                
    Case 8, 9   '8= Informe de Resultados de facturas de SOCIOS
                '9= Informe de Retenciones de facturas de SOCIOS
        If OpcionListado = 8 Then
            Label8.Caption = "Listado de Resultados"
        Else
            Label8.Caption = "Listado de Retenciones"
        End If
        FrameResultadosVisible True, h, w
        Tabla = "rfactsoc"
        CargarListView (1)
        
    Case 10, 11 '10 = grabacion modelo 190
                '11 = grabacion modelo 346
        If OpcionListado = 10 Then
            Label9.Caption = "Grabación Modelo 190"
        Else
            Label9.Caption = "Grabación Modelo 346"
        End If
        FrameGrabacionModelosVisible True, h, w
        Tabla = "rfactsoc"
    
    Case 15   ' Deshacer Proceso de facturación de Liquidacion
        ActivarCLAVE
        FrameTipoFactura.visible = False
        FrameDesFacturacionVisible True, h, w
        Tabla = "rfactsoc"
        Me.Caption = "Deshacer Proceso Facturación de Liquidación"
        
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case OpcionListado
        Case 3
            DesBloqueoManual ("FACANT")
        Case 14
            DesBloqueoManual ("FACLIQ")
    End Select
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de calidades
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub Image1_Click(Index As Integer)
Dim I As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        ' reimpresion de facturas socios
        Case 0
            For I = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(I).Checked = True
            Next I
        Case 1
            For I = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(I).Checked = False
            Next I
        ' informe de resultados y listado de retenciones
        Case 2
            For I = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(I).Checked = True
            Next I
        Case 3
            For I = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(I).Checked = False
            Next I
    End Select
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 18, 19, 20, 21, 28, 29 'Clases
            AbrirFrmClase (Index)
        
        Case 0, 1, 12, 13, 16, 17, 24, 25 'SOCIOS
            AbrirFrmSocios (Index)
        
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
        Case 0
            indice = 6
        Case 1
            indice = 7
        Case 2
            indice = 15
        Case 3, 4
            indice = Index - 1
        Case 5
            indice = 14
        Case 6
            indice = 11
        Case 7, 8
            indice = Index + 15
        Case 9, 10
            indice = Index + 17
        Case 11, 12
            indice = Index + 21
    End Select

    imgFec(0).Tag = indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(indice).Text <> "" Then frmC.NovaData = txtcodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(CByte(imgFec(0).Tag)) '<===
    ' ********************************************

End Sub



Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
    If OpcionListado = 10 Then
        If Index = 40 Then
            BarraEst.SimpleText = " CL = Calle    AV = Avenida."
        Else
            BarraEst.SimpleText = ""
        End If
        BarraEst.visible = (BarraEst.SimpleText <> "")
    End If
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
            Case 0: KEYBusqueda KeyAscii, 0 'socio desde
            Case 1: KEYBusqueda KeyAscii, 1 'socio hasta
            Case 12: KEYBusqueda KeyAscii, 12 'socio desde
            Case 13: KEYBusqueda KeyAscii, 13 'socio hasta
            Case 16: KEYBusqueda KeyAscii, 16 'socio desde
            Case 17: KEYBusqueda KeyAscii, 17 'socio hasta
            Case 24: KEYBusqueda KeyAscii, 24 'socio desde
            Case 25: KEYBusqueda KeyAscii, 25 'socio hasta
            Case 34: KEYBusqueda KeyAscii, 34 'socio desde
            Case 35: KEYBusqueda KeyAscii, 35 'socio hasta
            Case 18: KEYBusqueda KeyAscii, 18 'clase desde
            Case 19: KEYBusqueda KeyAscii, 19 'clase hasta
            Case 20: KEYBusqueda KeyAscii, 20 'clase desde
            Case 21: KEYBusqueda KeyAscii, 21 'clase hasta
            Case 28: KEYBusqueda KeyAscii, 28 'clase desde
            Case 29: KEYBusqueda KeyAscii, 29 'clase hasta
            Case 26: KEYFecha KeyAscii, 9 'fecha desde
            Case 27: KEYFecha KeyAscii, 10 'fecha hasta
            Case 6: KEYFecha KeyAscii, 0 'fecha desde
            Case 7: KEYFecha KeyAscii, 1 'fecha hasta
            Case 22: KEYFecha KeyAscii, 7 'fecha desde
            Case 23: KEYFecha KeyAscii, 8 'fecha hasta
            Case 2: KEYFecha KeyAscii, 3 'fecha desde
            Case 3: KEYFecha KeyAscii, 4 'fecha hasta
            Case 32: KEYFecha KeyAscii, 11 'fecha desde
            Case 33: KEYFecha KeyAscii, 12 'fecha hasta
            
            Case 11: KEYFecha KeyAscii, 6 'fecha
            Case 14: KEYFecha KeyAscii, 5 'fecha
            Case 15: KEYFecha KeyAscii, 2 'fecha
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
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
    
        Case 0, 1, 12, 13, 16, 17, 24, 25, 34, 35 'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            
        Case 4, 5 ' NROS DE FACTURA
            PonerFormatoEntero txtcodigo(Index)
            
        Case 2, 3, 6, 7, 11, 15, 26, 27, 32, 33 'FECHAS
            b = True
            If txtcodigo(Index).Text <> "" Then b = PonerFormatoFecha(txtcodigo(Index))
            If b And Index = 7 And (Me.OpcionListado = 1 Or Me.OpcionListado = 3 Or Me.OpcionListado = 12 Or Me.OpcionListado = 14) Then PonerFoco txtcodigo(15)
            If b And Index = 15 And (Me.OpcionListado = 1 Or Me.OpcionListado = 3 Or Me.OpcionListado = 12 Or Me.OpcionListado = 14) Then
                If Not AnticipoGastos Then
                    cmdAceptarAnt.SetFocus
                Else
                    cmdAceptarAntGastos.SetFocus
                End If
            End If
            
        Case 14, 22, 23  ' FECHA DE GENERACION DE FACTURA
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
        Case 8 ' password de deshacer facturacion
            If txtcodigo(Index).Text = "" Then Exit Sub
            If Trim(txtcodigo(Index).Text) <> Trim(txtcodigo(Index).Tag) Then
                MsgBox "    ACCESO DENEGADO    ", vbExclamation
                txtcodigo(Index).Text = ""
                PonerFoco txtcodigo(Index)
            Else
                DesactivarCLAVE
                Select Case OpcionListado
                    Case 5, 15 '5 = anticipos
                               '15= liquidaciones
                        PonerFoco txtcodigo(9)
                    Case 7 ' venta campo
                        PonerFocoCmb Combo1(1)
                End Select
            End If
        
        Case 9, 10 ' numero de facturas
            If txtcodigo(Index).Text <> "" Then PonerFormatoEntero txtcodigo(Index)
        
        Case 30, 31, 37, 39 ' datos de modelo190 y modelo346
            If txtcodigo(Index).Text <> "" Then PonerFormatoEntero txtcodigo(Index)
            
        Case 18, 19, 20, 21, 28, 29 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
            
    End Select
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
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
          
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
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

Private Sub FrameAnticiposVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
Dim b As Boolean

'Frame para el listado de socios por seccion
    Me.FrameAnticipos.visible = visible
    If visible = True Then
        Me.FrameAnticipos.Top = -90
        Me.FrameAnticipos.Left = 0
        Me.FrameAnticipos.Height = 5640
        Me.FrameAnticipos.Width = 6615
        w = Me.FrameAnticipos.Width
        h = Me.FrameAnticipos.Height
        
        b = (OpcionListado = 1 Or OpcionListado = 2 Or OpcionListado = 3 Or _
             OpcionListado = 12 Or OpcionListado = 13 Or OpcionListado = 14) And _
             Not AnticipoGastos
             
        
        FrameRecolectado.Enabled = b
        FrameRecolectado.visible = b
    

    
    
        If AnticipoGastos Then
            ' desactivo los botones de anticipos normales
            Me.cmdAceptarAnt.visible = False
            Me.cmdAceptarAnt.Enabled = False
            ' activo los botones de anticipos de gastos
            Me.cmdAceptarAntGastos.visible = True
            Me.cmdAceptarAntGastos.Enabled = True
            ' los situo
            Me.cmdAceptarAntGastos.Left = 4110
            Me.cmdAceptarAntGastos.Caption = "&Aceptar"
        End If
    
    End If
End Sub


Private Sub FrameReimpresionVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameReimpresion.visible = visible
    If visible = True Then
        Me.FrameReimpresion.Top = -90
        Me.FrameReimpresion.Left = 0
        Me.FrameReimpresion.Height = 5640
        Me.FrameReimpresion.Width = 6675
        w = Me.FrameReimpresion.Width
        h = Me.FrameReimpresion.Height
    End If
End Sub


Private Sub FrameResultadosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameResultados.visible = visible
    If visible = True Then
        Me.FrameResultados.Top = -90
        Me.FrameResultados.Left = 0
        Me.FrameResultados.Height = 5490
        Me.FrameResultados.Width = 6675
        w = Me.FrameResultados.Width
        h = Me.FrameResultados.Height
    End If
End Sub

Private Sub FrameGrabacionModelosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameGrabacionModelos.visible = visible
    If visible = True Then
        Me.FrameGrabacionModelos.Top = -90
        Me.FrameGrabacionModelos.Left = 0
        Select Case OpcionListado
            Case 10
                Me.FrameGrabacionModelos.Height = 6480
                Me.CmdAcepModelo.Top = 5640
                Me.CmdCancelModelo.Top = 5640
            Case 11
                Me.FrameGrabacionModelos.Height = 5490
                Me.CmdAcepModelo.Top = 4740
                Me.CmdCancelModelo.Top = 4740
        End Select
        Me.FrameGrabacionModelos.Width = 6675
        w = Me.FrameGrabacionModelos.Width
        h = Me.FrameGrabacionModelos.Height
    End If
End Sub


Private Sub FrameDesFacturacionVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameDesFacturacion.visible = visible
    If visible = True Then
        Me.FrameDesFacturacion.Top = -90
        Me.FrameDesFacturacion.Left = 0
        Me.FrameDesFacturacion.Height = 4740
        Me.FrameDesFacturacion.Width = 6615
        w = Me.FrameDesFacturacion.Width
        h = Me.FrameDesFacturacion.Height
    End If
End Sub

Private Sub FrameGeneraFacturaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameGeneraFactura.visible = visible
    If visible = True Then
        Me.FrameGeneraFactura.Top = -90
        Me.FrameGeneraFactura.Left = 0
        Me.FrameGeneraFactura.Height = 5790
        Me.FrameGeneraFactura.Width = 6615
        w = Me.FrameGeneraFactura.Width
        h = Me.FrameGeneraFactura.Height
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

Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim Campo As String
Dim nomCampo As String

    Campo = "pGroup" & numGrupo & "="
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
            cadParam = cadParam & Campo & "{" & Tabla & ".codclase}" & "|"
            cadParam = cadParam & nomCampo & " {" & "clases" & ".nomclase}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Producto""" & "|"
            numParam = numParam + 3
            
        Case "Producto"
            cadParam = cadParam & Campo & "{" & Tabla & ".codprodu}" & "|"
            cadParam = cadParam & nomCampo & " {" & "productos" & ".nomprodu}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3

        'Informe de calibres
        Case "Seccion"
            cadParam = cadParam & Campo & "{" & Tabla & ".codsecci}" & "|"
            cadParam = cadParam & nomCampo & "{rseccion.nomsecci}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Seccion""" & "|"
            numParam = numParam + 3
            
        Case "Socio"
            cadParam = cadParam & Campo & "{" & Tabla & ".codsocio}" & "|"
            cadParam = cadParam & nomCampo & " {" & "rsocios" & ".nomsocio}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Socio""" & "|"
            numParam = numParam + 3
            
        'Informe de calidades
        Case "Variedad"
            cadParam = cadParam & Campo & "{" & Tabla & ".codvarie}" & "|"
            cadParam = cadParam & nomCampo & "{variedades.nomvarie}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Variedad""" & "|"
            numParam = numParam + 3
            
        Case "Calidad"
            cadParam = cadParam & Campo & "{" & Tabla & ".codcalid}" & "|"
            cadParam = cadParam & nomCampo & " {" & "rcalidad" & ".nomcalid}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Calidad""" & "|"
            numParam = numParam + 3
            
            
        'Informe de campos
        Case "Socios"
            cadParam = cadParam & Campo & "{rcampos.codsocio}" & "|"
            cadParam = cadParam & nomCampo & "{rsocios.nomsocio}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Socio""" & "|"
            numParam = numParam + 3
            
        Case "Clases"
            cadParam = cadParam & Campo & "{variedades.codclase}" & "|"
            cadParam = cadParam & nomCampo & " {clases.nomclase}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3
            
        Case "Terminos"
            cadParam = cadParam & Campo & "{rpartida.codpobla}" & "|"
            cadParam = cadParam & nomCampo & " {" & "rpueblos" & ".despobla}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Termino Municipal""" & "|"
            numParam = numParam + 3
            
        Case "Zonas"
            cadParam = cadParam & Campo & "{rpartida.codzonas}" & "|"
            cadParam = cadParam & nomCampo & " {" & "rzonas" & ".nomzonas}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Zonas""" & "|"
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
Dim Campo As String
Dim nomCampo As String

    PonerOrden = 0

    Select Case cadgrupo
        Case "Codigo"
            cadParam = cadParam & "Orden" & "= {" & Tabla
            Select Case OpcionListado
                Case 10
                    cadParam = cadParam & ".codclien}|"
                Case 11
                    cadParam = cadParam & ".codprove}|"
            End Select
            tipo = "Código"
        Case "Alfabético"
            cadParam = cadParam & "Orden" & "= {" & Tabla
            Select Case OpcionListado
                Case 10
                    cadParam = cadParam & ".nomclien}|"
                Case 11
                    cadParam = cadParam & ".nomprove}|"
            End Select
            tipo = "Alfabético"
    End Select
    
    numParam = numParam + 1

End Function

Private Sub AbrirFrmCalidad(indice As Integer)
    indCodigo = indice
    Set frmCal = New frmManCalidades
    frmCal.DatosADevolverBusqueda = "2|3|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmCal.Show vbModal
    Set frmCal = Nothing
End Sub

Private Sub AbrirFrmSeccion(indice As Integer)
    indCodigo = indice
    Set frmSec = New frmManSeccion
    frmSec.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmSec.Show vbModal
    Set frmSec = Nothing
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
    
    AyudaClasesCom frmCla, txtcodigo(indice).Text
    
    Set frmCla = Nothing
End Sub



Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
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
Dim vClien As CSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String

    b = True
    Select Case OpcionListado
        Case 1, 3
            '1 - Informe de Anticipos
            '3 - Factura de Anticipos
            If b Then
                If txtcodigo(6).Text = "" Or txtcodigo(7) = "" Then
                    MsgBox "Debe introducir obligatoriamente el rango de fechas.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(6)
                End If
            End If
            If b Then
                If txtcodigo(15).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de Anticipo.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(15)
                End If
            End If
            
       Case 2 'Prevision de pagos
            If b Then
                If txtcodigo(6).Text = "" Or txtcodigo(7) = "" Then
                    MsgBox "Para realizar la Previsión de Pago de Anticipos debe introducir obligatoriamente el rango de fechas.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(6)
                End If
            End If
       
       Case 5 'Deshacer proceso de facturacion de anticipos
            If txtcodigo(9).Text = "" Or txtcodigo(10).Text = "" Then
                MsgBox "Debe introducir la primera y última factura de la Facturación de Anticipos", vbExclamation
                b = False
                PonerFoco txtcodigo(9)
'            Else
'                ' si la factura hasta no coincide con el contador de stipom no seguir
'                Set vCont = New CTiposMov
'                If vCont.leer("FAA") Then
'                    If vCont.Contador <> CLng(txtCodigo(10).Text) Then
'                        MsgBox "La Factura hasta no es el último número de Factura de Anticipos. Revise.", vbExclamation
'                        b = False
'                    End If
'                End If
'                Set vCont = Nothing
            End If
            
            If b Then
                If txtcodigo(11).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Anticipo.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(11)
                End If
            End If
    
        Case 6 ' factura de ventas campo (anticipo o liquidacion)
            If b Then
                If txtcodigo(14).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de Factura.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(14)
                End If
            End If
        
        Case 7 'deshacer facturacion de venta campo ( anticipo o liquidacion )
            If txtcodigo(9).Text = "" Or txtcodigo(10).Text = "" Then
                MsgBox "Debe introducir la primera y última factura de la Facturación", vbExclamation
                b = False
                PonerFoco txtcodigo(9)
'            Else
'                ' si la factura hasta no coincide con el contador de stipom no seguir
'                Select Case Combo1(1).ListIndex
'                    Case 0
'                        TipoMov = "FAC"
'                    Case 1
'                        TipoMov = "FLC"
'                End Select
'
'                Set vCont = New CTiposMov
'                If vCont.leer(TipoMov) Then
'                    If vCont.Contador <> CLng(txtCodigo(10).Text) Then
'                        MsgBox "La Factura hasta no es el último número de Factura. Revise.", vbExclamation
'                        b = False
'                    End If
'                End If
'                Set vCont = Nothing
            End If
            
            If b Then
                If txtcodigo(11).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Factura.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(11)
                End If
            End If
            ' comprobamos que si son anticipos no esten liquidados
            If b And tipoMov = "FAC" Then
                If AnticiposLiquidados(tipoMov, txtcodigo(9).Text, txtcodigo(10).Text, txtcodigo(11).Text) Then
                    MsgBox "Hay Facturas de Anticipos que han sido liquidadas. Revise.", vbExclamation
                    b = False
                    PonerFocoBtn cmdCancelDesF
                End If
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

Private Function CargarTemporalAnticipos(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As CSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticipos = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid,"
    SQL = SQL & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilos "
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto, baseimpo, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionHorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!CodSocio
        VarieAnt = RS!CodVarie
        NVarieAnt = RS!nomvarie
        
        Set vSocio = Nothing
        Set vSocio = New CSocio
        If vSocio.LeerDatos(RS!CodSocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!CodSocio), vParamAplic.SeccionHorto) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIVA, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not RS.EOF
        '++monica:28/07/2009 añadida la segunda condicion
        If VarieAnt <> RS!CodVarie Or SocioAnt <> RS!CodSocio Then
            
            ImpoIva = Round2(baseimpo * ImporteSinFormato(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
            
            VarieAnt = RS!CodVarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
        End If
        
        If RS!CodSocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New CSocio
            If vSocio.LeerDatos(RS!CodSocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!CodSocio), vParamAplic.SeccionHorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIVA, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        Recolect = DBLet(RS!Recolect, "N")
        Select Case Recolect
            Case 0
                baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * RS!precoop, 2)
            Case 1
                baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * RS!presocio, 2)
        End Select
            
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        ImpoIva = Round2(baseimpo * ImporteSinFormato(vPorcIva) / 100, 2)
    
        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                PorcReten = 0
        End Select
    
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalAnticipos = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function




Private Function HayPreciosVariedades(tipo As Byte, cTabla As String, cWhere As String, TipoPrecio As Byte) As Boolean
'Comprobar si hay precios para cada una de las variedades seleccionadas
' tipo: 0=anticipos
'       1=liquidaciones
' tipoprecio: 0 = precio recolectado cooperativa
'             1 = precio recolectado socio
'             2 = precio recolectado socio y cooperativa
Dim SQL As String
Dim vPrecios As CPrecios
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim Sql2 As String

    On Error GoTo eHayPreciosVariedades
    
    HayPreciosVariedades = False
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    SQL = "Select distinct rhisfruta.codvarie FROM " & QuitarCaracterACadena(cTabla, "_1")
    Sql2 = "Select distinct rhisfruta_clasif.codvarie, rhisfruta_clasif.codcalid FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
        Sql2 = Sql2 & " where " & cWhere
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True
    ' comprobamos que existen registros para todos las variedades / calidades seleccionadas
    While Not RS.EOF And b
        Set vPrecios = New CPrecios
        b = vPrecios.Leer(CStr(tipo), CStr(RS.Fields(0).Value), txtcodigo(6).Text, txtcodigo(7).Text)
'        If b Then b = vPrecios.ExistenPreciosCalidades
        If b Then b = vPrecios.ExisteAlgunPrecioCalidad(Sql2, TipoPrecio)
        Set vPrecios = Nothing
        
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
    
    HayPreciosVariedades = b
    Exit Function
    
eHayPreciosVariedades:
    MuestraError Err.nume, "Comprobando si hay precios en variedades", Err.Description
End Function


Private Function TotalFacturas(cTabla As String, cWhere As String) As Long
Dim SQL As String

    TotalFacturas = 0
    
    SQL = "SELECT  count(distinct rhisfruta.codsocio) "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If

    TotalFacturas = TotalRegistros(SQL)

End Function

Private Sub ActivarCLAVE()
Dim I As Integer
    
    For I = 9 To 11
        txtcodigo(I).Enabled = False
    Next I
    txtcodigo(8).Enabled = True
    imgFec(6).Enabled = False
    CmdAcepDesF.Enabled = False
    cmdCancelDesF.Enabled = True
    Combo1(1).Enabled = False
End Sub

Private Sub DesactivarCLAVE()
Dim I As Integer

    For I = 9 To 11
        txtcodigo(I).Enabled = True
    Next I
    txtcodigo(8).Enabled = False
    imgFec(6).Enabled = True
    CmdAcepDesF.Enabled = True
    Combo1(1).Enabled = True
End Sub

Private Sub CargaCombo()

    On Error GoTo ErrCarga
    
    ' Tipo de facturacion venta campo (anticipo o liquidacion)
    ' para generacion de factura
    Combo1(0).Clear
    Combo1(0).AddItem "Anticipo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Liquidación"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    ' Tipo de facturacion venta campo (anticipo o liquidacion)
    ' para deshacer proceso de facturacion de venta campo
    Combo1(1).Clear
    Combo1(1).AddItem "Anticipo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Liquidación"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    'recolectado por
    Combo1(2).Clear
    Combo1(2).AddItem "Cooperativa"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Socio"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Todos"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
    
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub

Private Sub CargarListView(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

 '   ListView1.ColumnHeaders.Add , , "Tipo", 650
    ListView1(Index).ColumnHeaders.Add , , "Tipo Movimiento", 2750
    
    SQL = "SELECT codtipom, nomtipom "
    SQL = SQL & " FROM usuarios.stipom "
    SQL = SQL & " WHERE stipom.tipodocu in (1,2,3,4)"
    SQL = SQL & " ORDER BY codtipom "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Set ItmX = ListView1(Index).ListItems.Add
        ItmX.Text = RS.Fields(1).Value ' Format(Rs.Fields(0).Value)
        ItmX.Key = RS.Fields(0).Value
'        ItmX.SubItems(1) = Rs.Fields(1).Value
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Tipo de Facturas.", Err.Description
    End If
End Sub


Private Function NroTotalMovimientos(tipo As Byte) As Long
' Tipo: 1 - anticipos
'       2 - liquidacion
'       3 - anticipos venta campo
'       4 - liquidacion venta campo
Dim SQL As String
    
    SQL = "select distinct "
    Select Case tipo
        Case 1
            SQL = SQL & " CodTipomAnt "
        Case 2
            SQL = SQL & " codtipomliq "
        Case 3
            SQL = SQL & " codtipomantvc "
        Case 4
            SQL = SQL & " codtipomliqvc "
    End Select
    
    SQL = SQL & " from rcoope, usuarios.stipom stipom "
    SQL = SQL & " WHERE stipom.tipodocu=" & tipo
    SQL = SQL & " and stipom.codtipom = rcoope."
    Select Case tipo
        Case 1
            SQL = SQL & "CodTipomAnt "
        Case 2
            SQL = SQL & "codtipomliq "
        Case 3
            SQL = SQL & "codtipomantvc "
        Case 4
            SQL = SQL & "codtipomliqvc "
    End Select
    
    NroTotalMovimientos = TotalRegistrosConsulta(SQL)

End Function



Private Function GeneraFicheroModelo(tipo As Byte, pTabla As String, pWhere As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim Cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim vSocio As CSocio
Dim b As Boolean
Dim NRegs As Long
Dim total As Variant

Dim cTabla As String
Dim vWhere As String


    On Error GoTo EGen
    GeneraFicheroModelo = False
    
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
    
    Open App.Path & "\modelo.txt" For Output As #NFic
    
    Select Case tipo
        Case 0 ' MODELO 190
            Aux = "select count(*), sum(rfactsoc.basereten), sum(rfactsoc.impreten) "
            Aux = Aux & " from " & cTabla
            Aux = Aux & " where " & vWhere
                
            Set RS = New ADODB.Recordset
            RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
            'CABECERA
            Cabecera190a NFic, CLng(DBLet(RS.Fields(0).Value, "N"))
            Cabecera190b NFic, CLng(DBLet(RS.Fields(0).Value, "N")), CCur(DBLet(RS.Fields(1).Value, "N")), CCur(DBLet(RS.Fields(2).Value, "N"))
            
            Set RS = Nothing
            
            'Imprimimos las lineas
            Aux = "select rfactsoc.codsocio, sum(rfactsoc.basereten), sum(rfactsoc.impreten) "
            Aux = Aux & " from " & cTabla
            Aux = Aux & " where " & vWhere
            Aux = Aux & " group by 1 "
            Aux = Aux & " having sum(rfactsoc.basereten) <> 0 "
            Aux = Aux & " order by 1 "
            
            Set RS = New ADODB.Recordset
            RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If RS.EOF Then
                'No hayningun registro
            Else
                b = True
                Regs = 0
                While Not RS.EOF And b
                    Regs = Regs + 1
                    Set vSocio = New CSocio
                    
                    If vSocio.LeerDatos(DBLet(RS!CodSocio, "N")) Then
                        Linea190 NFic, vSocio, RS
                    Else
                        b = False
                    End If
                    
                    Set vSocio = Nothing
                    RS.MoveNext
                Wend
            End If
            RS.Close
            Set RS = Nothing
            
        Case 1 ' MODELO 346
            cTabla = "(" & cTabla & ") INNER JOIN variedades ON rfactsoc_variedad.codvarie = variedades.codvarie "
            cTabla = "(" & cTabla & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
            cTabla = "(" & cTabla & ") INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
            Aux = "select rfactsoc.codsocio, grupopro.codgrupo, sum(rfactsoc_variedad.imporvar) "
            Aux = Aux & " from " & cTabla
            Aux = Aux & " where " & vWhere & " and grupopro.codgrupo in (4,5) " ' algarrobos y olivos
            Aux = Aux & " group by rfactsoc.codsocio, grupopro.codgrupo "
            Aux = Aux & "  union "
            Aux = Aux & " select rfactsoc.codsocio, 0, sum(rfactsoc_variedad.imporvar) "
            Aux = Aux & " from " & cTabla
            Aux = Aux & " where " & vWhere & " and not grupopro.codgrupo in (4,5) " ' el resto
            Aux = Aux & " group by rfactsoc.codsocio, grupopro.codgrupo "
            Aux = Aux & " order by 1,2"
        
            NRegs = TotalRegistrosConsulta(Aux)
        
            If NRegs <> 0 Then
                Aux2 = "select sum(rfactsoc_variedad.imporvar) from " & cTabla
                Aux2 = Aux2 & " where " & vWhere
                
                total = DevuelveValor(Aux2)
            
                Cabecera346 NFic, NRegs, CCur(total)
            
                Set RS = New ADODB.Recordset
                RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If RS.EOF Then
                    'No hayningun registro
                Else
                    b = True
                    Regs = 0
                    While Not RS.EOF And b
                        Regs = Regs + 1
                        Set vSocio = New CSocio
                        
                        If vSocio.LeerDatos(DBLet(RS!CodSocio, "N")) Then
                            Linea346 NFic, vSocio, RS
                        Else
                            b = False
                        End If
                        
                        Set vSocio = Nothing
                        RS.MoveNext
                    Wend
                End If
                RS.Close
                Set RS = Nothing
                
            End If
    End Select
    Close (NFic)
    
    If Regs > 0 Then GeneraFicheroModelo = True
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

    Me.CommonDialog1.DefaultExt = "txt"
    
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    ' copiamos el primer fichero
    Select Case OpcionListado
        Case 10
            CommonDialog1.FileName = "modelo190.txt"
        Case 11
            CommonDialog1.FileName = "modelo346.txt"
    End Select
        
    Me.CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        FileCopy App.Path & "\modelo.txt", CommonDialog1.FileName
    End If
    
    CopiarFichero = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function


Private Sub Cabecera190a(NFich As Integer, NRegs As Long)
Dim Cad As String

   'TIPO DE REGISTRO 0:PRESENTACION COLECTIVA
    Cad = "0190"                                                'p.1
    Cad = Cad & Format(txtcodigo(30).Text, "0000")              'p.5 año ejercicio
    Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 9)     'p.9 cif empresa
    Cad = Cad & RellenaABlancos(vParam.NombreEmpresa, True, 40) 'p.18 nombre empresa
    Cad = Cad & RellenaABlancos(txtcodigo(40).Text, True, 2)    'p.58 siglas
    Cad = Cad & RellenaABlancos(txtcodigo(38).Text, True, 20)   'p.60 nombre de la via publica
    Cad = Cad & RellenaAceros(txtcodigo(39).Text, False, 5)     'p.80 numero de la via
    Cad = Cad & Space(6)                                        'p.85 6 blancos en la posicion 84
    Cad = Cad & RellenaABlancos(vParam.CPostal, True, 5)        'p.91 codigo postal
    Cad = Cad & RellenaABlancos(vParam.Poblacion, True, 12)     'p.96 poblacion
    Cad = Cad & RellenaABlancos(Mid(vParam.CPostal, 1, 2), True, 2)    'p.108 46
    Cad = Cad & "00001"                                         'p.110
    
    Cad = Cad & RellenaAceros(CStr(NRegs), False, 9)            'p.115 numero de registros
    Cad = Cad & "D"                                             'p.124
    Cad = Cad & RellenaAceros(txtcodigo(37).Text, False, 9)     'p.125 telefono
    Cad = Cad & RellenaABlancos(txtcodigo(36).Text, True, 40)   'p.134 persona de contacto
    Cad = Cad & RellenaABlancos(" ", True, 64)                  'p.174
    Cad = Cad & RellenaABlancos(" ", True, 13)                  'p.238

    Print #NFich, Cad
End Sub

Private Sub Cabecera190b(NFich As Integer, NRegs As Currency, ImpReten As Currency, BaseReten As Currency)
Dim Cad As String

'TIPO DE REGISTRO 1:REGISTRO DEL RETENEDOR}
    
    Cad = "1190"                                                  'p.1
    Cad = Cad & Format(txtcodigo(30).Text, "0000")                'p.5 año de ejercicio
    Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 9)       'p.9 cif empresa
    Cad = Cad & RellenaABlancos(vParam.NombreEmpresa, True, 40)   'p.18 nombre de empresa
    Cad = Cad & "D"                                               'p.58
    Cad = Cad & RellenaAceros(txtcodigo(37).Text, True, 9)        'p.59 telefono
    Cad = Cad & RellenaABlancos(txtcodigo(36).Text, True, 40)     'p.68 persona de contacto
    Cad = Cad & RellenaAceros(txtcodigo(31).Text, True, 13)       'p.108 nro de justificante
    Cad = Cad & Space(2)                                          'p.121
    Cad = Cad & RellenaAceros("0", True, 13)                      'p.123 13 ceros
    Cad = Cad & Format(NRegs, "000000000")                        'p.136 nro de registros

    If BaseReten < 0 Then
        Cad = Cad & "N"                                           'p.145
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(BaseReten * (-1) * 100)), False, 15)    'p.145
    Else
        Cad = Cad & " "                                           'p.145
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(BaseReten * 100)), False, 15)           'p.145
    End If
              
    If ImpReten < 0 Then                                          'p.161
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(ImpReten * (-1) * 100)), False, 15)
    Else
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(ImpReten * 100)), False, 15)
    End If
    Cad = Cad & Space(62)                                         'p.176
    Cad = Cad & Space(13)                                         'p.238

    Print #NFich, Cad

End Sub


Private Sub Linea190(NFich As Integer, vSocio As CSocio, ByRef RS As ADODB.Recordset)
Dim Cad As String

    Cad = "2190"                                                'p.1
    Cad = Cad & Format(txtcodigo(30).Text, "0000")              'p.5 año ejercicio
    Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 9)     'p.9 cif empresa
    Cad = Cad & RellenaABlancos(vSocio.NIF, True, 9)            'p.18 nifsocio
    Cad = Cad & Space(9)                                        'p.27 nif del representante legal
    Cad = Cad & RellenaABlancos(vSocio.Nombre, True, 40)        'p.36 nombre socio
    Cad = Cad & RellenaABlancos(Mid(vSocio.CPostal, 1, 2), True, 2) 'p.76 codpobla[1,2]
    Cad = Cad & "H"                                             'p.78
    Cad = Cad & "01"                                            'p.79
    Cad = Cad & " "                                             'p.81
    
    If DBLet(RS.Fields(1).Value, "N") < 0 Then                  'p.82 base de retencion
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(1).Value, "N") * (-1) * 100)), False, 13)
    Else
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(1).Value, "N") * 100)), False, 13)
    End If
    
    If DBLet(RS.Fields(2).Value, "N") < 0 Then                  'p.95 importe de retencion
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * (-1) * 100)), False, 13)
    Else
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * 100)), False, 13)
    End If
    
    Cad = Cad & " "                                             'p.108
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.109
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.122
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.135
    Cad = Cad & RellenaAceros("0", True, 4)                     'p.148
    Cad = Cad & "0"                                             'p.152
    Cad = Cad & RellenaAceros("0", True, 5)                     'p.153
    Cad = Cad & RellenaABlancos(" ", True, 9)                   'p.158
    Cad = Cad & String(84, "0")                                 'p.167
    
    Print #NFich, Cad
End Sub


Private Sub Cabecera346(NFich As Integer, NRegs As Long, total As Currency)
Dim Cad As String

   'TIPO DE REGISTRO 0:PRESENTACION COLECTIVA
    Cad = "1346"                                                'p.1
    Cad = Cad & Format(txtcodigo(30).Text, "0000")              'p.5 año ejercicio
    Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 9)     'p.9 cif empresa
    Cad = Cad & RellenaABlancos(vParam.NombreEmpresa, True, 40) 'p.18 nombre empresa
    Cad = Cad & "D"    'p.58 siglas
    Cad = Cad & RellenaAceros(txtcodigo(37).Text, False, 9)     'p.59 telefono
    Cad = Cad & RellenaABlancos(txtcodigo(36).Text, True, 40)   'p.68 persona de contacto
    Cad = Cad & RellenaAceros(txtcodigo(31).Text, False, 13)    'p.108 nro justificante
    Cad = Cad & Space(2)                                        ' contar posiciones en multibase
    Cad = Cad & String(13, "0")                                 'p.122
    Cad = Cad & RellenaAceros(CStr(NRegs), False, 9)            'p.136 numero de registros
    Cad = Cad & Space(1)                                        ' contar posiciones en multibase
    Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(total * 100)), False, 17)  'p.146 importe total
    Cad = Cad & Space(87)                                       'p.163
    
    Print #NFich, Cad
End Sub


Private Sub Linea346(NFich As Integer, vSocio As CSocio, ByRef RS As ADODB.Recordset)
Dim Cad As String
          
    Cad = "2346"                                                'p.1
    Cad = Cad & Format(txtcodigo(30).Text, "0000")              'p.5 año ejercicio
    Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 9)     'p.9 cif empresa
    Cad = Cad & RellenaABlancos(vSocio.NIF, True, 18)            'p.18 nifsocio
    Cad = Cad & RellenaABlancos(vSocio.Nombre, True, 40)        'p.36 nombre socio
    Cad = Cad & RellenaABlancos(Mid(vSocio.CPostal, 1, 2), True, 2) 'p.76 codpobla[1,2]
    Cad = Cad & "A"                                             'p.78
    
    Select Case DBLet(RS.Fields(1).Value, "N")
        Case 0
            Cad = Cad & "6"                                             'p.79
        Case 4
            Cad = Cad & "1"                                             'p.79
        Case 5
            Cad = Cad & "1"                                             'p.79
    End Select
    
    Cad = Cad & " "                                             ' contar posiciones en multibase
    Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * 100)), False, 14) 'p.81 base imponible
    Cad = Cad & RellenaAceros("0", True, 4)                     'p.95
    
    Select Case DBLet(RS.Fields(1).Value, "N")
        Case 0
            Cad = Cad & RellenaABlancos("INDEMNIZACION AGROSEGURO", True, 57)   'p.99
        Case 4
            Cad = Cad & RellenaABlancos("CULTIVO ALGARROBO", True, 57)          'p.99
        Case 5
            Cad = Cad & RellenaABlancos("CULTIVO OLIVO", True, 57)              'p.99
    End Select
        
    Cad = Cad & Space(94)                                       'p.156
    
    
    Print #NFich, Cad
End Sub


Private Function CargarTemporalLiquidacion(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CampoAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bruto As Currency
Dim ImpoIva As Currency
Dim ImpoGastos As Currency
Dim ImpoReten As Currency
Dim ImpoAport As Currency
Dim Anticipos As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim vPorcGasto As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

Dim vSocio As CSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacion = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid,"
    SQL = SQL & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilos "
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  gastos,    anticipos, baseimpo, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionHorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!CodSocio
        VarieAnt = RS!CodVarie
        NVarieAnt = RS!nomvarie
        CampoAnt = RS!CodCampo
        
        Set vSocio = Nothing
        Set vSocio = New CSocio
        If vSocio.LeerDatos(RS!CodSocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!CodSocio), vParamAplic.SeccionHorto) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIVA, "N")
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not RS.EOF
        If CampoAnt <> RS!CodCampo Then
            ' gastos por campo
            Sql4 = "select sum(imptrans) + sum(impacarr) + sum(imprecol) + sum(imppenal) from rhisfruta "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
            Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
            Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F")
            
            ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
            
            CampoAnt = RS!CodCampo
        End If
    
        ' 23/07/2009: añadido el or con la segunda condicion
        If VarieAnt <> RS!CodVarie Or SocioAnt <> RS!CodSocio Then
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
            
            Anticipos = DevuelveValor(Sql4)
            
            Bruto = baseimpo
            
            ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
            
            baseimpo = baseimpo - ImpoGastos - Anticipos
            
            ImpoIva = Round2((baseimpo) * ImporteSinFormato(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    PorcReten = 0
            End Select
        
            ImpoAport = Round2((Bruto - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
            
            VarieAnt = RS!CodVarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
            
            ImpoGastos = 0
            Anticipos = 0
            
        End If
        
        If RS!CodSocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New CSocio
            If vSocio.LeerDatos(RS!CodSocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!CodSocio), vParamAplic.SeccionHorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIVA, "N")
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        Recolect = DBLet(RS!Recolect, "N")
        Select Case Recolect
            Case 0
                baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * RS!precoop, 2)
            Case 1
                baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * RS!presocio, 2)
        End Select
            
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        Bruto = baseimpo
        
        ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
        
        baseimpo = baseimpo - ImpoGastos - Anticipos
        
        ImpoIva = Round2((baseimpo) * ImporteSinFormato(vPorcIva) / 100, 2)
        
        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                PorcReten = 0
        End Select
    
        ImpoAport = Round2((Bruto - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
    
        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
        Sql1 = Sql1 & DBSet(Bruto, "N") & ","
        Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
        Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
        Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalLiquidacion = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function




Private Function ActualizarRegistrosFac(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim SQL As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistrosFac = False
    SQL = "update " & cTabla & ", usuarios.stipom set impreso = 1 "
    SQL = SQL & " where usuarios.stipom.codtipom = rfactsoc.codtipom "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " and " & cWhere
    End If
    
    conn.Execute SQL
    
    ActualizarRegistrosFac = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function



Private Function CargarTemporalAnticiposGastos(cTabla As String, cWhere As String, Cad As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim baseimpo As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim KilosGastos As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency
    
Dim vSocio As CSocio
Dim vSeccion As CSeccion
Dim HayReg As Boolean

Dim Sql3 As String
Dim Importe As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposGastos = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,  "
    SQL = SQL & "rcalidad.gastosrec, " ' sum(rhisfruta.imprecol) as importe, "
    SQL = SQL & "sum(rhisfruta_clasif.kilosnet) as kilos"
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5 "
    SQL = SQL & " order by 1, 2, 3, 4, 5 "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie,kggastos, baseimpo, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac, kgneto
    Sql2 = Sql2 & " porcen2, importeb1, importeb2, importeb3) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionHorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!CodSocio
        VarieAnt = RS!CodVarie
        NVarieAnt = RS!nomvarie
        
        Set vSocio = Nothing
        Set vSocio = New CSocio
        If vSocio.LeerDatos(RS!CodSocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!CodSocio), vParamAplic.SeccionHorto) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIVA, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not RS.EOF
        If VarieAnt <> RS!CodVarie Or SocioAnt <> RS!CodSocio Then
            
            ImpoIva = Round2(baseimpo * ImporteSinFormato(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosGastos, "N") & "," & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "," & DBSet(KilosNet, "N") & "),"
            
            VarieAnt = RS!CodVarie
            
            baseimpo = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
            KilosGastos = 0
        End If
        
        If RS!CodSocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New CSocio
            If vSocio.LeerDatos(RS!CodSocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!CodSocio), vParamAplic.SeccionHorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIVA, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        If DBLet(RS!gastosrec, "N") = 1 Then
            KilosGastos = KilosGastos + DBLet(RS!Kilos, "N")
        
        
            ' insertar linea de variedad, campo
            Sql3 = "select sum(imprecol) from rhisfruta where "
            If Cad <> "" Then Sql3 = Sql3 & Cad & " and "
            Sql3 = Sql3 & " rhisfruta.codvarie = " & DBSet(RS!CodVarie, "N")
            Sql3 = Sql3 & " and codcampo = " & DBSet(RS!CodCampo, "N") & " and codsocio = " & DBSet(RS!CodSocio, "N")
            
            Importe = DevuelveValor(Sql3)
        
        
        
        
            baseimpo = baseimpo + Importe  '+ DBLet(rs!Importe, "N")
        End If
        
            
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        ImpoIva = Round2(baseimpo * ImporteSinFormato(vPorcIva) / 100, 2)
    
        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                PorcReten = 0
        End Select
    
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosGastos, "N") & "," & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "," & DBSet(KilosNet, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalAnticiposGastos = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function



