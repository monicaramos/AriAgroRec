VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTRAFactAlb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   13215
   Icon            =   "frmTRAFactAlb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7800
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameFacturar 
      Height          =   5610
      Left            =   30
      TabIndex        =   11
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3285
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5190
         TabIndex        =   10
         Top             =   4935
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1095
         Width           =   990
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1455
         Width           =   990
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   1095
         Width           =   3195
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   1455
         Width           =   3195
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmTRAFactAlb.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Frame FrameFechaAnt 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   645
         Left            =   390
         TabIndex        =   13
         Top             =   3630
         Width           =   2745
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   960
            Picture         =   "frmTRAFactAlb.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Factura"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   25
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1035
         End
      End
      Begin VB.Frame FrameOpciones 
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   3780
         TabIndex        =   12
         Top             =   3270
         Width           =   2115
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Factura"
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   8
            Top             =   300
            Width           =   1995
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   2
            Left            =   270
            TabIndex        =   7
            Top             =   0
            Width           =   1965
         End
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   9
         Top             =   4950
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   4620
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1350
         Picture         =   "frmTRAFactAlb.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   405
         TabIndex        =   33
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   735
         TabIndex        =   32
         Top             =   2940
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   735
         TabIndex        =   31
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   720
         TabIndex        =   30
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   720
         TabIndex        =   29
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Facturación de Transporte"
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
         TabIndex        =   28
         Top             =   345
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Transportista"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   390
         TabIndex        =   27
         Top             =   900
         Width           =   915
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1350
         MouseIcon       =   "frmTRAFactAlb.frx":01AD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar transportista"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1350
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar transportista"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1350
         Picture         =   "frmTRAFactAlb.frx":01B1
         ToolTipText     =   "Buscar fecha"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   390
         TabIndex        =   26
         Top             =   1830
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   765
         TabIndex        =   25
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   765
         TabIndex        =   24
         Top             =   2475
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1350
         MouseIcon       =   "frmTRAFactAlb.frx":023C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1350
         MouseIcon       =   "frmTRAFactAlb.frx":038E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   10
         Left            =   390
         TabIndex        =   23
         Top             =   4950
         Width           =   3525
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   12
         Left            =   390
         TabIndex        =   22
         Top             =   5160
         Width           =   3615
      End
   End
   Begin VB.Frame FrameFacturarSocio 
      Height          =   5610
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   525
         Left            =   3600
         TabIndex        =   89
         Top             =   2910
         Width           =   2505
         Begin VB.CheckBox Check1 
            Caption         =   "Terceros"
            Height          =   195
            Index           =   11
            Left            =   420
            TabIndex        =   90
            Top             =   240
            Width           =   1035
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   1470
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.CommandButton CmdAcepFTraSoc 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   63
         Top             =   4950
         Width           =   975
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   3750
         TabIndex        =   73
         Top             =   3540
         Width           =   2655
         Begin VB.CheckBox Check1 
            Caption         =   "Reimprimir Rectificadas"
            Height          =   195
            Index           =   4
            Left            =   270
            TabIndex        =   91
            Top             =   570
            Width           =   1995
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   1
            Left            =   270
            TabIndex        =   75
            Top             =   0
            Width           =   1965
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Factura"
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   74
            Top             =   300
            Width           =   1995
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   1
            Left            =   2280
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   540
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   645
         Left            =   390
         TabIndex        =   71
         Top             =   3630
         Width           =   2745
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   62
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Factura"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Width           =   1035
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   960
            Picture         =   "frmTRAFactAlb.frx":04E0
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text5"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   59
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   58
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Height          =   440
         Left            =   7860
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmTRAFactAlb.frx":056B
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "Text5"
         Top             =   1455
         Width           =   3195
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text5"
         Top             =   1095
         Width           =   3195
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   57
         Top             =   1455
         Width           =   990
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   56
         Top             =   1095
         Width           =   990
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5190
         TabIndex        =   65
         Top             =   4950
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   61
         Top             =   3285
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   60
         Top             =   2880
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   255
         Left            =   360
         TabIndex        =   76
         Top             =   4620
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   14
         Left            =   390
         TabIndex        =   88
         Top             =   5160
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   13
         Left            =   390
         TabIndex        =   87
         Top             =   4950
         Width           =   3525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1350
         MouseIcon       =   "frmTRAFactAlb.frx":05F6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1350
         MouseIcon       =   "frmTRAFactAlb.frx":0748
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   9
         Left            =   765
         TabIndex        =   86
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   8
         Left            =   765
         TabIndex        =   85
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   7
         Left            =   390
         TabIndex        =   84
         Top             =   1830
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1350
         Picture         =   "frmTRAFactAlb.frx":089A
         ToolTipText     =   "Buscar fecha"
         Top             =   3270
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1350
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1350
         MouseIcon       =   "frmTRAFactAlb.frx":0925
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   6
         Left            =   390
         TabIndex        =   83
         Top             =   900
         Width           =   405
      End
      Begin VB.Label Label5 
         Caption         =   "Facturación de Transporte a Socio"
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
         TabIndex        =   82
         Top             =   345
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   81
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   80
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   735
         TabIndex        =   79
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   735
         TabIndex        =   78
         Top             =   2940
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   1
         Left            =   405
         TabIndex        =   77
         Top             =   2700
         Width           =   450
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1350
         Picture         =   "frmTRAFactAlb.frx":0929
         ToolTipText     =   "Buscar fecha"
         Top             =   2880
         Width           =   240
      End
   End
   Begin VB.Frame FrameReimpresion 
      Height          =   5220
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   6675
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   40
         Top             =   3780
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   39
         Top             =   3405
         Width           =   830
      End
      Begin VB.CommandButton cmdAceptarReimp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   41
         Top             =   4275
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelReimp 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5340
         TabIndex        =   42
         Top             =   4275
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   37
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
         TabIndex        =   38
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
         TabIndex        =   35
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
         TabIndex        =   36
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1740
         Width           =   830
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1470
         MouseIcon       =   "frmTRAFactAlb.frx":09B4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar transportista"
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1470
         MouseIcon       =   "frmTRAFactAlb.frx":0B06
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trasnportista"
         Top             =   3405
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Transportista"
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
         TabIndex        =   54
         Top             =   3165
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   870
         TabIndex        =   53
         Top             =   3780
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   855
         TabIndex        =   52
         Top             =   3405
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1485
         Picture         =   "frmTRAFactAlb.frx":0C58
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1485
         Picture         =   "frmTRAFactAlb.frx":0CE3
         ToolTipText     =   "Buscar fecha"
         Top             =   2415
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   825
         TabIndex        =   51
         Top             =   2775
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   825
         TabIndex        =   50
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   465
         TabIndex        =   49
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
         TabIndex        =   48
         Top             =   1125
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   47
         Top             =   1395
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   46
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Reimpresión de Facturas Transporte"
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
         TabIndex        =   45
         Top             =   315
         Width           =   5760
      End
   End
End
Attribute VB_Name = "frmTRAFactAlb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
'(ver opciones en frmListado)

' 1.- Facturacion de albaranes de transporte
' 2.- Reimpresion de facturas de transporte

      
      
'Alguna opcion mas
'                   1000.-  Es cuando paso pedido a albaran y este a factura en el mismo proceso
'                   1001.-  Facturar un unico albaran
      
      
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir


Public CodClien As String 'Para seleccionar inicialmente las ofertas del Proveedor

'#Laura 14/11/2006 Recuperar facturas Alzira
Public EstaRecupFact As Boolean ' si esta recuperando facturas (para albaranes de mostrador)


Private WithEvents frmCla As frmComercial 'Ayuda de Clases de comercial
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMens4 As frmMensajes 'Mensajes
Attribute frmMens4.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1


'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

'Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Private WithEvents frmTra As frmManTranspor
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios
Attribute frmSoc.VB_VarHelpID = -1
'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean
'-------------------------------------



Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim indice As Integer



Dim indCodigo As Integer 'indice para txtCodigo

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


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
Dim Sql2 As String
Dim vClien As cSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String

    b = True
    
    Select Case OpcionListado
        Case 1 ' facturacion de albaranes de transporte
            If txtcodigo(15).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la Fecha de Liquidación.", vbExclamation
                b = False
                PonerFoco txtcodigo(15)
            End If
    
        Case 3 ' facturacion de albaranes de transporte y acarreo a socio
            If txtcodigo(17).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la Fecha de Liquidación.", vbExclamation
                b = False
                PonerFoco txtcodigo(17)
            End If
    
    End Select
    DatosOk = b

End Function



Private Sub CmdAcepFTraSoc_Click()
'Facturacion de Albaranes
Dim campo As String, Cad As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
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
Dim Tabla1 As String


Dim Nregs As Long
Dim FecFac As Date
Dim TipoPrec As Byte ' 0 anticipos
                     ' 1 liquidaciones
Dim b As Boolean
Dim Sql2 As String
    
    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(10).Text)
        cHasta = Trim(txtcodigo(11).Text)
        nDesde = txtNombre(10).Text
        nHasta = txtNombre(11).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rclasifica.codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(14).Text)
        cHasta = Trim(txtcodigo(16).Text)
        nDesde = txtNombre(14).Text
        nHasta = txtNombre(16).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(14).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(14).Text, "N")
        If txtcodigo(16).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(16).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(8).Text)
        cHasta = Trim(txtcodigo(9).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rclasifica.fechaent}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        nTabla = "(((rclasifica INNER JOIN rsocios ON rclasifica.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN variedades ON rclasifica.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        
        Tabla1 = "(((rhisfruta INNER JOIN rhisfruta_entradas ON rhisfruta.numalbar = rhisfruta_entradas.numalbar and (rhisfruta.transportadopor = 1 or rhisfruta.recolect = 1) "
        Tabla1 = Tabla1 & " INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        Tabla1 = Tabla1 & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        Tabla1 = Tabla1 & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        Tabla1 = Tabla1 & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        Tabla1 = Tabla1 & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        
        cadSelect1 = Replace(Replace(cadSelect, "rclasifica", "rhisfruta_entradas"), "rhisfruta_entradas.codsocio", "rhisfruta.codsocio")
        
        If Not AnyadirAFormula(cadSelect, "({rclasifica.transportadopor} = 1 or {rclasifica.recolect} = 1)") Then Exit Sub
        
        
        If Not AnyadirAFormula(cadSelect, "{rclasifica.tipoentr} <> 1 ") Then Exit Sub
        If Not AnyadirAFormula(cadSelect1, "{rhisfruta.tipoentr} <> 1 ") Then Exit Sub
        
        
        '[Monica]10/10/2013: añadimos la condicion de que el socio sea tercero si lo han marcado (solo PICASSENT)
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
            If Check1(11).Value = 1 Then ' socios terceros
                If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadSelect1, "{rsocios.tipoprod} = 1") Then Exit Sub
            Else  ' socios no terceros
                If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
                If Not AnyadirAFormula(cadSelect1, "{rsocios.tipoprod} <> 1") Then Exit Sub
            End If
        End If
        
' no sé como

        '[Monica]26/12/2011: cambio estas dos condiciones:
'        If Not AnyadirAFormula(cadSelect, "{rclasifica.numnotac} not in (select numalbar from rfactsoc_albaran) ") Then Exit Sub
'
'        If Not AnyadirAFormula(cadSelect1, "{rhisfruta_entradas.numnotac} not in (select numalbar from rfactsoc_albaran) ") Then Exit Sub
'
'         por las siguientes: Solo si cambiamos el socio de las entradas podemos volver a facturarlas

        '[Monica]23/11/2015: si la factura esta rectificada solo cojo los albaranes de la factura rectificada
        If Me.Check1(4).Value = 0 Then  ' como estaba antes
            If Not AnyadirAFormula(cadSelect, "({rclasifica.numnotac},{rclasifica.codsocio}) not in (select numalbar, codsocio from rfactsoc_albaran INNER JOIN rfactsoc ON rfactsoc_albaran.codtipom = rfactsoc.codtipom and rfactsoc_albaran.numfactu = rfactsoc.numfactu and rfactsoc_albaran.fecfactu = rfactsoc.fecfactu) ") Then Exit Sub
    
            If Not AnyadirAFormula(cadSelect1, "({rhisfruta_entradas.numnotac},{rhisfruta.codsocio}) not in (select numalbar, codsocio from rfactsoc_albaran INNER JOIN rfactsoc ON rfactsoc_albaran.codtipom = rfactsoc.codtipom and rfactsoc_albaran.numfactu = rfactsoc.numfactu and rfactsoc_albaran.fecfactu = rfactsoc.fecfactu) ") Then Exit Sub
        Else
            ' si está marcado miramos los albaranes que estén en la rectificativa
            If Not AnyadirAFormula(cadSelect, "({rclasifica.numnotac},{rclasifica.codsocio}) in (select numalbar, codsocio from rfactsoc_albaran INNER JOIN rfactsoc ON rfactsoc_albaran.codtipom = rfactsoc.codtipom and rfactsoc_albaran.numfactu = rfactsoc.numfactu and rfactsoc_albaran.fecfactu = rfactsoc.fecfactu and rfactsoc.codtipom = 'FRS') ") Then Exit Sub
    
            If Not AnyadirAFormula(cadSelect1, "({rhisfruta_entradas.numnotac},{rhisfruta.codsocio}) in (select numalbar, codsocio from rfactsoc_albaran INNER JOIN rfactsoc ON rfactsoc_albaran.codtipom = rfactsoc.codtipom and rfactsoc_albaran.numfactu = rfactsoc.numfactu and rfactsoc_albaran.fecfactu = rfactsoc.fecfactu and rfactsoc.codtipom = 'FRS') ") Then Exit Sub
        End If

        cadNombreRPT = "rResumFacturas.rpt"
        
        If Check1(11).Value = 1 Then
            cadTitulo = "Resumen Informe Transporte/Recoleccion Socios "
        Else
            cadTitulo = "Resumen Facturas Transporte/Recoleccion Socios "
        End If
                
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInformeNew(nTabla, cadSelect, Tabla1, cadSelect1) Then
            b = FacturacionTransporteSocio(nTabla, cadSelect, Tabla1, cadSelect1, txtcodigo(17).Text, Me.Pb1, txtcodigo(8).Text, txtcodigo(9).Text, Check1(11).Value = 1)
            If b Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                               
                'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                If Me.Check1(1).Value Then
                    cadFormula = ""
                    CadParam = CadParam & "pFecFac= """ & txtcodigo(17).Text & """|"
                    numParam = numParam + 1
                    
                    '[Monica]10/10/2013: si son socios terceros (Picassent)
                    If Check1(11).Value = 1 Then
                        CadParam = CadParam & "pTitulo= ""Resumen Informe Transporte Tercero""|"
                        CadParam = CadParam & "pEsInforme=1|"
                    Else
                        CadParam = CadParam & "pTitulo= ""Resumen Facturación Transporte Socio""|"
                        CadParam = CadParam & "pEsInforme=0|"
                    End If
                    numParam = numParam + 2
                    
                    FecFac = CDate(txtcodigo(17).Text)
                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                    ConSubInforme = False
                    
                    LlamarImprimir
                End If
                'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE TRANSPORTE
                If Me.Check1(0).Value Then
                    cadFormula = ""
                    cadSelect = ""
                    
                    '[Monica]10/10/2013: si son socios terceros (Picassent)
                    If Check1(11).Value = 1 Then
                        cadAux = "({stipom.codtipom} = 'FTT')"
                    Else
                        cadAux = "({stipom.codtipom} = 'FTS')"
                    End If
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                    
                    '[Monica]10/10/2013: si son socios terceros (Picassent)
                    If Check1(11).Value = 1 Then
                        'Nº Factura
                        cadAux = "({rfactsoc.numfactu} IN [" & ListaFacturasGeneradas("FTT") & "])"
                    Else
                        'Nº Factura
                        cadAux = "({rfactsoc.numfactu} IN [" & ListaFacturasGeneradas("FTS") & "])"
                    End If
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                     
                    'Fecha de Factura
                    FecFac = CDate(txtcodigo(17).Text)
                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                   
                    indRPT = 23 'Impresion de facturas de transporte/recoleccion a socios
                    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                    'Nombre fichero .rpt a Imprimir
                    cadNombreRPT = nomDocu
                    
                    '[Monica]10/10/2013: si son socios terceros (Picassent)
                    'Nombre fichero .rpt a Imprimir
                    If Check1(11).Value = 0 Then
                        cadTitulo = "Reimpresión Facturas Transporte Recolección Socios"
                    Else
                        cadTitulo = "Reimpresión de Informes de Recolección a Socios"
                    End If
                    
                    ConSubInforme = True
                    
                    conSubRPT = ConSubInforme
                    
                    LlamarImprimir
                    
                    If frmVisReport.EstaImpreso Then
                        ActualizarRegistrosFacSoc "rfactsoc", cadSelect
                    End If
                End If
                'SALIR DE LA FACTURACION DE tranporte a socios
                cmdCancel_Click (1)
            End If
        Else
            MsgBox "No hay entradas a facturar.", vbExclamation
        End If
    End If
    Label2(13).visible = False
    Label2(14).visible = False

End Sub

Private Sub cmdAceptar_Click()
'Facturacion de Albaranes
Dim campo As String, Cad As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
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
Dim Tabla1 As String


Dim Nregs As Long
Dim FecFac As Date
Dim TipoPrec As Byte ' 0 anticipos
                     ' 1 liquidaciones
Dim b As Boolean
Dim Sql2 As String

Dim Sql4 As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H TRANSPORTISTA
        cDesde = Trim(txtcodigo(12).Text)
        cHasta = Trim(txtcodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rclasifica.codtrans}"
            TipCod = "T"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        '[Monica]07/11/2013: Listview para pocer seleccionar los transportistas a facturar
        Sql4 = ""
        If txtcodigo(12).Text <> "" Then Sql4 = Sql4 & " and rtransporte.codtrans >=" & DBSet(txtcodigo(12).Text, "T")
        If txtcodigo(13).Text <> "" Then Sql4 = Sql4 & " and rtransporte.codtrans <=" & DBSet(txtcodigo(13).Text, "T")
        
        
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
            Codigo = "{rclasifica.fechaent}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        nTabla = "(((rclasifica INNER JOIN rtransporte ON rclasifica.codtrans = rtransporte.codtrans) "
        nTabla = nTabla & " INNER JOIN variedades ON rclasifica.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        '[Monica]08/06/2010:
        ' solo seleccionamos los registros de los transportistas que se van a facturar
        If Not AnyadirAFormula(cadSelect, "{rtransporte.sefactura} = 1 ") Then Exit Sub
        
        Tabla1 = "((((rhisfruta INNER JOIN rhisfruta_entradas ON rhisfruta.numalbar = rhisfruta_entradas.numalbar and rhisfruta.transportadopor = 0) "
        Tabla1 = Tabla1 & " INNER JOIN rtransporte ON rhisfruta_entradas.codtrans = rtransporte.codtrans) "
        Tabla1 = Tabla1 & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        Tabla1 = Tabla1 & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        Tabla1 = Tabla1 & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        Tabla1 = Tabla1 & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        
        cadSelect1 = Replace(cadSelect, "rclasifica", "rhisfruta_entradas")
        
        If Not AnyadirAFormula(cadSelect, "{rclasifica.transportadopor} = 0") Then Exit Sub
        
        
        If Not AnyadirAFormula(cadSelect, "{rclasifica.tipoentr} <> 1 ") Then Exit Sub
        If Not AnyadirAFormula(cadSelect1, "{rhisfruta.tipoentr} <> 1 ") Then Exit Sub
        
        If Not AnyadirAFormula(cadSelect, "not ({rclasifica.numnotac}, {rclasifica.fechaent}) in (select numnotac, fechaent from rfacttra_albaran) ") Then Exit Sub
        
        If Not AnyadirAFormula(cadSelect1, "not ({rhisfruta_entradas.numnotac},{rhisfruta_entradas.fechaent}) in (select numnotac, fechaent from rfacttra_albaran) ") Then Exit Sub
        
        
        cadNombreRPT = "rResumFacturasTrans.rpt"
        
        cadTitulo = "Resumen de Facturas de Transporte"
                
                
        '[Monica]07/11/2013: añadimos el poder seleccionar los transportistas a facturar
        Set frmMens4 = New frmMensajes
        frmMens4.OpcionMensaje = 54
        frmMens4.cadWHERE = Sql4
        frmMens4.Show vbModal
        Set frmMens4 = Nothing
                
                
                
        Set frmMens = New frmMensajes
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = Sql2
        frmMens.Show vbModal
        Set frmMens = Nothing
        
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInformeNew(nTabla, cadSelect, Tabla1, cadSelect1) Then
            b = FacturacionTransporte(nTabla, cadSelect, Tabla1, cadSelect1, txtcodigo(15).Text, Me.Pb1, txtcodigo(6).Text, txtcodigo(7).Text)
            If b Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                               
                'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                If Me.Check1(2).Value Then
                    cadFormula = ""
                    CadParam = CadParam & "pFecFac= """ & txtcodigo(15).Text & """|"
                    numParam = numParam + 1
                    CadParam = CadParam & "pTitulo= ""Resumen Facturación Transporte""|"
                    numParam = numParam + 1
                    
                    FecFac = CDate(txtcodigo(15).Text)
                    cadAux = "{rfacttra.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                    ConSubInforme = False
                    
                    LlamarImprimir
                End If
                'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE TRANSPORTE
                If Me.Check1(3).Value Then
                    cadFormula = ""
                    cadSelect = ""
                    cadAux = "({stipom.codtipom} = 'FTR')"
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                    'Nº Factura
                    cadAux = "({rfacttra.numfactu} IN [" & ListaFacturasGeneradas("FTR") & "])"
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                     
                    'Fecha de Factura
                    FecFac = CDate(txtcodigo(15).Text)
                    cadAux = "{rfacttra.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                    cadAux = "{rfacttra.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                   
                    indRPT = 49 'Impresion de facturas de transportistas
                    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                    'Nombre fichero .rpt a Imprimir
                    cadNombreRPT = nomDocu
                    'Nombre fichero .rpt a Imprimir
                    cadTitulo = "Reimpresión de Facturas Transporte"
                    ConSubInforme = False
                    
                    conSubRPT = ConSubInforme
                    
                    LlamarImprimir
                    
                    If frmVisReport.EstaImpreso Then
                        ActualizarRegistrosFac "rfacttra", cadSelect
                    End If
                End If
                'SALIR DE LA FACTURACION DE tranportistas
                cmdCancel_Click (0)
            End If
        Else
            MsgBox "No hay entradas a facturar.", vbExclamation
        End If
    End If
    Label2(10).visible = False
    Label2(12).visible = False


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



Private Sub cmdAceptarReimp_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String


InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    Tipos = "{rfacttra.codtipom} = 'FTR' "
    If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
    
    'D/H Transportista
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codtrans}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTranspor= """) Then Exit Sub
    End If
    
    'D/H Nro de Factura
    cDesde = Trim(txtcodigo(4).Text)
    cHasta = Trim(txtcodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfacttra.numfactu}"
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
        indRPT = 49 'Impresion de Factura transporte
        ConSubInforme = False
        cadTitulo = "Reimpresión de Facturas transporte"
        
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
          
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
          
          
        'Nombre fichero .rpt a Imprimir
        
        LlamarImprimir
        
        If frmVisReport.EstaImpreso Then
            ActualizarRegistros "rfacttra", cadSelect
        End If
    End If

End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub cmdCancelReimp_Click()
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
Dim i As Integer
Dim indFrame As Single


    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    Me.FrameFacturar.visible = False
    Me.FrameReimpresion.visible = False
    
    ConexionConta
    
    For i = 0 To 5
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 12 To 13
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 20 To 21
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i
    

    ' Necesitamos la conexion a la contabilidad de la seccion de adv
    ' para sacar los porcentajes de iva de los articulos y calcular
    ' los datos de la factura
    
    
    NomTabla = "rhisfruta"
    NomTablaLin = "rhisfruta_entradas"
        
'    OpcionListado = 52
    Me.FrameFacturar.visible = False
    Me.FrameFacturarSocio.visible = False
    Me.FrameReimpresion.visible = False
    
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
        Case 1 ' Facturacion de Albaranes de transporte
            PonerFrameFacVisible True, H, W
            txtcodigo(15).Text = Format(Now, "dd/mm/yyyy")
            txtcodigo(7).Text = Format(CDate(txtcodigo(15).Text) - 1, "dd/mm/yyyy")
            indFrame = 6
            
            Me.Pb1.visible = False
            Me.Check1(2).Value = 1
            Me.Check1(3).Value = 1
            
        Case 2 ' Reimpresion de facturas de transporte
            FrameReimpresionVisible True, H, W
            Tabla = "rfacttra"
            
        Case 3 ' Facturacion de Albaranes de transporte a socios
            PonerFrameFacTraSocVisible True, H, W
            txtcodigo(17).Text = Format(Now, "dd/mm/yyyy")
            txtcodigo(9).Text = Format(CDate(txtcodigo(17).Text) - 1, "dd/mm/yyyy")
            indFrame = 6
            
            Me.pb2.visible = False
            Me.Check1(0).Value = 1
            Me.Check1(1).Value = 1
            
            Me.Frame1.Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
            Me.Frame1.visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
           
    End Select
    
    '[Monica]23/11/2015: solo en el caso de sea la facturacion de albaranes de trasporte a socios
    Me.Check1(4).visible = (OpcionListado = 3)
    Me.Check1(4).Value = 0
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(0).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Sub


Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
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
    If Not AnyadirAFormula(cadSelect1, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub


Private Sub frmMens4_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {rtransporte.codtrans} in (" & CadenaSeleccion & ")"
        Sql2 = " {rtransporte.codtrans} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {rtransporte.codtrans} = -1 "
    End If
    
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadSelect1, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

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
'            Select Case Index
'                Case 11, 12: indCodigo = Index + 9
'                Case 14, 15: indCodigo = Index + 14
'                Case 20, 21: indCodigo = Index + 20
'                Case 27, 28: indCodigo = Index + 21
'                Case 32: indCodigo = 8
'            End Select
'            Set frmSoc = New frmManSocios
'            frmSoc.DatosADevolverBusqueda = "0|2|"
'            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
'            frmSoc.Show vbModal
'            Set frmSoc = Nothing
            
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
    
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
    
    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmF.Left = esq + imgFec(Index).Parent.Left + 30
    frmF.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
   
   frmF.NovaData = Now
   
   Select Case Index
        Case 0 'FramePreFacturar
            indCodigo = 6
        Case 1 'FramePreFacturar
            indCodigo = 7
        Case 15 'Frame Factura
            indCodigo = 15
   End Select
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)


'++
'
'
'
'
'
'   Screen.MousePointer = vbHourglass
'   Set frmF = New frmCal
'   frmF.Fecha = Now
'
'
'   Select Case Index
'        Case 10 'FramePreFacturar
'            indCodigo = 26
'        Case 11 'FramePreFacturar
'            indCodigo = 27
'        Case 12 'Frame Factura
'            indCodigo = 38
'        Case 13 'Frame Factura
'            indCodigo = 39
'        Case 14 'FrameFactura
'            indCodigo = 34
'   End Select
'
'   PonerFormatoFecha txtCodigo(indCodigo)
'   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)
'
'   Screen.MousePointer = vbDefault
'   frmF.Show vbModal
'   Set frmF = Nothing
'   PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub OptTipoInf_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptDetalle_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si está marcado se liquidan los socios que sean terceros con contador " & vbCrLf & _
                      "diferente. " & vbCrLf & vbCrLf & _
                      "En caso contrario, sólo se liquidan los socios que no sean terceros." & vbCrLf & vbCrLf
    
        Case 1
           ' "____________________________________________________________"
            vCadena = "Si está marcado se liquidan los albaranes que esten en una factura  " & vbCrLf & _
                      "rectificada. " & vbCrLf & vbCrLf & _
                      "En caso contrario, sólo se liquidan los que no estén facturados." & vbCrLf & vbCrLf
    
    
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    

End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 18, 19, 20, 21, 28, 29 'Clases
            AbrirFrmClase (Index)
        
        Case 4 ' clases
            AbrirFrmClase (Index + 10)
        Case 5 ' clases
            AbrirFrmClase (Index + 11)
        
        Case 0, 1, 12, 13, 16, 17, 24, 25 'transportistas
            AbrirFrmTransportistas (Index)
        
        Case 2, 3 ' socios
            AbrirFrmSocios (Index + 8)
        
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
            Case 0: KEYBusqueda KeyAscii, 0 'transportista desde
            Case 1: KEYBusqueda KeyAscii, 1 'transportista hasta
            Case 12: KEYBusqueda KeyAscii, 12 'transportista desde
            Case 13: KEYBusqueda KeyAscii, 13 'transportista hasta
            Case 20: KEYBusqueda KeyAscii, 20 'clase desde
            Case 21: KEYBusqueda KeyAscii, 21 'clase hasta
            
            Case 2: KEYFecha KeyAscii, 3 'fecha desde
            Case 3: KEYFecha KeyAscii, 4 'fecha hasta
            Case 6: KEYFecha KeyAscii, 0 'fecha desde
            Case 7: KEYFecha KeyAscii, 1 'fecha hasta
            Case 15: KEYFecha KeyAscii, 2 'fecha hasta
            
            ' facturas de transporte a socios
            Case 10: KEYBusqueda KeyAscii, 2 'socio desde
            Case 11: KEYBusqueda KeyAscii, 3 'socio hasta
            Case 14: KEYBusqueda KeyAscii, 4 'clase desde
            Case 16: KEYBusqueda KeyAscii, 5 'clase hasta
            Case 8: KEYFecha KeyAscii, 6 'fecha desde
            Case 9: KEYFecha KeyAscii, 7 'fecha hasta
            Case 17: KEYFecha KeyAscii, 5 'fecha factura
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
Dim devuelve As String
Dim codcampo As String, nomCampo As String
Dim Tabla As String
      
    Select Case Index
        Case 0, 1, 12, 13 'transportistas
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rtransporte", "nomtrans", "codtrans", "T")
    
    
        'FECHA Desde Hasta
        Case 2, 3, 6, 7, 15, 8, 9, 17
            If txtcodigo(Index).Text <> "" Then
                PonerFormatoFecha txtcodigo(Index)
            End If
        
        Case 4, 5
            PonerFormatoEntero txtcodigo(Index)
        
        
        Case 36, 37  'Nº de Parte
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000000")
            End If
            
        Case 20, 21, 14, 16
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
        
        
        Case 40, 41, 10, 11 'Cod. Socio
            If PonerFormatoEntero(txtcodigo(Index)) Then
                nomCampo = "nomsocio"
                Tabla = "rsocios"
                codcampo = "codsocio"
                txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), Tabla, nomCampo, codcampo, "N")
                If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 42  'Cod. Formas de PAGO de comercial
            If PonerFormatoEntero(txtcodigo(Index)) Then
'                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "forpago", "nomforpa", "codforpa", "N")
'[Monica] 09/02/2010 no es de comercial sino de la contabilidad de adv
                txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(Index), "N")
                If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
        
    End Select
End Sub



Private Sub PonerFrameFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim Cad As String

    H = 6015
    W = 6735
    
    If visible = True Then
         Select Case CodClien 'aqui guardamos el tipo de movimiento
            Case "FTR": Cad = "(TRA)"
                
        End Select
        
        Me.Label2(10).Caption = "Factura de Transporte " & Cad
        Me.Label2(12).Caption = ""
        
        Me.Caption = "Facturación"
    End If
    
    PonerFrameVisible Me.FrameFacturar, visible, H, W
End Sub


Private Sub PonerFrameFacTraSocVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim Cad As String

    H = 6015
    W = 6735
    
    If visible = True Then
         Select Case CodClien 'aqui guardamos el tipo de movimiento
            Case "FTS": Cad = "(TRA)"
                
        End Select
        
        Me.Label2(13).Caption = "Factura de Transporte a Socio" & Cad
        Me.Label2(14).Caption = ""
        
        Me.Caption = "Facturación"
    End If
    
    PonerFrameVisible Me.FrameFacturarSocio, visible, H, W
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


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadSelect1 = ""
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
        .Opcion = 0
        .Titulo = cadTitulo
        .ConSubInforme = conSubRPT
        .NombreRPT = cadNombreRPT  'nombre del informe
        .Show vbModal
    End With
End Sub

Private Sub txtCodigo_Validate(Index As Integer, Cancel As Boolean)
'    Select Case Index
'           Case 15, 16 'ARTICULO
'            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "sartic", "nomartic", "codartic", "Articulo", "T")
'            If txtNombre(Index).Text = "" And txtcodigo(Index) <> "" Then Cancel = True
'    End Select
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




Private Function ActualizarRegistrosFac(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim Sql As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistrosFac = False
    Sql = "update " & cTabla & ", usuarios.stipom set impreso = 1 "
    Sql = Sql & " where usuarios.stipom.codtipom = rfacttra.codtipom "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " and " & cWhere
    End If
    
    conn.Execute Sql
    
    ActualizarRegistrosFac = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function

Private Function ActualizarRegistrosFacSoc(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim Sql As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistrosFacSoc = False
    Sql = "update " & cTabla & ", usuarios.stipom set impreso = 1 "
    Sql = Sql & " where usuarios.stipom.codtipom = rfactsoc.codtipom "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " and " & cWhere
    End If
    
    conn.Execute Sql
    
    ActualizarRegistrosFacSoc = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
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


Private Sub AbrirFrmClase(indice As Integer)
    indCodigo = indice
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtcodigo(indice).Text
    
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmTransportistas(indice As Integer)
    indCodigo = indice
    Set frmTra = New frmManTranspor
    frmTra.DatosADevolverBusqueda = "0|1|"
    frmTra.Show vbModal
    Set frmTra = Nothing
End Sub

Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub



Private Sub FrameReimpresionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameReimpresion.visible = visible
    If visible = True Then
        Me.FrameReimpresion.Top = -90
        Me.FrameReimpresion.Left = 0
        Me.FrameReimpresion.Height = 5640
        Me.FrameReimpresion.Width = 6675
        W = Me.FrameReimpresion.Width
        H = Me.FrameReimpresion.Height
    End If
End Sub



