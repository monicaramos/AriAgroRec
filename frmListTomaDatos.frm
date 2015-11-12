VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListTomaDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7290
   Icon            =   "frmListTomaDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameClasificaSocio 
      Height          =   5010
      Left            =   30
      TabIndex        =   90
      Top             =   60
      Width           =   6285
      Begin VB.CommandButton CmdCancelClas 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4770
         TabIndex        =   112
         Top             =   4170
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarClas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   110
         Top             =   4170
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   102
         Top             =   1560
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   101
         Top             =   1200
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   100
         Text            =   "Text5"
         Top             =   1215
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   99
         Text            =   "Text5"
         Top             =   1575
         Width           =   3375
      End
      Begin VB.CommandButton Command6 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTomaDatos.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command5 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTomaDatos.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   103
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1605
         MaxLength       =   6
         TabIndex        =   104
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   96
         Text            =   "Text5"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   95
         Text            =   "Text5"
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   1035
         Left            =   180
         TabIndex        =   91
         Top             =   2820
         Width           =   2565
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   31
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   108
            Top             =   675
            Width           =   1095
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   1395
            MaxLength       =   10
            TabIndex        =   106
            Top             =   270
            Width           =   1095
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   1080
            Picture         =   "frmListTomaDatos.frx":0620
            ToolTipText     =   "Buscar fecha"
            Top             =   690
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   4
            Left            =   1095
            Picture         =   "frmListTomaDatos.frx":06AB
            ToolTipText     =   "Buscar fecha"
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   94
            Top             =   675
            Width           =   420
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   93
            Top             =   330
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   92
            Top             =   90
            Width           =   450
         End
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   2
         Left            =   330
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   3900
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   24
         Left            =   630
         TabIndex        =   115
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   17
         Left            =   630
         TabIndex        =   114
         Top             =   1620
         Width           =   420
      End
      Begin VB.Label Label5 
         Caption         =   "Informe de Clasificación por Socio"
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
         TabIndex        =   113
         Top             =   330
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   16
         Left            =   300
         TabIndex        =   111
         Top             =   1020
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1290
         MouseIcon       =   "frmListTomaDatos.frx":0736
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1290
         MouseIcon       =   "frmListTomaDatos.frx":0888
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1605
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   15
         Left            =   300
         TabIndex        =   109
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   675
         TabIndex        =   107
         Top             =   2205
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   675
         TabIndex        =   105
         Top             =   2595
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1290
         MouseIcon       =   "frmListTomaDatos.frx":09DA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1290
         MouseIcon       =   "frmListTomaDatos.frx":0B2C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2550
         Width           =   240
      End
   End
   Begin VB.Frame FrameTomaDatos 
      Height          =   7320
      Left            =   30
      TabIndex        =   13
      Top             =   0
      Width           =   6285
      Begin VB.CheckBox Check3 
         Caption         =   "Informe para Seguro"
         Height          =   225
         Left            =   3810
         TabIndex        =   122
         Top             =   6000
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Nro.Campo Asignado"
         Height          =   225
         Left            =   3810
         TabIndex        =   121
         Top             =   5580
         Width           =   1815
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   3
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   8
         Top             =   5025
         Width           =   4200
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Recolectado|N|N|0|1|rcampos|recolect||N|"
         Top             =   2160
         Width           =   1680
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Recolectado|N|N|0|1|rcampos|recolect||N|"
         Top             =   2580
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo Superficie"
         ForeColor       =   &H00972E0B&
         Height          =   1455
         Left            =   2040
         TabIndex        =   61
         Top             =   630
         Width           =   1815
         Begin VB.OptionButton Option3 
            Caption         =   "Cultivable"
            Height          =   225
            Index           =   3
            Left            =   300
            TabIndex        =   119
            Top             =   1110
            Width           =   1305
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Cooperativa"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   64
            Top             =   300
            Width           =   1305
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Sigpac"
            Height          =   225
            Index           =   1
            Left            =   300
            TabIndex        =   63
            Top             =   570
            Width           =   1305
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Catastro"
            Height          =   225
            Index           =   2
            Left            =   300
            TabIndex        =   62
            Top             =   840
            Width           =   1035
         End
      End
      Begin VB.Frame FrameFecha 
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   150
         TabIndex        =   57
         Top             =   5430
         Width           =   2565
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   1485
            MaxLength       =   10
            TabIndex        =   9
            Top             =   210
            Width           =   1095
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   1500
            MaxLength       =   10
            TabIndex        =   10
            Top             =   615
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   19
            Left            =   210
            TabIndex        =   60
            Top             =   30
            Width           =   450
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   20
            Left            =   540
            TabIndex        =   59
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   21
            Left            =   540
            TabIndex        =   58
            Top             =   615
            Width           =   420
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1185
            Picture         =   "frmListTomaDatos.frx":0C7E
            ToolTipText     =   "Buscar fecha"
            Top             =   615
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1185
            Picture         =   "frmListTomaDatos.frx":0D09
            ToolTipText     =   "Buscar fecha"
            Top             =   210
            Width           =   240
         End
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   390
         TabIndex        =   56
         Top             =   6510
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo Listado"
         ForeColor       =   &H00972E0B&
         Height          =   1455
         Left            =   300
         TabIndex        =   53
         Top             =   630
         Width           =   1575
         Begin VB.OptionButton Option2 
            Caption         =   "Socio"
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   390
            Width           =   1125
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Partida"
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   54
            Top             =   810
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Imprime"
         ForeColor       =   &H00972E0B&
         Height          =   1455
         Left            =   4020
         TabIndex        =   49
         Top             =   630
         Width           =   2085
         Begin VB.OptionButton Option1 
            Caption         =   "Producción Real"
            Height          =   225
            Index           =   2
            Left            =   330
            TabIndex        =   52
            Top             =   990
            Width           =   1485
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Con Aforo"
            Height          =   225
            Index           =   1
            Left            =   330
            TabIndex        =   51
            Top             =   660
            Width           =   1305
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Sin Aforo"
            Height          =   225
            Index           =   0
            Left            =   330
            TabIndex        =   50
            Top             =   330
            Width           =   1305
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2550
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   0
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   4470
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   4110
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1665
         MaxLength       =   3
         TabIndex        =   7
         Top             =   4470
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1665
         MaxLength       =   3
         TabIndex        =   6
         Top             =   4110
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTomaDatos.frx":0D94
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTomaDatos.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   16
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
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   3495
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   3135
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3495
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   4
         Top             =   3135
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptarTom 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3750
         TabIndex        =   11
         Top             =   6840
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelTom 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4830
         TabIndex        =   12
         Top             =   6840
         Width           =   975
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   1
         Left            =   5760
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   5970
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   5760
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   5580
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Título del informe"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   26
         Left            =   330
         TabIndex        =   118
         Top             =   5040
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   4
         Left            =   3450
         TabIndex        =   117
         Top             =   2190
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Recolectado"
         Enabled         =   0   'False
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   3450
         TabIndex        =   116
         Top             =   2610
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Presentación"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   48
         Top             =   2550
         Width           =   1425
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1920
         Picture         =   "frmListTomaDatos.frx":13A8
         ToolTipText     =   "Buscar fecha"
         Top             =   2550
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Entrega"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   25
         Left            =   360
         TabIndex        =   47
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1920
         Picture         =   "frmListTomaDatos.frx":1433
         ToolTipText     =   "Buscar fecha"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1350
         MouseIcon       =   "frmListTomaDatos.frx":14BE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1350
         MouseIcon       =   "frmListTomaDatos.frx":1610
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   4110
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   735
         TabIndex        =   26
         Top             =   4545
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   735
         TabIndex        =   25
         Top             =   4155
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   24
         Top             =   3900
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1350
         MouseIcon       =   "frmListTomaDatos.frx":1762
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3525
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1350
         MouseIcon       =   "frmListTomaDatos.frx":18B4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3135
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   360
         TabIndex        =   21
         Top             =   2940
         Width           =   405
      End
      Begin VB.Label Label3 
         Caption         =   "Informe de Toma de Datos"
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
         TabIndex        =   20
         Top             =   240
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   690
         TabIndex        =   19
         Top             =   3540
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   690
         TabIndex        =   18
         Top             =   3180
         Width           =   465
      End
   End
   Begin VB.Frame FrameDesviacionAforos 
      Height          =   5220
      Left            =   60
      TabIndex        =   65
      Top             =   30
      Width           =   6285
      Begin VB.CheckBox Check2 
         Caption         =   "Salta página por Socio"
         Height          =   255
         Left            =   630
         TabIndex        =   72
         Top             =   3930
         Width           =   2265
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo Hanegadas"
         ForeColor       =   &H00972E0B&
         Height          =   885
         Left            =   330
         TabIndex        =   80
         Top             =   2790
         Width           =   5475
         Begin VB.OptionButton Option4 
            Caption         =   "Cultivable"
            Height          =   285
            Index           =   3
            Left            =   3960
            TabIndex        =   120
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Cooperativa"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   71
            Top             =   390
            Width           =   1305
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Sigpac"
            Height          =   225
            Index           =   1
            Left            =   1680
            TabIndex        =   82
            Top             =   390
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Catastro"
            Height          =   225
            Index           =   2
            Left            =   2820
            TabIndex        =   81
            Top             =   390
            Width           =   945
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   79
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
         TabIndex        =   78
         Text            =   "Text5"
         Top             =   2430
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1635
         MaxLength       =   3
         TabIndex        =   69
         Top             =   2070
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   70
         Top             =   2430
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTomaDatos.frx":1A06
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTomaDatos.frx":1D10
         Style           =   1  'Graphical
         TabIndex        =   75
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
         TabIndex        =   73
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
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   1470
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   66
         Top             =   1110
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   68
         Top             =   1470
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptarDesv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3780
         TabIndex        =   74
         Top             =   4605
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelDesv 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4860
         TabIndex        =   76
         Top             =   4605
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1350
         MouseIcon       =   "frmListTomaDatos.frx":201A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1350
         MouseIcon       =   "frmListTomaDatos.frx":216C
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
         TabIndex        =   89
         Top             =   2505
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   705
         TabIndex        =   88
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
         TabIndex        =   87
         Top             =   1860
         Width           =   390
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1320
         MouseIcon       =   "frmListTomaDatos.frx":22BE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1320
         MouseIcon       =   "frmListTomaDatos.frx":2410
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
         TabIndex        =   86
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
         TabIndex        =   85
         Top             =   330
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   9
         Left            =   660
         TabIndex        =   84
         Top             =   1530
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   12
         Left            =   660
         TabIndex        =   83
         Top             =   1170
         Width           =   465
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
      Height          =   4230
      Left            =   60
      TabIndex        =   27
      Top             =   0
      Width           =   6645
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   33
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
         TabIndex        =   32
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2055
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   34
         Tag             =   "Código Postal|T|S|||clientes|codposta|##0.00||"
         Top             =   2970
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancelResul 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5340
         TabIndex        =   36
         Top             =   3525
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepResul 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   35
         Top             =   3525
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   31
         Top             =   1485
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
         Text            =   "Text5"
         Top             =   1110
         Width           =   3675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   29
         Left            =   1440
         MouseIcon       =   "frmListTomaDatos.frx":2562
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2460
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1440
         MouseIcon       =   "frmListTomaDatos.frx":26B4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Incrementar / Decrementar Aforo"
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
         TabIndex        =   44
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   795
         TabIndex        =   43
         Top             =   2445
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   795
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   1830
         Width           =   390
      End
      Begin VB.Label Label4 
         Caption         =   "Porcentaje"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   23
         Left            =   435
         TabIndex        =   40
         Top             =   2865
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   795
         TabIndex        =   39
         Top             =   1155
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   19
         Left            =   795
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   915
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1440
         MouseIcon       =   "frmListTomaDatos.frx":2806
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1440
         MouseIcon       =   "frmListTomaDatos.frx":2958
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1110
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListTomaDatos"
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
    ' 1 .- Informe de Toma de Datos
    ' 2 .- Informe de Desviación de Aforos
    ' 3 .- Informe de Clasificación Socio
    
    ' 4 .- Proceso de incremento o decremento de porcentaje de aforo sobre campos
    
    
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
Private WithEvents frmCla As frmComercial 'Ayuda Clases de comercial
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmPro As frmComercial 'Ayuda Productos de comercial
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMens1 As frmMensajes 'Mensajes
Attribute frmMens1.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadparam As String 'Cadena con los parametros para Crystal Report
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
Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadparam = cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    'D/H Socios
    cDesde = Trim(txtCodigo(24).Text)
    cHasta = Trim(txtCodigo(25).Text)
    nDesde = txtNombre(24).Text
    nHasta = txtNombre(25).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Clase
    cDesde = Trim(txtCodigo(28).Text)
    cHasta = Trim(txtCodigo(29).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
    End If
    
        
    nTabla = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
    
    If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
    
    vSQL = ""
    If txtCodigo(28).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(28).Text, "N")
    If txtCodigo(29).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(29).Text, "N")
    
    Set frmMens = New frmMensajes
    
    frmMens.OpcionMensaje = 16
    frmMens.cadwhere = vSQL
    frmMens.Show vbModal
    
    Set frmMens = Nothing
    
    If HayRegistros(nTabla, cadSelect) Then
        ProcesarCambiosAforos nTabla, cadSelect
        CmdCancelResul_Click
    End If

End Sub



Private Sub CmdAceptarClas_Click()
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

Dim Nregs As Long
Dim FecFac As Date

Dim b As Boolean
Dim TipoPrec As Byte

Dim vSQL As String

    vSQL = ""

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadparam = cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtCodigo(16).Text)
        cHasta = Trim(txtCodigo(17).Text)
        nDesde = txtNombre(16).Text
        nHasta = txtNombre(17).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        If txtCodigo(16).Text <> "" Then vSQL = vSQL & " and rcampos.codsocio >= " & DBSet(txtCodigo(16).Text, "N")
        If txtCodigo(17).Text <> "" Then vSQL = vSQL & " and rcampos.codsocio <= " & DBSet(txtCodigo(17).Text, "N")
        
        
        'D/H VARIEDAD
        cDesde = Trim(txtCodigo(18).Text)
        cHasta = Trim(txtCodigo(19).Text)
        nDesde = txtNombre(18).Text
        nHasta = txtNombre(19).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rcampos.codvarie}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
        End If
        
        If txtCodigo(18).Text <> "" Then vSQL = vSQL & " and rcampos.codvarie >= " & DBSet(txtCodigo(18).Text, "N")
        If txtCodigo(19).Text <> "" Then vSQL = vSQL & " and rcampos.codvarie <= " & DBSet(txtCodigo(19).Text, "N")
        
        
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        
        'CAMPOS DADOS DE ALTA
        If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null ") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub
        
        vSQL = vSQL & " and rcampos.fecbajas is null"
        
        
        nTabla = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rcampos.codsocio = rsocios_seccion.codsocio "


        cadNombreRPT = "rInfClasSocios.rpt"
        cadTitulo = "Informe de Clasificación por Socio"
             
        Set frmMens1 = New frmMensajes
        
        frmMens1.OpcionMensaje = 15
        frmMens1.cadwhere = vSQL
        frmMens1.Show vbModal
        
        Set frmMens1 = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
            ConSubInforme = False
            LlamarImprimir
        End If
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

Dim Nregs As Long
Dim FecFac As Date

Dim b As Boolean
Dim TipoPrec As Byte

Dim vSQL As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadparam = cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
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
        If Option4(0).Value Then cadparam = cadparam & "pTipoHa=0|"
        If Option4(1).Value Then cadparam = cadparam & "pTipoHa=1|"
        If Option4(2).Value Then cadparam = cadparam & "pTipoHa=2|"
        If Option4(3).Value Then cadparam = cadparam & "pTipoHa=3|"
        numParam = numParam + 1
             
        If Check2.Value Then
            cadparam = cadparam & "pSaltoSocio=1|"
        Else
            cadparam = cadparam & "pSaltoSocio=0|"
        End If
        numParam = numParam + 1
             
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadwhere = vSQL
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
            ConSubInforme = True
            LlamarImprimir
        End If
    End If


End Sub

Private Sub cmdAceptarTom_Click()
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

Dim Nregs As Long
Dim FecFac As Date

Dim b As Boolean
Dim TipoPrec As Byte

Dim vSQL As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadparam = cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtCodigo(12).Text)
        cHasta = Trim(txtCodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H PRODUCTO
        cDesde = Trim(txtCodigo(20).Text)
        cHasta = Trim(txtCodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codprodu}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto=""") Then Exit Sub
        End If
        
        
        If Option1(2).Value Or Option2(1).Value Then
            'D/H fecha
            cDesde = Trim(txtCodigo(6).Text)
            cHasta = Trim(txtCodigo(7).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rhisfruta.fecalbar}"
                TipCod = "F"
                
                devuelve = CadenaDesdeHasta(cDesde, cHasta, Codigo, TipCod)
                If devuelve = "Error" Then Exit Sub
                
'                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
        End If
        
        
'[Monica]19/12/2011: quito el control de la seccion de horto pq en Quatretonda tienen campos de horto y de almazara
'                    hacen entradas de almazara por entradas de bascula
'        'SECCION
'        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
'        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        'CAMPOS DADOS DE ALTA
        If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null ") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas}) ") Then Exit Sub
        
        nTabla = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
'[Monica]01/03/2012: por quitar lo de la seccion horto
'       nTabla = nTabla & " INNER JOIN rsocios_seccion ON rcampos.codsocio = rsocios_seccion.codsocio "
        nTabla = nTabla & " INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
        

        Select Case OpcionListado
            Case 1 ' Listado de toma de datos
                'Nombre fichero .rpt a Imprimir
               
                cadparam = cadparam & "pTipo=" & Combo1(0).ListIndex & "|"
                numParam = numParam + 1
                
                cadparam = cadparam & "pTitulo=""" & txtCodigo(3).Text & """|"
                numParam = numParam + 1
                
                'tipo de listado
                If Option2(0).Value Then
                    indRPT = 71 ' informe de toma de datos por socio
                    If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
                    cadNombreRPT = nomDocu ' "rInfTomaDatos.rpt"
                    
                    'tipo de listado
                    If Option1(0).Value Then cadparam = cadparam & "pTipoListado=0|"
                    If Option1(1).Value Then cadparam = cadparam & "pTipoListado=1|"
                    If Option1(2).Value Then cadparam = cadparam & "pTipoListado=2|"
                    numParam = numParam + 1
                    
                    cadparam = cadparam & "pFecentre=""" & txtCodigo(15).Text & """|"
                    numParam = numParam + 1
                    cadparam = cadparam & "pFecprese=""" & txtCodigo(2).Text & """|"
                    numParam = numParam + 1
                    
                    '[Monica]01/04/2011: indicamos si es o no seguro
                    If Check3.Value = 1 Then
                        cadparam = cadparam & "pSeguro=1|"
                    Else
                        cadparam = cadparam & "pSeguro=0|"
                    End If
                    numParam = numParam + 1
                End If
                
                If Option2(1).Value Then
                    cadNombreRPT = "rInfTomaDatos1.rpt"
                    
                    ' campo recolectado por socio cooperativa o ambos
                    Select Case Combo1(1).ListIndex
                        Case 0 ' cooperativa
                            If Not AnyadirAFormula(cadSelect, "{rcampos.recolect}=0") Then Exit Sub
                            If Not AnyadirAFormula(cadFormula, "{rcampos.recolect}=0") Then Exit Sub
                        
                        Case 1 ' socio
                            If Not AnyadirAFormula(cadSelect, "{rcampos.recolect}=1") Then Exit Sub
                            If Not AnyadirAFormula(cadFormula, "{rcampos.recolect}=1") Then Exit Sub
                            
                        Case 2 ' ambos
                    
                    End Select
                    
                    cadparam = cadparam & "pRecolect=" & Combo1(1).ListIndex & "|"
                    numParam = numParam + 1
                
                End If
                
                cadTitulo = "Informe de Toma de Datos"
                
                If Check1.Value = 1 Then
                    If Not AnyadirAFormula(cadSelect, "{rcampos.nrocampo}<>0") Then Exit Sub
                    If Not AnyadirAFormula(cadFormula, "{rcampos.nrocampo}<>0") Then Exit Sub
                End If
        End Select
                    
        vSQL = ""
        If txtCodigo(20).Text <> "" Then vSQL = vSQL & " and variedades.codprodu >= " & DBSet(txtCodigo(20).Text, "N")
        If txtCodigo(21).Text <> "" Then vSQL = vSQL & " and variedades.codprodu <= " & DBSet(txtCodigo(21).Text, "N")
        
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadwhere = vSQL
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
            If OpcionListado = 1 Then
                If CargarTemporalProdReal(nTabla, cadSelect) Then
                    '[Monica]15/07/2011: tenemos que cargar las parcelas para picassent
                    ' si no es seguros tenemos que cargar el anexo con los registros de rcampos_parcelas
                    If Check3.Value = 0 And vParamAplic.Cooperativa = 2 Then
                        If Not CargarSubparcelas(nTabla, cadSelect) Then Exit Sub
                    End If
                    
                    'tipo de hanegada
                    If Option3(0).Value Then cadparam = cadparam & "pTipoHa=0|"
                    If Option3(1).Value Then cadparam = cadparam & "pTipoHa=1|"
                    If Option3(2).Value Then cadparam = cadparam & "pTipoHa=2|"
                    If Option3(3).Value Then cadparam = cadparam & "pTipoHa=3|"
                    
                    numParam = numParam + 1
                    
                    cadFormula = ""
                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                    ConSubInforme = True
                    
                    LlamarImprimir
                End If
            End If
        End If
    End If

End Sub

Private Function CargarSubparcelas(nTabla, cadSelect) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset

    On Error GoTo eCargarSubparcelas


    CargarSubparcelas = False

    Sql = "delete from tmpinfkilos where codusu = " & vUsu.Codigo
    conn.Execute Sql
                                                                'poligono,parcela
    Sql = "insert into tmpinfkilos (codusu, codsocio, codcampo, codprodu, kilosnet) select " & vUsu.Codigo & ", importe1,"
    Sql = Sql & " importe2, poligono, parcela from  rcampos_parcelas, tmpinformes where tmpinformes.codusu = " & vUsu.Codigo & " and  "
    Sql = Sql & " rcampos_parcelas.codcampo = tmpinformes.importe2 "

    conn.Execute Sql

    'borramos los campos que tienen solo una parcela ( porque se ponen en la linea pppal del listado no en el subreport de anexos )
    Sql = "delete from tmpinfkilos where (codusu, codcampo, codprodu, kilosnet) in (select " & vUsu.Codigo & ","
    Sql = Sql & " rcampos.codcampo, rcampos.poligono, rcampos.parcela from rcampos "
    Sql = Sql & " where rcampos.codcampo = tmpinfkilos.codcampo)"
     
    conn.Execute Sql
    CargarSubparcelas = True
    Exit Function
    
eCargarSubparcelas:
    MuestraError Err.Number, "Cargar subparcelas", Err.Description
End Function


Private Sub CmdCancelClas_Click()
    Unload Me
End Sub

Private Sub cmdCancelDesv_Click()
    Unload Me
End Sub

Private Sub cmdCancelTom_Click()
    Unload Me
End Sub


Private Sub CmdCancelResul_Click()
    Unload Me
End Sub




Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 1 ' 1-Toma de Datos
                txtCodigo(15).Text = Format(Now, "dd/mm/yyyy")
                txtCodigo(2).Text = Format(Now, "dd/mm/yyyy")
                
                Me.Option1(0).Value = True
                Me.Option2(0).Value = True
                Me.Option3(0).Value = True
                
                PonerFoco txtCodigo(15)
                FrameFecha.visible = False
                FrameFecha.Enabled = False
                Combo1(0).ListIndex = 1
                Combo1(1).ListIndex = 2
                    
            Case 2  ' 2-Desviacion de aforos
                Option4(0).Value = True
                Check2.Value = 1
                PonerFoco txtCodigo(9)
                
            Case 3  ' 3-Clasificacion por socio
                PonerFoco txtCodigo(16)
            
            Case 4  ' 4-Incrementar/Decrementar aforos
                PonerFoco txtCodigo(24)
            
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
    For h = 9 To 10
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
    
    For h = 0 To imgAyuda.Count - 1
        imgAyuda(h).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next h
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    FrameTomaDatos.visible = False
    FrameResultados.visible = False
    FrameDesviacionAforos.visible = False
    FrameClasificaSocio.visible = False
    
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 1   '1- Informe de Toma de Datos
        FrameTomaDatosVisible True, h, w
        Tabla = "rcampos"
        Me.Label3.Caption = "Informe de Toma de Datos"
        Me.pb1.visible = False
        CargaCombo
        txtCodigo(3).Text = "Listado de Kilos Estimados"
        
    Case 2   '2 - Listado de Desviación de Aforos
        FrameDesviacionAforosVisible True, h, w
        Tabla = "rcampos"
        Me.Label3.Caption = "Informe de Desviación de Aforos"
        Me.pb1.visible = False
        
    Case 3   '3 - Informe de Clasificación de Socios
        FrameClasificacionVisible True, h, w
        Tabla = "rcampos"
        Me.Label3.Caption = "Informe de Clasificación de Socios"
        
        
    Case 4   '4 - Incrementar decrementar aforo
        FrameResultadosVisible True, h, w
        Tabla = "rcampos"
        
    End Select
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
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


Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {rcampos.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {rcampos.codvarie} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {rcampos.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub


Private Sub frmMens1_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {rcampos.codcampo} in (" & CadenaSeleccion & ")"
        Sql2 = " {rcampos.codcampo} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {rcampos.codcampo} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
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

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Marcándolo únicamente saldrán los campos que tengan el Nro.Campo" & vbCrLf & _
                      "distinto de cero. Son los campos que tienen asignada ya parcela y  " & vbCrLf & _
                      "término." & vbCrLf & vbCrLf
                      
        Case 1
           ' "____________________________________________________________"
            vCadena = "Marcándolo si seleccionamos la producción real, se calcularán " & vbCrLf & _
                      "los kilos de la campaña actual. Además sacaremos las subparcelas " & vbCrLf & _
                      "en línea." & vbCrLf & vbCrLf & _
                      "Si no lo marcamos los kilos son de la campaña anterior y las " & vbCrLf & _
                      "subparcelas nos apareceran al final de cada socio. " & vbCrLf & vbCrLf & _
                      "Sólo tiene efecto para el tipo listado Socio." & vbCrLf & vbCrLf
                      
        Case 2
           ' "____________________________________________________________"
            vCadena = "Este informe enlaza con la clasificacion de campos, cuando la " & vbCrLf & _
                      "entrada de ese campo se clasificaba en él." & vbCrLf & _
                      "" & vbCrLf
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1, 28, 29   'Clases
            AbrirFrmClase (Index)
        
        Case 20, 21 'Productos
            AbrirFrmProducto (Index)
        
        Case 9, 10, 12, 13, 16, 17, 24, 25 'SOCIOS
            AbrirFrmSocios (Index)
            
        Case 18, 19 ' Variedades
            AbrirFrmVariedad (Index)
        
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
        Case 0
            indice = 6
        Case 1
            indice = 7
        Case 2
            indice = 15
        Case 3
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
        Case 4, 5
            indice = Index + 26
    End Select

    imgFec(0).Tag = indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(indice).Text <> "" Then frmC.NovaData = txtCodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFec(0).Tag)) '<===
    ' ********************************************

End Sub



Private Sub Option1_Click(Index As Integer)
Dim b As Boolean
'    b = (Option1(2).Value = True Or Option2(1).Value = True)
    b = (Option1(2).Value = True)
    FrameFecha.visible = b
    FrameFecha.Enabled = b
    If Not b Then ' limpiamos los campos de fechas
        txtCodigo(6).Text = ""
        txtCodigo(7).Text = ""
    End If
End Sub

Private Sub Option2_Click(Index As Integer)
Dim b As Boolean
'    b = (Option1(2).Value = True Or Option2(1).Value = True)
    b = (Option1(2).Value = True)
    FrameFecha.visible = b
    FrameFecha.Enabled = b
    If Not b Then ' limpiamos los campos de fechas
        txtCodigo(6).Text = ""
        txtCodigo(7).Text = ""
    End If
    
    ' recolectado solo si es partida
    b = (Option2(1).Value = True)
    Combo1(1).visible = b
    Combo1(1).Enabled = b
    Label2(3).visible = b
    Label2(3).Enabled = b
    
End Sub

Private Sub Option3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Option3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 9: KEYBusqueda KeyAscii, 9 'socio desde
            Case 10: KEYBusqueda KeyAscii, 10 'socio hasta
            Case 12: KEYBusqueda KeyAscii, 12 'socio desde
            Case 13: KEYBusqueda KeyAscii, 13 'socio hasta
            Case 16: KEYBusqueda KeyAscii, 16 'socio desde
            Case 17: KEYBusqueda KeyAscii, 17 'socio hasta
            Case 24: KEYBusqueda KeyAscii, 24 'socio desde
            Case 25: KEYBusqueda KeyAscii, 25 'socio hasta
            Case 0: KEYBusqueda KeyAscii, 0 'clase desde
            Case 1: KEYBusqueda KeyAscii, 1 'clase hasta
            Case 18: KEYBusqueda KeyAscii, 18 'variedad desde
            Case 19: KEYBusqueda KeyAscii, 19 'variedad hasta
            Case 20: KEYBusqueda KeyAscii, 20 'producto desde
            Case 21: KEYBusqueda KeyAscii, 21 'producto hasta
            Case 28: KEYBusqueda KeyAscii, 28 'clase desde
            Case 29: KEYBusqueda KeyAscii, 29 'clase hasta
            Case 27: KEYFecha KeyAscii, 10 'fecha hasta
            Case 6: KEYFecha KeyAscii, 0 'fecha desde
            Case 7: KEYFecha KeyAscii, 1 'fecha hasta
            Case 2: KEYFecha KeyAscii, 3 'fecha desde
            Case 3: KEYFecha KeyAscii, 4 'fecha hasta
            Case 30: KEYFecha KeyAscii, 4 'fecha desde
            Case 31: KEYFecha KeyAscii, 5 'fecha hasta
            
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
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
    
        Case 9, 10, 12, 13, 16, 17, 24, 25  'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
        Case 4, 5 ' NROS DE FACTURA
            PonerFormatoEntero txtCodigo(Index)
            
        Case 2, 6, 7, 11, 15, 30, 31   'FECHAS
            b = True
            If txtCodigo(Index).Text <> "" Then
                If Index = 6 Or Index = 7 Then
                    b = PonerFormatoFecha(txtCodigo(Index), True)
                Else
                    b = PonerFormatoFecha(txtCodigo(Index))
                End If
            End If
            
        Case 0, 1, 28, 29  'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 20, 21  'PRODUCTOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "productos", "nomprodu", "codprodu", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
        Case 18, 19 ' variedades
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 26 ' porcentaje de incremento/decremento de afor
            PonerFormatoDecimal txtCodigo(Index), 4
        
    End Select
End Sub

Private Sub FrameTomaDatosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameTomaDatos.visible = visible
    If visible = True Then
        Me.FrameTomaDatos.Top = -90
        Me.FrameTomaDatos.Left = 0
        Me.FrameTomaDatos.Height = 7320
        Me.FrameTomaDatos.Width = 6615
        w = Me.FrameTomaDatos.Width
        h = Me.FrameTomaDatos.Height
    End If
End Sub

Private Sub FrameDesviacionAforosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de diferencias de produccion
    Me.FrameDesviacionAforos.visible = visible
    If visible = True Then
        Me.FrameDesviacionAforos.Top = -90
        Me.FrameDesviacionAforos.Left = 0
        Me.FrameDesviacionAforos.Height = 5220
        Me.FrameDesviacionAforos.Width = 6285
        w = Me.FrameDesviacionAforos.Width
        h = Me.FrameDesviacionAforos.Height
    End If
End Sub

Private Sub FrameClasificacionVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de diferencias de produccion
    Me.FrameClasificaSocio.visible = visible
    If visible = True Then
        Me.FrameClasificaSocio.Top = -90
        Me.FrameClasificaSocio.Left = 0
        Me.FrameClasificaSocio.Height = 5010
        Me.FrameClasificaSocio.Width = 6285
        w = Me.FrameClasificaSocio.Width
        h = Me.FrameClasificaSocio.Height
    End If
End Sub



Private Sub FrameResultadosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameResultados.visible = visible
    If visible = True Then
        Me.FrameResultados.Top = -90
        Me.FrameResultados.Left = 0
        Me.FrameResultados.Height = 4230
        Me.FrameResultados.Width = 6645
        w = Me.FrameResultados.Width
        h = Me.FrameResultados.Height
    End If
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadparam = ""
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
            cadparam = cadparam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadparam
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
    frmCal.Show vbModal
    Set frmCal = Nothing
End Sub

Private Sub AbrirFrmSeccion(indice As Integer)
    indCodigo = indice
    Set frmSec = New frmManSeccion
    frmSec.DatosADevolverBusqueda = "0|1|"
    frmSec.Show vbModal
    Set frmSec = Nothing
End Sub

Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
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

Private Sub AbrirFrmProducto(indice As Integer)
    indCodigo = indice
    Set frmPro = New frmComercial
    
    AyudaProductosCom frmPro, txtCodigo(indice).Text
    
    Set frmPro = Nothing
End Sub

Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
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
Dim TipoMov As String

    '[Monica] 24/06/2010
    If vParamAplic.Seccionhorto = "" Then
        MsgBox "No tiene asignada la sección de Horto en parámetros. Revise.", vbExclamation
        DatosOk = False
        Exit Function
    End If


    b = True
    Select Case OpcionListado
        Case 1
            '1 - Informe de Toma de Datos
            If b Then
                If Option1(2).Value Then ' solo si estamos mirando la produccion real
                    If txtCodigo(6).Text = "" Or txtCodigo(7) = "" Then
                        MsgBox "Debe introducir obligatoriamente el rango de fechas.", vbExclamation
                        b = False
                        PonerFoco txtCodigo(6)
                    End If
                 End If
            End If
        Case 4
            ' incrementar / decrementar porcentaje de aforo
            If txtCodigo(26).Text = "" Then
                MsgBox "Debe introducir obligatoriamente un porcentaje de incremento/decremento", vbExclamation
                b = False
                PonerFoco txtCodigo(26)
            End If
    End Select
    DatosOk = b

End Function



Private Function ConcatenarCampos(cTabla As String, cWhere As String) As String
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String

    ConcatenarCampos = ""

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    
    Sql = "select distinct rcampos.codcampo  from " & cTabla & " where " & cWhere
    Set RS = New ADODB.Recordset
    
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not RS.EOF
        Sql1 = Sql1 & DBLet(RS.Fields(0).Value, "N") & ","
        RS.MoveNext
    Wend
    Set RS = Nothing
    'quitamos el ultimo or
    ConcatenarCampos = Mid(Sql1, 1, Len(Sql1) - 1)
    
End Function

Private Function CargarTemporalProdReal(cTabla As String, cWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String

Dim vCampAnt As CCampAnt

Dim Cad As String
Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalProdReal = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
    End If
        
    ' insertamos en la temporal con los kilos por defecto a cero
    '                                       codsocio, codcampo, kilos
    Sql = "insert into tmpinformes (codusu, importe1, importe2, importe3)    "
    Sql = Sql & "select " & DBSet(vUsu.Codigo, "N") & ",rcampos.codsocio, rcampos.codcampo, 0 from " & cTabla
    Sql = Sql & " where " & cWhere
    Sql = Sql & " group by 1,2,3 "
    Sql = Sql & " order by 1,2,3 "
    
    conn.Execute Sql
    
    If Option1(2).Value Then
        ' recordset para actualizar los kilos de la campaña anterior de la temporal
        Sql = "SELECT rhisfruta.codsocio, rhisfruta.codcampo, sum(kilosnet) as kilos "
        Sql = Sql & " FROM  (" & cTabla & ") INNER JOIN rhisfruta ON rhisfruta.codcampo = rcampos.codcampo "
        Sql = Sql & " and rhisfruta.codsocio = rcampos.codsocio where " & cWhere
        
        If txtCodigo(6).Text <> "" Then Sql = Sql & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(6).Text, "F")
        If txtCodigo(7).Text <> "" Then Sql = Sql & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(7).Text, "F")
        
        Sql = Sql & " group by 1,2 "
        Sql = Sql & " order by 1,2 "
    
        '[Monica]01/04/2011: en el caso de informe para seguros se coge de la campaña actual
        If Check3.Value = 1 Then
            Set RS = New ADODB.Recordset
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Cad = ""
            While Not RS.EOF
                Cad = Cad & "(" & vUsu.Codigo & "," & DBLet(RS!Codsocio, "N") & "," & DBLet(RS!codcampo, "N") & ","
                Cad = Cad & DBLet(RS!Kilos, "N") & "),"
                
                Sql2 = "update tmpinformes set importe3 = importe3 + " & DBLet(RS!Kilos, "N")
                Sql2 = Sql2 & " where codusu = " & DBSet(vUsu.Codigo, "N")
                Sql2 = Sql2 & " and importe1 = " & DBSet(RS!Codsocio, "N")
                Sql2 = Sql2 & " and importe2 = " & DBSet(RS!codcampo, "N")
                
                conn.Execute Sql2
                
                RS.MoveNext
            Wend
            
            Set RS = Nothing
        Else
            Set vCampAnt = New CCampAnt
            
            If vCampAnt.Leer = 0 Then
                If AbrirConexionCampAnterior(vCampAnt.BaseDatos) Then
                    Set RS = New ADODB.Recordset
                    RS.Open Sql, ConnCAnt, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    Cad = ""
                    While Not RS.EOF
                        Cad = Cad & "(" & vUsu.Codigo & "," & DBLet(RS!Codsocio, "N") & "," & DBLet(RS!codcampo, "N") & ","
                        Cad = Cad & DBLet(RS!Kilos, "N") & "),"
                        
                        Sql2 = "update tmpinformes set importe3 = importe3 + " & DBLet(RS!Kilos, "N")
                        Sql2 = Sql2 & " where codusu = " & DBSet(vUsu.Codigo, "N")
                        Sql2 = Sql2 & " and importe1 = " & DBSet(RS!Codsocio, "N")
                        Sql2 = Sql2 & " and importe2 = " & DBSet(RS!codcampo, "N")
                        
                        conn.Execute Sql2
                        
                        RS.MoveNext
                    Wend
                    
                    Set RS = Nothing
                    
                    CerrarConexionCampAnterior
                End If
            End If
        End If
        Set vCampAnt = Nothing
    End If
    
    CargarTemporalProdReal = True
    Exit Function
    
eCargarTemporal:
    CargarTemporalProdReal = False
    MuestraError "Cargando temporal", Err.Description
End Function



Private Sub CargaCombo()
        
    Combo1(0).Clear
    'tipo de superficie
    Combo1(0).AddItem "Hectáreas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Hanegadas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
        
    Combo1(1).Clear
    'tipo de recoleccion
    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Ambos"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2

End Sub
    
Private Sub ProcesarCambiosAforos(cTabla As String, cWhere As String)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Porcentaje As Currency

    On Error GoTo eProcesarCambiosAforos

    Sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not BloqueaRegistro(cTabla, cWhere) Then
        MsgBox "No se puede realizar el proceso. Hay registros de campos bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
       
    Porcentaje = CCur(ImporteSinFormato(txtCodigo(26).Text))
    
    
    Sql = " update rcampos, variedades "
    Sql = Sql & " set rcampos.canaforo = round(rcampos.canaforo * (1 + ((" & DBSet(Porcentaje, "N") & ") / 100)), 0)"
    Sql = Sql & " where " & cWhere & " and rcampos.codvarie = variedades.codvarie"
    
    conn.Execute Sql
    
    TerminaBloquear
    
    MsgBox "Proceso realizado correctamente.", vbExclamation
    Exit Sub
    
eProcesarCambiosAforos:
    MuestraError Err.Number, "Procesar cambios aforos", Err.Description
End Sub
