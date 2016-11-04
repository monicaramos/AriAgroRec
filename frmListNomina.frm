VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListNomina 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8145
   Icon            =   "frmListNomina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameInfDiasTrabajados 
      Height          =   4275
      Left            =   30
      TabIndex        =   259
      Top             =   30
      Width           =   6375
      Begin VB.CheckBox Check6 
         Caption         =   "Exportar Cadena para Excel"
         Height          =   375
         Left            =   570
         TabIndex        =   275
         Top             =   3420
         Width           =   2775
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   4770
         TabIndex        =   261
         Top             =   3375
         Width           =   975
      End
      Begin VB.CommandButton CmdDiasTrabajados 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3750
         TabIndex        =   260
         Top             =   3390
         Width           =   975
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2985
         Left            =   420
         TabIndex        =   262
         Top             =   900
         Width           =   5595
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   69
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   268
            Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   765
            Width           =   750
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   68
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   267
            Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   405
            Width           =   750
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   68
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   266
            Text            =   "Text5"
            Top             =   405
            Width           =   3015
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   69
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   265
            Text            =   "Text5"
            Top             =   780
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   67
            Left            =   1380
            MaxLength       =   4
            TabIndex        =   264
            Top             =   1380
            Width           =   840
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   1380
            TabIndex        =   263
            Text            =   "Combo2"
            Top             =   1950
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   92
            Left            =   390
            TabIndex        =   273
            Top             =   420
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   91
            Left            =   390
            TabIndex        =   272
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Trabajador"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   90
            Left            =   180
            TabIndex        =   271
            Top             =   60
            Width           =   765
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   30
            Left            =   1080
            MouseIcon       =   "frmListNomina.frx":000C
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   810
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   29
            Left            =   1080
            MouseIcon       =   "frmListNomina.frx":015E
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Año"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   89
            Left            =   180
            TabIndex        =   270
            Top             =   1410
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   88
            Left            =   180
            TabIndex        =   269
            Top             =   2010
            Width           =   300
         End
      End
      Begin VB.Label Label17 
         Caption         =   "Informe Mensual Días Trabajados"
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
         TabIndex        =   274
         Top             =   390
         Width           =   5595
      End
   End
   Begin VB.Frame FrameListMensAsesoria 
      Height          =   4275
      Left            =   30
      TabIndex        =   237
      Top             =   60
      Width           =   6375
      Begin VB.CommandButton CmdAcepInfAse 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   243
         Top             =   3390
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   4770
         TabIndex        =   244
         Top             =   3375
         Width           =   975
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2985
         Left            =   390
         TabIndex        =   238
         Top             =   930
         Width           =   5595
         Begin VB.CheckBox Check2 
            Caption         =   "Exportar Cadena para Excel"
            Height          =   375
            Left            =   180
            TabIndex        =   253
            Top             =   2550
            Width           =   2775
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   1380
            TabIndex        =   242
            Text            =   "Combo2"
            Top             =   1950
            Width           =   1575
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   61
            Left            =   1380
            MaxLength       =   4
            TabIndex        =   241
            Top             =   1380
            Width           =   840
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   65
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   246
            Text            =   "Text5"
            Top             =   780
            Width           =   3015
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   64
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   245
            Text            =   "Text5"
            Top             =   405
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   64
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   239
            Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   405
            Width           =   750
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   65
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   240
            Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   765
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   86
            Left            =   180
            TabIndex        =   252
            Top             =   2010
            Width           =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Año"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   85
            Left            =   180
            TabIndex        =   250
            Top             =   1410
            Width           =   285
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   28
            Left            =   1080
            MouseIcon       =   "frmListNomina.frx":02B0
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   780
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   27
            Left            =   1080
            MouseIcon       =   "frmListNomina.frx":0402
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Trabajador"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   84
            Left            =   180
            TabIndex        =   249
            Top             =   60
            Width           =   765
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   83
            Left            =   390
            TabIndex        =   248
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   82
            Left            =   390
            TabIndex        =   247
            Top             =   420
            Width           =   465
         End
      End
      Begin VB.Label Label15 
         Caption         =   "Informe Mensual Asesoria"
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
         TabIndex        =   251
         Top             =   390
         Width           =   5595
      End
   End
   Begin VB.Frame FrameEventuales 
      Height          =   5535
      Left            =   0
      TabIndex        =   130
      Top             =   0
      Width           =   6375
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3315
         Left            =   300
         TabIndex        =   140
         Top             =   1470
         Width           =   5595
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   1530
            MaxLength       =   6
            TabIndex        =   135
            Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   1995
            Width           =   750
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   1530
            MaxLength       =   6
            TabIndex        =   134
            Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   1635
            Width           =   750
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   41
            Left            =   2370
            Locked          =   -1  'True
            TabIndex        =   148
            Text            =   "Text5"
            Top             =   1635
            Width           =   3015
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   42
            Left            =   2370
            Locked          =   -1  'True
            TabIndex        =   147
            Text            =   "Text5"
            Top             =   2010
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   270
            MaxLength       =   13
            TabIndex        =   136
            Top             =   3000
            Width           =   1140
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   132
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   1545
            MaxLength       =   10
            TabIndex        =   133
            Top             =   900
            Width           =   1005
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   54
            Left            =   540
            TabIndex        =   151
            Top             =   1650
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   53
            Left            =   540
            TabIndex        =   150
            Top             =   2010
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Trabajador"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   52
            Left            =   240
            TabIndex        =   149
            Top             =   1410
            Width           =   765
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   13
            Left            =   1215
            MouseIcon       =   "frmListNomina.frx":0554
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   1995
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   1215
            MouseIcon       =   "frmListNomina.frx":06A6
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   1620
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   50
            Left            =   270
            TabIndex        =   144
            Top             =   2700
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   44
            Left            =   270
            TabIndex        =   143
            Top             =   180
            Width           =   450
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   12
            Left            =   1215
            Picture         =   "frmListNomina.frx":07F8
            Top             =   540
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   11
            Left            =   1200
            Picture         =   "frmListNomina.frx":0883
            Top             =   900
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   43
            Left            =   510
            TabIndex        =   142
            Top             =   930
            Width           =   420
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   33
            Left            =   510
            TabIndex        =   141
            Top             =   570
            Width           =   465
         End
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   4770
         TabIndex        =   138
         Top             =   4785
         Width           =   975
      End
      Begin VB.CommandButton CmdEventuales 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   137
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   28
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   139
         Text            =   "Text5"
         Top             =   1110
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   131
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1110
         Width           =   750
      End
      Begin VB.Label Label9 
         Caption         =   "Alta Eventuales"
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
         TabIndex        =   146
         Top             =   390
         Width           =   5595
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":090E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   51
         Left            =   570
         TabIndex        =   145
         Top             =   1140
         Width           =   630
      End
   End
   Begin VB.Frame FrameTrabajadoresCapataz 
      Height          =   5055
      Left            =   -60
      TabIndex        =   152
      Top             =   90
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   47
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   156
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1110
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   47
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   155
         Text            =   "Text5"
         Top             =   1110
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   46
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   157
         Top             =   1605
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   45
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   154
         Text            =   "Text5"
         Top             =   2085
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   45
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   158
         Tag             =   "Código|N|N|0|9999|straba|codtraba|0000|S|"
         Top             =   2070
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepTrabajCapataz 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   160
         Top             =   4230
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   4770
         TabIndex        =   153
         Top             =   4245
         Width           =   975
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1395
         Left            =   300
         TabIndex        =   161
         Top             =   2670
         Width           =   5085
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   270
            MaxLength       =   13
            TabIndex        =   159
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   60
            Left            =   270
            TabIndex        =   162
            Top             =   120
            Width           =   525
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   63
         Left            =   570
         TabIndex        =   166
         Top             =   1140
         Width           =   630
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":0A60
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1125
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   1515
         Picture         =   "frmListNomina.frx":0BB2
         Top             =   1605
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1515
         MouseIcon       =   "frmListNomina.frx":0C3D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   2085
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   62
         Left            =   570
         TabIndex        =   165
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label Label10 
         Caption         =   "Trabajadores de un capataz"
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
         TabIndex        =   164
         Top             =   390
         Width           =   5595
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   61
         Left            =   570
         TabIndex        =   163
         Top             =   2100
         Width           =   585
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6750
      Top             =   4980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FramePaseABanco 
      Height          =   5490
      Left            =   60
      TabIndex        =   215
      Top             =   30
      Width           =   6435
      Begin VB.Frame FrameConcep 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   540
         TabIndex        =   257
         Top             =   4050
         Width           =   5205
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   66
            Left            =   1290
            MaxLength       =   30
            TabIndex        =   228
            Top             =   120
            Width           =   3870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   87
            Left            =   0
            TabIndex        =   258
            Top             =   90
            Width           =   840
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   227
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   3720
         Width           =   1665
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   62
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   234
         Text            =   "Text5"
         Top             =   1230
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   63
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   235
         Text            =   "Text5"
         Top             =   1575
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   63
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   223
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1575
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   62
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   222
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1215
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepPaseBanco 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   3570
         TabIndex        =   229
         Top             =   4875
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   4650
         TabIndex        =   230
         Top             =   4890
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   225
         Top             =   2670
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   59
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   224
         Top             =   2220
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   226
         Tag             =   "Código|N|N|0|9999|rcapataz|codcapat|0000|S|"
         Top             =   3120
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   58
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   236
         Text            =   "Text5"
         Top             =   3120
         Width           =   3015
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   240
         Left            =   480
         TabIndex        =   216
         Top             =   4560
         Visible         =   0   'False
         Width           =   5250
         _ExtentX        =   9260
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
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   1530
         MouseIcon       =   "frmListNomina.frx":0D8F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1575
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1530
         MouseIcon       =   "frmListNomina.frx":0EE1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   81
         Left            =   510
         TabIndex        =   233
         Top             =   990
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   80
         Left            =   870
         TabIndex        =   232
         Top             =   1590
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   79
         Left            =   870
         TabIndex        =   231
         Top             =   1230
         Width           =   465
      End
      Begin VB.Label Label16 
         Caption         =   "Pase a Banco"
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
         Left            =   510
         TabIndex        =   221
         Top             =   405
         Width           =   5835
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1515
         Picture         =   "frmListNomina.frx":1033
         ToolTipText     =   "Buscar fecha"
         Top             =   2670
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pago"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   78
         Left            =   540
         TabIndex        =   220
         Top             =   2490
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   77
         Left            =   540
         TabIndex        =   219
         Top             =   1950
         Width           =   1005
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   1530
         Picture         =   "frmListNomina.frx":10BE
         ToolTipText     =   "Buscar fecha"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Banco "
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   76
         Left            =   540
         TabIndex        =   218
         Top             =   3090
         Width           =   510
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1515
         MouseIcon       =   "frmListNomina.frx":1149
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar banco"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "Concepto Transferencia "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   540
         TabIndex        =   217
         Top             =   3450
         Width           =   1875
      End
   End
   Begin VB.Frame FrameEntradasCapataz 
      Height          =   3885
      Left            =   0
      TabIndex        =   183
      Top             =   0
      Width           =   6285
      Begin VB.CheckBox Check4 
         Caption         =   "Imprimir resumen"
         Height          =   225
         Left            =   630
         TabIndex        =   254
         Top             =   3180
         Width           =   1995
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   53
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   187
         Top             =   2670
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   52
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   186
         Top             =   2340
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   38
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   191
         Text            =   "Text5"
         Top             =   1290
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   43
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   190
         Text            =   "Text5"
         Top             =   1665
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   185
         Tag             =   "Código|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   184
         Tag             =   "Código|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepEntCapataz 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   188
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   4770
         TabIndex        =   189
         Top             =   3135
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1545
         Picture         =   "frmListNomina.frx":129B
         Top             =   2670
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1560
         Picture         =   "frmListNomina.frx":1326
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1560
         MouseIcon       =   "frmListNomina.frx":13B1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":1503
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   69
         Left            =   600
         TabIndex        =   198
         Top             =   2130
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   68
         Left            =   960
         TabIndex        =   197
         Top             =   2685
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   67
         Left            =   960
         TabIndex        =   196
         Top             =   2370
         Width           =   465
      End
      Begin VB.Label Label12 
         Caption         =   "Entradas Capataz"
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
         TabIndex        =   195
         Top             =   390
         Width           =   5505
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   56
         Left            =   600
         TabIndex        =   194
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   55
         Left            =   960
         TabIndex        =   193
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   45
         Left            =   960
         TabIndex        =   192
         Top             =   1320
         Width           =   465
      End
   End
   Begin VB.Frame FrameHorasDestajo 
      Height          =   5565
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   7515
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   47
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   2685
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   46
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   2310
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text5"
         Top             =   2325
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text5"
         Top             =   2700
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Informe para el Trabajador"
         Height          =   195
         Left            =   630
         TabIndex        =   61
         Top             =   4320
         Width           =   2220
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   49
         Top             =   3690
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   48
         Top             =   3345
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Text5"
         Top             =   1305
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   44
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1305
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   45
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1665
         Width           =   750
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   3690
         TabIndex        =   50
         Top             =   4650
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4770
         TabIndex        =   51
         Top             =   4665
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   66
         Top             =   2340
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   65
         Top             =   2700
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   12
         Left            =   600
         TabIndex        =   64
         Top             =   2100
         Width           =   630
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":1655
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2685
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":17A7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2325
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1575
         Picture         =   "frmListNomina.frx":18F9
         Top             =   3690
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   1575
         Picture         =   "frmListNomina.frx":1984
         Top             =   3345
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":1A0F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":1B61
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   600
         TabIndex        =   60
         Top             =   3120
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   59
         Top             =   3675
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   960
         TabIndex        =   58
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Informe de Horas Trabajadas Destajo"
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
         TabIndex        =   57
         Top             =   390
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   56
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   55
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   54
         Top             =   1320
         Width           =   465
      End
   End
   Begin VB.Frame FrameBorradoMasivoETT 
      Height          =   3885
      Left            =   0
      TabIndex        =   96
      Top             =   -60
      Width           =   6285
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4770
         TabIndex        =   106
         Top             =   3135
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepBorradoMasivo 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   3690
         TabIndex        =   104
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   32
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   100
         Tag             =   "Código|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   99
         Tag             =   "Código|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   1305
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   31
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   98
         Text            =   "Text5"
         Top             =   1305
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   32
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   97
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   102
         Top             =   2700
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   101
         Top             =   2370
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   42
         Left            =   960
         TabIndex        =   111
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   40
         Left            =   960
         TabIndex        =   110
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   39
         Left            =   600
         TabIndex        =   109
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label5 
         Caption         =   "Borrado Masivo ETT"
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
         TabIndex        =   108
         Top             =   390
         Width           =   5505
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   37
         Left            =   960
         TabIndex        =   107
         Top             =   2370
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   35
         Left            =   960
         TabIndex        =   105
         Top             =   2685
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   34
         Left            =   600
         TabIndex        =   103
         Top             =   2130
         Width           =   450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":1CB3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":1E05
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   1560
         Picture         =   "frmListNomina.frx":1F57
         Top             =   2700
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1575
         Picture         =   "frmListNomina.frx":1FE2
         Top             =   2370
         Width           =   240
      End
   End
   Begin VB.Frame FrameInfComprobacion 
      Height          =   4425
      Left            =   0
      TabIndex        =   167
      Top             =   0
      Width           =   6915
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   4770
         TabIndex        =   177
         Top             =   3435
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepInfComprob 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   175
         Top             =   3420
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   50
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   171
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   170
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1305
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   49
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   169
         Text            =   "Text5"
         Top             =   1305
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   50
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   168
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   172
         Top             =   2175
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   48
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   173
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   66
         Left            =   960
         TabIndex        =   182
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   65
         Left            =   960
         TabIndex        =   181
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   64
         Left            =   600
         TabIndex        =   180
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label11 
         Caption         =   "Informe de Comprobación"
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
         TabIndex        =   179
         Top             =   390
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   59
         Left            =   960
         TabIndex        =   178
         Top             =   2190
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   58
         Left            =   960
         TabIndex        =   176
         Top             =   2505
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   57
         Left            =   600
         TabIndex        =   174
         Top             =   1950
         Width           =   450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":206D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":21BF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   1560
         Picture         =   "frmListNomina.frx":2311
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1560
         Picture         =   "frmListNomina.frx":239C
         Top             =   2160
         Width           =   240
      End
   End
   Begin VB.Frame FrameHorasTrabajadas 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   7515
      Begin VB.CheckBox Check3 
         Caption         =   "Sobre Horas Productivas"
         Height          =   195
         Left            =   600
         TabIndex        =   26
         Top             =   3360
         Width           =   2220
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   4560
         TabIndex        =   10
         Top             =   3735
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   8
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1305
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   1305
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2745
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2370
         Width           =   1005
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   4
         Left            =   5355
         TabIndex        =   16
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
         Height          =   195
         Index           =   29
         Left            =   960
         TabIndex        =   15
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   960
         TabIndex        =   14
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   600
         TabIndex        =   13
         Top             =   1080
         Width           =   765
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
         TabIndex        =   12
         Top             =   405
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   11
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   960
         TabIndex        =   9
         Top             =   2715
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   600
         TabIndex        =   7
         Top             =   2160
         Width           =   450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1620
         MouseIcon       =   "frmListNomina.frx":2427
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1620
         MouseIcon       =   "frmListNomina.frx":2579
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1575
         Picture         =   "frmListNomina.frx":26CB
         Top             =   2745
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1575
         Picture         =   "frmListNomina.frx":2756
         Top             =   2340
         Width           =   240
      End
   End
   Begin VB.Frame FramePagoPartesCampo 
      Height          =   4455
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   6345
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   360
         TabIndex        =   255
         Top             =   3090
         Width           =   4155
         Begin VB.CheckBox Check5 
            Caption         =   "Prevision de Pago de Partes"
            Height          =   225
            Left            =   270
            TabIndex        =   256
            Top             =   240
            Width           =   2445
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   2820
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   33
         Top             =   2745
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   32
         Top             =   2340
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   31
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   30
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1260
         Width           =   750
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   35
         Top             =   3690
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   37
         Top             =   3690
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1575
         Picture         =   "frmListNomina.frx":27E1
         Top             =   2745
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1575
         Picture         =   "frmListNomina.frx":286C
         Top             =   2340
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   42
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   41
         Top             =   2715
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   40
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Pago de Partes Campo"
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
         TabIndex        =   39
         Top             =   450
         Width           =   4350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parte"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   38
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   36
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   34
         Top             =   1320
         Width           =   465
      End
   End
   Begin VB.Frame FrameCalculoETT 
      Height          =   5055
      Left            =   0
      TabIndex        =   67
      Top             =   30
      Width           =   6375
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4770
         TabIndex        =   75
         Top             =   4245
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepCalculoETT 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   74
         Top             =   4230
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   70
         Tag             =   "Código|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   2070
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Text5"
         Top             =   2085
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   69
         Top             =   1605
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text5"
         Top             =   1110
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   68
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1110
         Width           =   750
      End
      Begin VB.Frame FrameDestajo 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1395
         Left            =   270
         TabIndex        =   82
         Top             =   2700
         Width           =   5085
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   3240
            MaxLength       =   13
            TabIndex        =   73
            Top             =   720
            Width           =   1140
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   1830
            MaxLength       =   6
            TabIndex        =   72
            Top             =   720
            Width           =   1140
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   360
            MaxLength       =   6
            TabIndex        =   71
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   19
            Left            =   3240
            TabIndex        =   85
            Top             =   390
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   16
            Left            =   1830
            TabIndex        =   84
            Top             =   390
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kilos Entrados"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   15
            Left            =   360
            TabIndex        =   83
            Top             =   390
            Width           =   1005
         End
      End
      Begin VB.Frame FrameBonificacion 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1395
         Left            =   300
         TabIndex        =   93
         Top             =   2670
         Width           =   5085
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   270
            MaxLength       =   13
            TabIndex        =   94
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   30
            Left            =   270
            TabIndex        =   95
            Top             =   120
            Width           =   525
         End
      End
      Begin VB.Frame FramePenalizacion 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1395
         Left            =   180
         TabIndex        =   86
         Top             =   2550
         Width           =   5085
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   22
            Left            =   360
            MaxLength       =   6
            TabIndex        =   89
            Top             =   720
            Width           =   1140
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   1830
            MaxLength       =   6
            TabIndex        =   88
            Top             =   720
            Width           =   1140
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   20
            Left            =   3240
            MaxLength       =   11
            TabIndex        =   87
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kilos Entrados"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   23
            Left            =   360
            TabIndex        =   92
            Top             =   390
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "% Penalizacion"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   22
            Left            =   1830
            TabIndex        =   91
            Top             =   390
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   20
            Left            =   3240
            TabIndex        =   90
            Top             =   390
            Width           =   525
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   21
         Left            =   570
         TabIndex        =   81
         Top             =   2100
         Width           =   585
      End
      Begin VB.Label Label4 
         Caption         =   "Destajo Alicatado"
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
         TabIndex        =   80
         Top             =   390
         Width           =   5595
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   18
         Left            =   570
         TabIndex        =   79
         Top             =   1620
         Width           =   450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1515
         MouseIcon       =   "frmListNomina.frx":28F7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   2085
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1515
         Picture         =   "frmListNomina.frx":2A49
         Top             =   1605
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":2AD4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   17
         Left            =   570
         TabIndex        =   78
         Top             =   1140
         Width           =   630
      End
   End
   Begin VB.Frame FrameCalculoHorasProductivas 
      Height          =   3525
      Left            =   90
      TabIndex        =   17
      Top             =   30
      Width           =   5835
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   24
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   2190
         Width           =   2955
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   20
         Top             =   2190
         Width           =   960
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1290
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   19
         Top             =   1740
         Width           =   990
      End
      Begin VB.CommandButton CmdAcepCalHProd 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2790
         TabIndex        =   21
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelCalHProd 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3900
         TabIndex        =   22
         Top             =   2760
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1290
         MouseIcon       =   "frmListNomina.frx":2C26
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar almacén"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   36
         Left            =   570
         TabIndex        =   28
         Top             =   2250
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1320
         Picture         =   "frmListNomina.frx":2D78
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   41
         Left            =   570
         TabIndex        =   25
         Top             =   1290
         Width           =   450
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
         TabIndex        =   24
         Top             =   480
         Width           =   4725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   38
         Left            =   570
         TabIndex        =   23
         Top             =   1800
         Width           =   765
      End
   End
   Begin VB.Frame FrameAltaRapida 
      Height          =   5055
      Left            =   30
      TabIndex        =   112
      Top             =   90
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   114
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1110
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   36
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text5"
         Top             =   1110
         Width           =   3015
      End
      Begin VB.CommandButton CmdAltaRapida 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   119
         Top             =   4260
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   4770
         TabIndex        =   120
         Top             =   4245
         Width           =   975
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2415
         Left            =   300
         TabIndex        =   121
         Top             =   1620
         Width           =   5595
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   1545
            MaxLength       =   10
            TabIndex        =   116
            Top             =   900
            Width           =   1005
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   115
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   34
            Left            =   2370
            Locked          =   -1  'True
            TabIndex        =   125
            Text            =   "Text5"
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   1530
            MaxLength       =   6
            TabIndex        =   117
            Tag             =   "Código|N|N|0|999999|straba|codtraba|0000|S|"
            Top             =   1305
            Width           =   750
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   270
            MaxLength       =   13
            TabIndex        =   118
            Top             =   2070
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   32
            Left            =   510
            TabIndex        =   129
            Top             =   570
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   31
            Left            =   510
            TabIndex        =   128
            Top             =   930
            Width           =   420
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   10
            Left            =   1200
            Picture         =   "frmListNomina.frx":2E03
            Top             =   900
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   9
            Left            =   1215
            Picture         =   "frmListNomina.frx":2E8E
            Top             =   540
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1215
            MouseIcon       =   "frmListNomina.frx":2F19
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   48
            Left            =   270
            TabIndex        =   127
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Capataz"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   47
            Left            =   270
            TabIndex        =   126
            Top             =   1335
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   46
            Left            =   270
            TabIndex        =   122
            Top             =   1770
            Width           =   525
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   49
         Left            =   570
         TabIndex        =   124
         Top             =   1140
         Width           =   630
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":306B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Alta Rápida"
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
         TabIndex        =   123
         Top             =   390
         Width           =   5595
      End
   End
   Begin VB.Frame FrameBorradoAsesoria 
      Height          =   4215
      Left            =   60
      TabIndex        =   199
      Top             =   0
      Width           =   6705
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   57
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   203
         Top             =   2730
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   56
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   202
         Top             =   2385
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   55
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   206
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   54
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   204
         Text            =   "Text5"
         Top             =   1305
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   55
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   201
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   54
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   200
         Tag             =   "Código|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1305
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepBorrAse 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3630
         TabIndex        =   205
         Top             =   3330
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   4710
         TabIndex        =   207
         Top             =   3345
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   1605
         Picture         =   "frmListNomina.frx":31BD
         Top             =   2730
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   1605
         Picture         =   "frmListNomina.frx":3248
         Top             =   2385
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   1620
         MouseIcon       =   "frmListNomina.frx":32D3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   22
         Left            =   1620
         MouseIcon       =   "frmListNomina.frx":3425
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   75
         Left            =   600
         TabIndex        =   214
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   74
         Left            =   960
         TabIndex        =   213
         Top             =   2715
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   73
         Left            =   960
         TabIndex        =   212
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label13 
         Caption         =   "Borrado de Movimientos"
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
         TabIndex        =   211
         Top             =   405
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   72
         Left            =   600
         TabIndex        =   210
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   71
         Left            =   960
         TabIndex        =   209
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   70
         Left            =   960
         TabIndex        =   208
         Top             =   1320
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmListNomina"
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
    ' 16 .- Calculo de Horas productivas
    ' 17 .- Proceso de pago de partes de campo
    
    ' 18 .- Listado de Horas trabajadas destajo
    ' 19 .- Actualizar entradas de horas de destajo
    
    '==== HORAS ETT
    '=============================
    ' 20 .- Introduccion masiva de horas ett (destajo alicatado)
    ' 21 .- Calculo de penalizacion
    ' 22 .- calculo de bonificacion
    ' 23 .- Borrado masivo de ETT
    
    ' 29 .- Listado de entradas capataz
    ' 38 .- Rendimiento por capataz
    
    '==== HORAS
    '=============================
    ' 24 .- Introduccion masiva de horas (alta rapida)
    ' 25 .- Eventuales
    ' 26 .- Trabajador de un capataz
    ' 27 .- Borrado masivo
    
    ' 28 .- Listado de Comprobacion
    
    '==== HORAS DESTAJO
    '=============================
    ' 30 .- Introduccion masiva de horas (destajo alicatado)
    ' 31 .- Calculo de penalizacion
    ' 32 .- Calculo de bonificacion
    ' 33 .- Borrado masivo
    
    
    '==== ASESORIA
    '=============================
    ' 34 .- Listado para Asesoria
    
    ' 35 .- Borrado Masivo de movimientos de Asesoria
    ' 36 .- Pase a Banco de movimientos de Asesoria
    
    ' 37 .- Listado mensual de horas para asesoria
    
    
    ' 39 .- Informe de dias trabajados dentro del mto de partes
    
    
    
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
Private WithEvents frmAlm As frmComercial 'mantenimiento de almacenes propios de comercial
Attribute frmAlm.VB_VarHelpID = -1
 
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCap As frmManCapataz 'capataces
Attribute frmCap.VB_VarHelpID = -1

Private WithEvents frmBan As frmComercial 'Banco propio
Attribute frmBan.VB_VarHelpID = -1

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
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub


Private Sub CmdAcepBorradoMasivo_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim cWhere As String
Dim Sql As String

       InicializarVbles
       
        'D/H Capataz
        cDesde = Trim(txtCodigo(31).Text)
        cHasta = Trim(txtCodigo(32).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codcapat}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        'D/H Fecha
        cDesde = Trim(txtCodigo(29).Text)
        cHasta = Trim(txtCodigo(30).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fechahora}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If

        If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".fecharec} is null ") Then Exit Sub


        cTabla = Tabla
        cWhere = cadSelect
        
        cTabla = QuitarCaracterACadena(cTabla, "{")
        cTabla = QuitarCaracterACadena(cTabla, "}")
        Sql = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            cWhere = QuitarCaracterACadena(cWhere, "{")
            cWhere = QuitarCaracterACadena(cWhere, "}")
            cWhere = QuitarCaracterACadena(cWhere, "_1")
            Sql = Sql & " WHERE " & cWhere
        End If
    
        If RegistrosAListar(Sql) <> 0 Then
            If ProcesoBorradoMasivo(cTabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
                Exit Sub
            Else
                MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "No hay registros entre esos límites.", vbExclamation
        End If

End Sub

Private Sub CmdAcepBorrAse_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim cWhere As String
Dim Sql As String


       InicializarVbles
       
        'D/H Trabajador
        cDesde = Trim(txtCodigo(54).Text)
        cHasta = Trim(txtCodigo(55).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codtraba}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        'D/H Fecha
        cDesde = Trim(txtCodigo(56).Text)
        cHasta = Trim(txtCodigo(57).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fechahora}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If

        If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".idconta} = 1") Then Exit Sub


        cTabla = Tabla
        cWhere = cadSelect
        
        cTabla = QuitarCaracterACadena(cTabla, "{")
        cTabla = QuitarCaracterACadena(cTabla, "}")
        Sql = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            cWhere = QuitarCaracterACadena(cWhere, "{")
            cWhere = QuitarCaracterACadena(cWhere, "}")
            cWhere = QuitarCaracterACadena(cWhere, "_1")
            Sql = Sql & " WHERE " & cWhere
        End If
    
        If RegistrosAListar(Sql) <> 0 Then
            If ProcesoBorradoMasivo(cTabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
                Exit Sub
            Else
                MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "No hay registros entre esos límites.", vbExclamation
        End If

        

End Sub

Private Sub CmdAcepCalculoETT_Click()
Dim Sql As String
Dim CodigoETT As String

    If txtCodigo(9).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo Variedad.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(11).Text = "" Then
        MsgBox "Debe introducir una Fecha para realizar el cálculo.", vbExclamation
        Exit Sub
    End If

    If txtCodigo(12).Text = "" Then
        MsgBox "Debe introducir el capataz para realizar el cálculo.", vbExclamation
        Exit Sub
    End If
    
    Select Case OpcionListado
        Case 20 'horasett: calculo de destajo alicatado ett
            If CalculoDestajoETT(True) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
               
                cmdCancel_Click (2)
            End If
            
        Case 30 ' horas: calculo de destajo alicatado
            If CalculoDestajo(True) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
               
                cmdCancel_Click (2)
            End If
            
        Case 21 'horasett: calculo de penalizacion ett
            Sql = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
            
            CodigoETT = DevuelveValor(Sql)
        
            Sql = "select count(*) from horasett where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            Sql = Sql & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            Sql = Sql & " and codigoett = " & DBSet(CodigoETT, "N")
            If TotalRegistros(Sql) = 0 Then
                MsgBox "No existe registro para realizar la penalización. Revise.", vbExclamation
            Else
                If CalculoPenalizacionETT(True) Then
                     MsgBox "Proceso realizado correctamente.", vbExclamation
                    
                     cmdCancel_Click (2)
                End If
            End If
        
        Case 31 'horas: calculo de penalizacion
            Sql = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            Sql = Sql & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            If TotalRegistros(Sql) = 0 Then
                MsgBox "No existen registros para realizar la penalización. Revise.", vbExclamation
            Else
                If CalculoPenalizacion(True) Then
                     MsgBox "Proceso realizado correctamente.", vbExclamation
                    
                     cmdCancel_Click (2)
                End If
            End If
        
        Case 22 ' horasett: calculo de bonificacion
            Sql = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
            
            CodigoETT = DevuelveValor(Sql)
        
            Sql = "select count(*) from horasett where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            Sql = Sql & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            Sql = Sql & " and codigoett = " & DBSet(CodigoETT, "N")
            If TotalRegistros(Sql) = 0 Then
                MsgBox "No existen registros para realizar la bonificación. Revise.", vbExclamation
            Else
                If CalculoBonificacionETT(True) Then
                     MsgBox "Proceso realizado correctamente.", vbExclamation
                    
                     cmdCancel_Click (2)
                End If
            End If
        
        Case 32 ' horas: calculo de bonificacion
            Sql = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            Sql = Sql & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            If TotalRegistros(Sql) = 0 Then
                MsgBox "No existen registros para realizar la bonificación. Revise.", vbExclamation
            Else
                If CalculoBonificacion(True) Then
                     MsgBox "Proceso realizado correctamente.", vbExclamation
                    
                     cmdCancel_Click (2)
                End If
            End If
        
    End Select
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

Private Sub CmdAcepEntCapataz_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim cWhere As String
Dim Sql As String

       InicializarVbles
       
        'Añadir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

        'D/H Capataz
        cDesde = Trim(txtCodigo(38).Text)
        cHasta = Trim(txtCodigo(43).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codcapat}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        'D/H Fecha
        cDesde = Trim(txtCodigo(52).Text)
        cHasta = Trim(txtCodigo(53).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fechahora}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If

'?????        If Not AnyadirAFormula(cadSelect, "{" & Tabla & ".fecharec} is null ") Then Exit Sub


        CadParam = CadParam & "pResumen=" & Check4.Value & "|"
        numParam = numParam + 1


        cTabla = Tabla
        cWhere = cadSelect
        
        cTabla = QuitarCaracterACadena(cTabla, "{")
        cTabla = QuitarCaracterACadena(cTabla, "}")
        Sql = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            cWhere = QuitarCaracterACadena(cWhere, "{")
            cWhere = QuitarCaracterACadena(cWhere, "}")
            cWhere = QuitarCaracterACadena(cWhere, "_1")
            Sql = Sql & " WHERE " & cWhere
        End If
    
        If OpcionListado = 29 Then
            ' entradas por capataz
            If ProcesoEntradasCapataz(cTabla, cadSelect) Then
                If HayRegParaInforme("tmpinformes", "{tmpinformes.codusu} = " & vUsu.Codigo) Then
                    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                    
                    cadNombreRPT = "rInfEntradasCapataz.rpt"
                    cadTitulo = "Informe de Entradas Capataz"
                    ConSubInforme = True
                    LlamarImprimir
                Else
                    MsgBox "No hay registros entre esos límites.", vbExclamation
                End If
            End If
        Else
            ' rendimiento por capataz
            If ProcesoEntradasCapatazRdto(cTabla, cadSelect) Then
                If HayRegParaInforme("tmpinformes", "{tmpinformes.codusu} = " & vUsu.Codigo) Then
                    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                    
                    cadNombreRPT = "rRdtoCapataz.rpt"
                    cadTitulo = "Rendimiento por Capataz"
                    ConSubInforme = False
                    LlamarImprimir
                Else
                    MsgBox "No hay registros entre esos límites.", vbExclamation
                End If
            End If
            
        
        End If


End Sub

Private Sub CmdAcepInfAse_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim Fdesde As Date
Dim Fhasta As Date
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
    
    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H TRABAJADOR
    cDesde = Trim(txtCodigo(64).Text)
    cHasta = Trim(txtCodigo(65).Text)
    nDesde = txtNombre(64).Text
    nHasta = txtNombre(65).Text
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        nDesde = ""
        nHasta = ""
    End If
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{horas.codtraba}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador=""") Then Exit Sub
    End If
    
    Fdesde = CDate("01/" & Format(Combo1(1).ListIndex + 1, "00") & "/" & txtCodigo(61).Text)
    Fhasta = DateAdd("m", 1, Fdesde) - 1
    
    nDesde = ""
    nHasta = ""
    
    'D/H fecha
    cDesde = Trim(Fdesde)
    cHasta = Trim(Fhasta)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{horas.fechahora}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
        
    ConSubInforme = False


    'Nombre fichero .rpt a Imprimir
    indRPT = 60 ' informe de asesoria
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    
    cadNombreRPT = nomDocu '"rInfAsesoriaNomiMes.rpt"
    cadTitulo = "Informe para Asesoria Mensual"
    If Me.Check2.Value = 1 Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "1.rpt") '"rInfAsesoriaNomiMes1.rpt"

    If CargarTemporalListAsesoria(cadSelect, Fdesde, Fhasta) Then
        Tabla = "{tmpinformes}"
        cadSelect = "{tmpinformes.codusu} = " & vUsu.Codigo
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        
        CadParam = CadParam & "pDias=" & Day(Fhasta) & "|"
        numParam = numParam + 1
    Else
        Exit Sub
    End If

    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(Tabla, cadSelect) Then
        If (vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16) And Me.Check2.Value = 1 Then
            If vParamAplic.Cooperativa = 4 Then ' Alzira
                Shell App.Path & "\nomina.exe /E|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
            Else ' Picassent
                Shell App.Path & "\nomina.exe /P|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
            End If
        Else
            LlamarImprimir
        End If
    End If


End Sub

Private Sub CmdAcepInfComprob_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H TRABAJADOR
    cDesde = Trim(txtCodigo(49).Text)
    cHasta = Trim(txtCodigo(50).Text)
    nDesde = txtNombre(49).Text
    nHasta = txtNombre(50).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{horas.codtraba}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador=""") Then Exit Sub
    End If
    
    'D/H fecha
    cDesde = Trim(txtCodigo(44).Text)
    cHasta = Trim(txtCodigo(48).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{horas.fechahora}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
    
    Select Case OpcionListado
        Case 28 ' informe de comprobacion
            ConSubInforme = False
        
            cadNombreRPT = "rInfComprobNomi.rpt"
        
            indRPT = 84 ' personalizamos el informe de comprobacion
            
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    
            cadTitulo = "Informe de Comprobación Nóminas"
            
            cadNombreRPT = nomDocu
        
        Case 34 ' informe para asesoria Picassent
            ConSubInforme = False
        
            cadNombreRPT = "rInfAsesoriaNomi.rpt"
            cadTitulo = "Informe para Asesoria"
        
            If CargarTemporalPicassent(cadSelect) Then
                Tabla = "{tmpinformes}"
                cadSelect = "{tmpinformes.codusu} = " & vUsu.Codigo
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            Else
                Exit Sub
            End If
    End Select

    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(Tabla, cadSelect) Then
        LlamarImprimir
    End If

End Sub

Private Function CargarTemporalPicassent(cadWHERE As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Cad As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
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

Dim Anticipado As Currency

On Error GoTo eProcesarCambiosPicassent
    
    CargarTemporalPicassent = False
    
    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    Sql = "select count(distinct horas.codtraba) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Rs.Close
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "select horas.codtraba,  sum(horasdia), sum(compleme), sum(penaliza), sum(importe) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    Sql = Sql & " group by horas.codtraba "
    Sql = Sql & " order by 1 "
        
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & Rs!CodTraba & vbCrLf
        
        Sql2 = "select salarios.*, straba.dtoreten, straba.dtosegso, straba.dtosirpf, straba.pluscapataz from salarios, straba where straba.codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql2 = Sql2 & " and salarios.codcateg = straba.codcateg "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        ImpHoras = Round2(DBLet(Rs.Fields(1).Value, "N") * DBLet(Rs2!impsalar, "N"), 2)
                                    ' importe + pluscapataz + complemento - penalizacion
        ImpBruto = Round2(ImpHoras + DBLet(Rs.Fields(4).Value, "N") + DBLet(Rs2!PlusCapataz, "N") + DBLet(Rs.Fields(2).Value, "N") - DBLet(Rs.Fields(3).Value, "N"), 2)
                                                'codtraba,bruto,    anticipado,diferencia
        
        '[Monica]05/10/2010: el importe bruto es el que le he pagaria sin cargar ningun dto
        Sql5 = "select sum(importe) from rrecibosnomina where codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql5 = Sql5 & " and fechahora >= " & DBSet(txtCodigo(44).Text, "F")
        Sql5 = Sql5 & " and fechahora <= " & DBSet(txtCodigo(48).Text, "F")
        ImpBruto = DevuelveValor(Sql5)
        
        '[Monica]05/10/2010: el importe anticipado es el importe liquido (antes sum(importe) era incorrecto)
        Sql5 = "select sum(neto34) from rrecibosnomina where codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql5 = Sql5 & " and fechahora >= " & DBSet(txtCodigo(44).Text, "F")
        Sql5 = Sql5 & " and fechahora <= " & DBSet(txtCodigo(48).Text, "F")
                                                
        Anticipado = DevuelveValor(Sql5)
        Diferencia = ImpBruto - Anticipado
                                                
        Sql3 = "insert into tmpinformes (codusu, codigo1, importe1, importe2, importe3) values ("
        Sql3 = Sql3 & vUsu.Codigo & ","
        Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & ","
        Sql3 = Sql3 & DBSet(ImpBruto, "N") & ","
        Sql3 = Sql3 & DBSet(Anticipado, "N") & ","
        Sql3 = Sql3 & DBSet(Diferencia, "N") & ")"
        
        conn.Execute Sql3

        Set Rs2 = Nothing
            
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    CargarTemporalPicassent = True
    Exit Function
    
eProcesarCambiosPicassent:
    If Err.Number <> 0 Then
        Mens = Err.Description
        MsgBox "Error " & Mens, vbExclamation
    End If
End Function



Private Sub CmdAcepPaseBanco_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim Sql As String

    If Not DatosOk Then Exit Sub
    
    
    InicializarVbles
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If vParamAplic.Cooperativa = 9 Then
        '======== FORMULA  ====================================
        'D/H TRABAJADOR
        cDesde = Trim(txtCodigo(62).Text)
        cHasta = Trim(txtCodigo(63).Text)
        nDesde = txtNombre(49).Text
        nHasta = txtNombre(50).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{horasanticipos.codtraba}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador=""") Then Exit Sub
        End If
    
        'La forma de pago tiene que ser de tipo Transferencia
        AnyadirAFormula cadSelect, "forpago.tipoforp = 1"
        
        AnyadirAFormula cadSelect, "horasanticipos.fechapago is null"
     
     
        Tabla = "(horasanticipos INNER JOIN straba ON horasanticipos.codtraba = straba.codtraba) INNER JOIN forpago ON straba.codforpa = forpago.codforpa "
                   
        cTabla = Tabla
        
        cTabla = QuitarCaracterACadena(cTabla, "{")
        cTabla = QuitarCaracterACadena(cTabla, "}")
        Sql = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cadSelect <> "" Then
            cadSelect = QuitarCaracterACadena(cadSelect, "{")
            cadSelect = QuitarCaracterACadena(cadSelect, "}")
            cadSelect = QuitarCaracterACadena(cadSelect, "_1")
            Sql = Sql & " WHERE " & cadSelect
        End If
        
        If RegistrosAListar(Sql) = 0 Then
            MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        Else
            ProcesoPaseABancoAnticipos (cadSelect)
        End If
    
    
    Else
        '======== FORMULA  ====================================
        'D/H TRABAJADOR
        cDesde = Trim(txtCodigo(62).Text)
        cHasta = Trim(txtCodigo(63).Text)
        nDesde = txtNombre(49).Text
        nHasta = txtNombre(50).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rrecasesoria.codtraba}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador=""") Then Exit Sub
        End If
    
        'La forma de pago tiene que ser de tipo Transferencia
        AnyadirAFormula cadSelect, "forpago.tipoforp = 1"
        
        AnyadirAFormula cadSelect, "rrecasesoria.idconta = 0"
     
        Tabla = "(rrecasesoria INNER JOIN straba ON rrecasesoria.codtraba = straba.codtraba) INNER JOIN forpago ON straba.codforpa = forpago.codforpa "
                   
        cTabla = Tabla
        
        cTabla = QuitarCaracterACadena(cTabla, "{")
        cTabla = QuitarCaracterACadena(cTabla, "}")
        Sql = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cadSelect <> "" Then
            cadSelect = QuitarCaracterACadena(cadSelect, "{")
            cadSelect = QuitarCaracterACadena(cadSelect, "}")
            cadSelect = QuitarCaracterACadena(cadSelect, "_1")
            Sql = Sql & " WHERE " & cadSelect
        End If
        
        If RegistrosAListar(Sql) = 0 Then
            MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        Else
            ProcesoPaseABanco (cadSelect)
        End If
    
    End If
    
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim Prevision As Boolean

    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Select Case Index
        Case 0 ' Proceso de pago de partes de campo
            NomAlmac = ""
            NomAlmac = DevuelveDesdeBDNew(cAgro, "salmpr", "nomalmac", "codalmac", vParamAplic.AlmacenNOMI, "N")
            If NomAlmac = "" Then
                MsgBox "Debe introducir un código de almacén de Nóminas en parámetros. Revise.", vbExclamation
                Exit Sub
            End If
        
            'D/H Parte
            cDesde = Trim(txtCodigo(0).Text)
            cHasta = Trim(txtCodigo(1).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rpartes.nroparte}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHParte=""") Then Exit Sub
            End If
            
            'D/H Fecha
            cDesde = Trim(txtCodigo(14).Text)
            cHasta = Trim(txtCodigo(15).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rpartes.fechapar}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
    
            cTabla = Tabla & " INNER JOIN rpartes_trabajador ON rpartes.nroparte = rpartes_trabajador.nroparte "
    
            If HayRegParaInforme(cTabla, cadSelect) Then
                If vParamAplic.Cooperativa = 4 Then ' Alzira
                    '[Monica]23/12/2011: sólo en el caso de que queramos la prevision
                    If Check5.Value = 1 Then
                        If ProcesoCargaHoras(cTabla, cadSelect, True) Then
                            ConSubInforme = False
                            cadNombreRPT = "rPrevPagoPartes.rpt"
                            cadTitulo = "Informe de Prevision Pago de Partes"
                            
                            cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
                            cadSelect = cadFormula
                            'Comprobar si hay registros a Mostrar antes de abrir el Informe
                            If HayRegParaInforme("tmpinformes", cadSelect) Then
                                LlamarImprimir
                            End If
                        End If
                    Else
                        If ProcesoCargaHoras(cTabla, cadSelect, False) Then
                            MsgBox "Proceso realizado correctamente.", vbExclamation
                            cmdCancel_Click (0)
                            Exit Sub
                        Else
                            MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
                            Exit Sub
                        End If
                    End If
                Else
                    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then  ' Picassent
                        If ProcesoCargaHorasPicassent(cTabla, cadSelect) Then
                            MsgBox "Proceso realizado correctamente.", vbExclamation
                            cmdCancel_Click (0)
                            Exit Sub
                        Else
                            MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
                            Exit Sub
                        End If
                    Else
                        '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
                        '                    Natural no tiene partes
                        If vParamAplic.Cooperativa = 0 Then ' catadau
                            If ProcesoCargaHorasCatadau(cTabla, cadSelect) Then
                                MsgBox "Proceso realizado correctamente.", vbExclamation
                                cmdCancel_Click (0)
                                Exit Sub
                            Else
                                MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
                                Exit Sub
                            End If
                        End If
                    End If
                End If
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

            CadParam = CadParam & "pHProductivas=" & Me.Check3.Value & "|"
            numParam = numParam + 1
            
            ConSubInforme = False
            cadNombreRPT = "rManHorasTrab.rpt"
            cadTitulo = "Informe de Horas Trabajadas"
            
            'Comprobar si hay registros a Mostrar antes de abrir el Informe
            If HayRegParaInforme(Tabla, cadSelect) Then
                LlamarImprimir
            End If
    
    
        Case 1 ' informe de horas destajo trabajadas
            '======== FORMULA  ====================================
            'D/H TRABAJADOR
            cDesde = Trim(txtCodigo(2).Text)
            cHasta = Trim(txtCodigo(3).Text)
            nDesde = txtNombre(2).Text
            nHasta = txtNombre(3).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{horasdestajo.codtraba}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador=""") Then Exit Sub
            End If
            
            'D/H fecha
            cDesde = Trim(txtCodigo(4).Text)
            cHasta = Trim(txtCodigo(5).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{horasdestajo.fechahora}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
            
            
            
            Select Case OpcionListado
                Case 18 ' informe de horas de destajo trabajadas
                    ConSubInforme = False
                
                    If Me.Check1.Value Then
                        cadNombreRPT = "rManHorasTrabDestajo.rpt"
                        cadTitulo = "Informe de Horas Destajo para trabajador"
                    Else
                        cadNombreRPT = "rManHorasDestajo.rpt"
                        cadTitulo = "Informe de Horas Destajo para trabajador"
                    End If
            
                    'Comprobar si hay registros a Mostrar antes de abrir el Informe
                    If HayRegParaInforme(Tabla, cadSelect) Then
                        LlamarImprimir
                    End If
                Case 19 ' actualizacion de horas de destajo al  fichero de horas
                    If ActualizarTabla(Tabla, cadSelect) Then
                        MsgBox "Proceso realizado correctamente.", vbExclamation
                        cmdCancel_Click (1)
                    Else
                        MsgBox "No se ha realizado el proceso. Llame a Ariadna.", vbExclamation
                    End If
                    DesBloqueoManual ("ACTDES") 'ACTualizacion DEStajo

            End Select
    End Select
    

End Sub


Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView1
End Sub

Private Sub CmdAcepTrabajCapataz_Click()
Dim Sql As String
Dim CodigoETT As String

    If txtCodigo(47).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo Variedad.", vbExclamation
        Exit Sub
    Else
        txtNombre(47).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", txtCodigo(47).Text, "N")
        If txtNombre(47).Text = "" Then
            MsgBox "No existe la variedad. Revise.", vbExclamation
            PonerFoco txtCodigo(47)
            Exit Sub
        End If
    End If
    
    If txtCodigo(46).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en la Fecha.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(45).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo capataz.", vbExclamation
        Exit Sub
    End If
    
    If CalculoTrabajCapatazNew() Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
       
        cmdCancel_Click (2)
    End If


End Sub

Private Sub CmdAltaRapida_Click()
Dim Sql As String
Dim CodigoETT As String

    If txtCodigo(36).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo Variedad.", vbExclamation
        Exit Sub
    Else
        txtNombre(36).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", txtCodigo(36).Text, "N")
        If txtNombre(36).Text = "" Then
            MsgBox "No existe la variedad. Revise.", vbExclamation
            PonerFoco txtCodigo(36)
            Exit Sub
        End If
    End If
    
    If txtCodigo(35).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en la Fecha desde.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(26).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en la Fecha hasta.", vbExclamation
        Exit Sub
    End If
    
    If CDate(txtCodigo(35).Text) > CDate(txtCodigo(26).Text) Then
        MsgBox "La fecha desde no puede ser superior a la fecha hasta.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(34).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo capataz.", vbExclamation
        Exit Sub
    End If
    
    If CalculoAltaRapida() Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
       
        cmdCancel_Click (2)
    End If

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

Private Sub CmdEventuales_Click()
    If txtCodigo(28).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo Variedad.", vbExclamation
        Exit Sub
    Else
        txtNombre(28).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", txtCodigo(28).Text, "N")
        If txtNombre(28).Text = "" Then
            MsgBox "No existe la variedad. Revise.", vbExclamation
            PonerFoco txtCodigo(28)
            Exit Sub
        End If
    End If
    
    If txtCodigo(37).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en la Fecha desde.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(33).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en la Fecha hasta.", vbExclamation
        Exit Sub
    End If
    
    If CDate(txtCodigo(37).Text) > CDate(txtCodigo(33).Text) Then
        MsgBox "La fecha desde no puede ser superior a la fecha hasta.", vbExclamation
        Exit Sub
    End If
    
    
    If txtCodigo(41).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el Trabajador desde.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(42).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el Trabajador hasta.", vbExclamation
        Exit Sub
    End If
    
    If CDate(txtCodigo(41).Text) > CDate(txtCodigo(42).Text) Then
        MsgBox "El código desde no puede ser superior al código hasta.", vbExclamation
        Exit Sub
    End If
    
    If CalculoEventuales() Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
       
        cmdCancel_Click (2)
    End If

End Sub



Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub CmdDiasTrabajados_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim Fdesde As Date
Dim Fhasta As Date
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
    
Dim cadSelect2 As String

    InicializarVbles
    
    cadSelect2 = "(1=1)"
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H TRABAJADOR
    cDesde = Trim(txtCodigo(68).Text)
    cHasta = Trim(txtCodigo(69).Text)
    nDesde = txtNombre(68).Text
    nHasta = txtNombre(69).Text
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        nDesde = ""
        nHasta = ""
    End If
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpartes_trabajador.codtraba}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador=""") Then Exit Sub
        
        cadSelect2 = Replace(cadSelect, "rpartes_trabajador", "horas")
    End If
    
    Fdesde = CDate("01/" & Format(Combo1(2).ListIndex + 1, "00") & "/" & txtCodigo(67).Text)
    Fhasta = DateAdd("m", 1, Fdesde) - 1
    
    nDesde = ""
    nHasta = ""
    
    'D/H fecha
    cDesde = Trim(Fdesde)
    cHasta = Trim(Fhasta)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpartes.fecentrada}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        
        If cDesde <> "" Then cadSelect2 = cadSelect2 & " and horas.fechahora >= " & DBSet(cDesde, "F")
        If cHasta <> "" Then cadSelect2 = cadSelect2 & " and horas.fechahora <= " & DBSet(cHasta, "F")
    End If
    
        
    ConSubInforme = False


    'Nombre fichero .rpt a Imprimir
    indRPT = 110 ' informe de dias trabajados
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    
    cadNombreRPT = nomDocu '"rInfAsesoriaNomiMes.rpt"
    cadTitulo = "Informe Mensual Días Trabajados"
    If Me.Check2.Value = 1 Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "1.rpt") '"rInfAsesoriaNomiMes1.rpt"

    If CargarTemporalListDiasTrabajados(cadSelect, Fdesde, Fhasta, cadSelect2) Then
        Tabla = "{tmpinformes}"
        cadSelect = "{tmpinformes.codusu} = " & vUsu.Codigo
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        
        CadParam = CadParam & "pDias=" & Day(Fhasta) & "|"
        numParam = numParam + 1
    Else
        Exit Sub
    End If

    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(Tabla, cadSelect) Then
        If (vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16) And Me.Check6.Value = 1 Then
            If vParamAplic.Cooperativa = 4 Then ' Alzira
                Shell App.Path & "\nomina.exe /E|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
            Else ' Picassent
                Shell App.Path & "\nomina.exe /P|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
            End If
        Else
            LlamarImprimir
        End If
    End If



End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 15 ' Informe de Horas Trabajadas
                PonerFoco txtCodigo(18)
                
            Case 16 ' calculo de horas productivas
                PonerFoco txtCodigo(27)
            
            Case 18, 19 ' 18 = informe de horas destajo
                        ' 19 = actualizacion de horas
                PonerFoco txtCodigo(2)
                
                If OpcionListado = 19 Then Label3.Caption = "Actualización Entradas de Destajo"
                
            Case 20, 21, 22, 30, 31, 32 '20,21,22= Horas ETT
                                        '30,31,32 = Horas
                PonerFoco txtCodigo(9)
                Select Case OpcionListado
                    Case 20, 30
                        Me.FrameDestajo.visible = True
                        Me.FramePenalizacion.visible = False
                        Me.FrameBonificacion.visible = False
                        Me.FrameDestajo.Enabled = True
                        Me.FramePenalizacion.Enabled = False
                        Me.FrameBonificacion.Enabled = False
                        
                    Case 21, 31
                        Me.FrameDestajo.visible = False
                        Me.FramePenalizacion.visible = True
                        Me.FrameBonificacion.visible = False
                        Me.FrameDestajo.Enabled = False
                        Me.FramePenalizacion.Enabled = True
                        Me.FrameBonificacion.Enabled = False
                        Label4.Caption = "Calculo Penalización"
                        
                    Case 22, 32
                        Me.FrameDestajo.visible = False
                        Me.FramePenalizacion.visible = False
                        Me.FrameBonificacion.visible = True
                        Me.FrameDestajo.Enabled = False
                        Me.FramePenalizacion.Enabled = False
                        Me.FrameBonificacion.Enabled = True
                        Label4.Caption = "Calculo Bonificación"
                            
                End Select
                
            Case 23, 27, 33 ' 23 borrado masivo de horas ett
                            ' 27 borrado masivo de horas
                            ' 33 borrado masivo de horas
                PonerFoco txtCodigo(31)
                
            Case 24 ' alta rapida
                PonerFoco txtCodigo(36)
                
            Case 25 ' eventuales
                PonerFoco txtCodigo(28)
            
            Case 26 ' trabajadores de un capataz
                PonerFoco txtCodigo(47)
                
            Case 28 ' Informe de comprobacion para picassent
                PonerFoco txtCodigo(49)
                
            Case 29 ' Listado de entradas capataz
                PonerFoco txtCodigo(38)
        
            Case 34 ' Informe para asesoria
                PonerFoco txtCodigo(49)
                
            Case 35 ' Borrado Masivo de Registros Asesoria
                PonerFoco txtCodigo(54)
                
            Case 36 ' Pase a banco de importes
                Combo1(0).ListIndex = 0
                txtCodigo(59).Text = Format(Now, "dd/mm/yyyy")
                txtCodigo(60).Text = Format(Now, "dd/mm/yyyy")
                PonerFoco txtCodigo(62)
                
                '[Monica]18/09/2013: anticipos para Natural
                FrameConcep.visible = (vParamAplic.Cooperativa = 9)
                FrameConcep.Enabled = (vParamAplic.Cooperativa = 9)
                If vParamAplic.Cooperativa = 9 Then
                    Label2(77).Caption = "Fecha"
                    txtCodigo(66).Text = "ANTICIPO " & UCase(MonthName(Month(Now))) & " " & Year(Now)
                End If
                
            Case 37 ' Informe de horas mensual para asesoria
                PonerFoco txtCodigo(64)
                
                txtCodigo(61).Text = Format(Year(Now), "0000")
                
                PosicionarCombo Combo1(1), Month(Now)
                
            Case 38 ' Informe de rendimiento por capataz
                txtCodigo(52).Text = Format(Now, "dd/mm/yyyy")
                txtCodigo(53).Text = txtCodigo(52).Text
                
            Case 39 ' Listado de horas trabajadas
                PonerFoco txtCodigo(68)
                
                txtCodigo(67).Text = Format(Year(Now), "0000")
                
                PosicionarCombo Combo1(2), Month(Now)
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
    
    For H = 0 To 30
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
   
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    Me.FrameHorasTrabajadas.visible = False
    Me.FrameCalculoHorasProductivas.visible = False
    Me.FramePagoPartesCampo.visible = False
    Me.FrameHorasDestajo.visible = False
    Me.FrameCalculoETT.visible = False
    Me.FrameBorradoMasivoETT.visible = False
    Me.FrameAltaRapida.visible = False
    Me.FrameEventuales.visible = False
    Me.FrameTrabajadoresCapataz.visible = False
    Me.FrameInfComprobacion.visible = False
    Me.FrameEntradasCapataz.visible = False
    Me.FrameBorradoAsesoria.visible = False
    Me.FramePaseABanco.visible = False
    Me.FrameListMensAsesoria.visible = False
    Me.FrameInfDiasTrabajados.visible = False
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 15 ' Informe de Horas Trabajadas
        FrameHorasTrabajadasVisible True, H, W
        indFrame = 0
        Tabla = "horas"
        
    Case 16 ' Proceso de Calculo de Horas Productivas
        FrameCalculoHorasProductivasVisible True, H, W
        indFrame = 0
        Tabla = "horas"
        
    Case 17 ' Proceso de Pago de Partes de Campo
        FramePagoPartesCampoVisible True, H, W
        indFrame = 0
        Tabla = "rpartes"
    
        '[Monica]23/12/2011: solo Alzira puede sacar la prevision de pago de partes
        Frame1.visible = (vParamAplic.Cooperativa = 4)
        Frame1.Enabled = (vParamAplic.Cooperativa = 4)
    
    
    Case 18 ' Informe de Horas Trabajadas destajo
        FrameHorasDestajoVisible True, H, W
        indFrame = 0
        Tabla = "horasdestajo"
    
        Check1.visible = True
        Check1.Enabled = True
        
    Case 19 ' Actualizar horas de destajo ( pasa a la tabla de horas )
        FrameHorasDestajoVisible True, H, W
        indFrame = 0
        Tabla = "horasdestajo"
    
        Check1.visible = False
        Check1.Enabled = False
    
    Case 20, 30 ' Horas ETT
        FrameHorasETTVisible True, H, W
        indFrame = 0
        If OpcionListado = 20 Then
            Tabla = "horasett"
        Else
            Tabla = "horas"
        End If
    
    Case 21, 31 ' Penalizacion ett
        FrameHorasETTVisible True, H, W
        indFrame = 0
        If OpcionListado = 21 Then
            Tabla = "horasett"
        Else
            Tabla = "horas"
        End If
    
    Case 22, 32 ' Bonificacion
        FrameHorasETTVisible True, H, W
        indFrame = 0
        If OpcionListado = 22 Then
            Tabla = "horasett"
        Else
            Tabla = "horas"
        End If
    
    Case 23, 33 ' Borrado Masivo ETT
        FrameBorradoMasivoETTVisible True, H, W
        indFrame = 0
        Select Case OpcionListado
            Case 23
                Tabla = "horasett"
            Case 33
                Tabla = "horas"
        End Select
        
    Case 24 ' alta rapida
        FrameAltaRapidaVisible True, H, W
        indFrame = 0
        Tabla = "horas"
        
    Case 25 ' eventuales
        FrameEventualesVisible True, H, W
        indFrame = 0
        Tabla = "horas"
    
    Case 26 ' trabaajdores de un capataz
        FrameTrabajadoresCapatazVisible True, H, W
        indFrame = 0
        Tabla = "horas"
    
    Case 27 ' Borrado Masivo Horas
        Label5.Caption = "Borrado Masivo Horas"
        FrameBorradoMasivoETTVisible True, H, W
        indFrame = 0
        Tabla = "horas"
        
    Case 28 ' Informe de Comprobacion
        FrameInfComprobacionVisible True, H, W
        indFrame = 0
        Tabla = "horas"
    
    Case 29 ' Informe de Entradas Capataz
        FrameEntradasCapatazVisible True, H, W
        indFrame = 0
        Tabla = "horas"
    
    Case 34 ' Informe para Asesoria
        FrameInfComprobacionVisible True, H, W
        indFrame = 0
        Tabla = "horas"
        Label11.Caption = "Informe para Asesoria"
    
    Case 35 ' Borrado masivo Asesoria
        FrameBorradoAsesoriaVisible True, H, W
        indFrame = 0
        Tabla = "rrecasesoria"
    
    Case 36 ' pase a banco
        CargaCombo
    
        FramePaseaBancoVisible True, H, W
        indFrame = 0
        Tabla = "rrecasesoria"
    
    Case 37 ' Informe de horas mensual para asesoria
        CargaCombo
    
        FrameListMensAsesoriaVisible True, H, W
        indFrame = 0
        Tabla = "rrecasesoria"
    
    Case 38 ' Rendimiento por Capataz
        Label12.Caption = "Rendimiento por Capataz"
        FrameEntradasCapatazVisible True, H, W
        Check4.visible = False
        Check4.Enabled = False
        
        indFrame = 0
        Tabla = "horas"

    Case 39 ' Informe de dias trabajados
        CargaCombo
    
        FrameInfDiasTrabajadosVisible True, H, W
        indFrame = 0
        Tabla = "rpartes"



    End Select
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.CmdCancel(0).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub



Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFecha(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCap_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 ' trabajadores
            AbrirFrmManTraba (Index + 2)
        
        Case 2, 3, 5 'variedades
            AbrirFrmVariedades Index + 4

        Case 7 ' variedades
            AbrirFrmVariedades Index + 29

        Case 6 'capataz
            AbrirFrmCapataces Index + 6

        Case 8, 9 'capataz
            AbrirFrmCapataces Index + 23

        Case 4 ' capataz
            AbrirFrmCapataces Index + 30

        Case 14, 15 'trabajadores
            AbrirFrmManTraba (Index + 4)
    
        Case 20
            AbrirFrmManAlmac (Index)
           
        Case 11 ' variedades
            AbrirFrmVariedades Index + 17
        
        Case 12, 13 'trabajadores
            AbrirFrmManTraba (Index + 29)
           
        Case 16 ' variedades
            AbrirFrmVariedades Index + 31
        
        Case 10 'capataz
            AbrirFrmCapataces Index + 35
            
        Case 19 'trabajadores
            AbrirFrmManTraba (49)
        
        Case 21 'trabajadores
            AbrirFrmManTraba (50)
    
        Case 17 'capataz
            AbrirFrmCapataces 38
        
        Case 18 'capataz
            AbrirFrmCapataces 43
            
        Case 22, 23 'trabajadores
            AbrirFrmManTraba (Index + 32)
        
        Case 25, 26 'trabajadores
            AbrirFrmManTraba (Index + 37)
    
        Case 24 ' banco
            AbrirFrmManBanco (Index + 34)
        
        Case 27, 28 ' trabajadores
            AbrirFrmManTraba (Index + 37)
        
        Case 29, 30 ' trabajadores
            AbrirFrmManTraba (Index + 39)
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object
    Dim indice As Integer
    
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
        Case 0, 1, 13, 2, 3
            indice = Index + 14
        Case 4, 5
            indice = Index
        Case 7
            indice = 11
        Case 6
            indice = 29
        Case 8
            indice = 30
        Case 9
            indice = 35
        Case 10
            indice = 26
        Case 12
            indice = 37
        Case 11
            indice = 33
        Case 14
            indice = 46
        Case 15
            indice = 44
        Case 16
            indice = 48
        Case 17, 18
            indice = Index + 35
        Case 19, 20
            indice = Index + 37
        Case 21, 22
            indice = Index + 38
    End Select
    imgFecha(0).Tag = indice '<===
    
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(indice).Text <> "" Then frmC.NovaData = txtCodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFecha(0).Tag)) '<===
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
            Case 2: KEYBusqueda KeyAscii, 0 'trabajador desde
            Case 3: KEYBusqueda KeyAscii, 1 'trabajador hasta
            Case 6:  KEYBusqueda KeyAscii, 2 'variedad desde
            Case 7:  KEYBusqueda KeyAscii, 3 'variedad hasta
            
            Case 18: KEYBusqueda KeyAscii, 14 'trabajador desde
            Case 19: KEYBusqueda KeyAscii, 15 'trabajador hasta
            
            Case 14: KEYFecha KeyAscii, 0 'fecha desde
            Case 15: KEYFecha KeyAscii, 1 'fecha hasta
            
            Case 24: KEYBusqueda KeyAscii, 20 'almacen para el calculo de horas productivas
        
            Case 9:  KEYBusqueda KeyAscii, 5 ' variedad
            Case 11: KEYFecha KeyAscii, 7 ' fecha
            Case 12: KEYBusqueda KeyAscii, 6 'capataz
        
            Case 35: KEYFecha KeyAscii, 9 ' fecha desde
            Case 26: KEYFecha KeyAscii, 10 ' fecha hasta
            
            Case 34:  KEYBusqueda KeyAscii, 4 'capataz
            Case 36: KEYBusqueda KeyAscii, 7 ' variedad
            
        
            Case 31: KEYBusqueda KeyAscii, 8 'capataz desde
            Case 32: KEYBusqueda KeyAscii, 9 'capataz hasta
            Case 29: KEYFecha KeyAscii, 6 'fecha desde
            Case 30: KEYFecha KeyAscii, 8 'fecha hasta
        
            Case 28:  KEYBusqueda KeyAscii, 11 ' variedad
            Case 37: KEYFecha KeyAscii, 12 ' fecha desde
            Case 33: KEYFecha KeyAscii, 11 ' fecha hasta
            Case 41: KEYBusqueda KeyAscii, 12 'trabajador desde
            Case 42: KEYBusqueda KeyAscii, 13 'trabajador hasta
        
            Case 47:  KEYBusqueda KeyAscii, 16 ' variedad
            Case 46: KEYFecha KeyAscii, 14 ' fecha desde
            Case 45: KEYBusqueda KeyAscii, 10 'capataz
        
            Case 44: KEYFecha KeyAscii, 15 ' fecha desde
            Case 48: KEYFecha KeyAscii, 16 ' fecha hasta
            Case 49: KEYBusqueda KeyAscii, 19 'trabajador desde
            Case 50: KEYBusqueda KeyAscii, 21 'trabajador hasta
        
            Case 38: KEYBusqueda KeyAscii, 17 'capataz desde
            Case 43: KEYBusqueda KeyAscii, 18 'capataz hasta
            Case 52: KEYFecha KeyAscii, 17 ' fecha desde
            Case 53: KEYFecha KeyAscii, 18 ' fecha hasta
        
            Case 54: KEYBusqueda KeyAscii, 22 'trabajador desde
            Case 55: KEYBusqueda KeyAscii, 23 'trabajador hasta
            Case 56: KEYFecha KeyAscii, 19 ' fecha desde
            Case 57: KEYFecha KeyAscii, 20 ' fecha hasta
            
            ' Pase a bancos
            Case 62: KEYBusqueda KeyAscii, 25 'trabajador desde
            Case 63: KEYBusqueda KeyAscii, 26 'trabajador hasta
            Case 59: KEYFecha KeyAscii, 21 ' fecha
            Case 60: KEYFecha KeyAscii, 22 ' fecha
            Case 58: KEYBusqueda KeyAscii, 24 'banco
        
            Case 64: KEYBusqueda KeyAscii, 27 'trabajador desde
            Case 65: KEYBusqueda KeyAscii, 28 'trabajador hasta
        
            Case 68: KEYBusqueda KeyAscii, 29 'trabajador desde
            Case 69: KEYBusqueda KeyAscii, 30 'trabajador hasta
        
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
        Case 0, 1 ' Nro.Partes
            PonerFormatoEntero txtCodigo(Index)
    
        Case 4, 5, 14, 15, 16, 17, 27, 11, 29, 30 'FECHAS
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoFecha txtCodigo(Index)
                If Index = 11 Then
                    Select Case OpcionListado
                        Case 20
                            CalculoDestajoETT False
                        Case 21
                            CalculoPenalizacionETT False
                        Case 30
                            CalculoDestajo False
                        Case 31
                            CalculoPenalizacion False
                    End Select
                End If
            End If
            
        Case 35, 26, 33, 37, 46, 44, 48, 52, 53, 56, 57, 59, 60 'FECHAS
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoFecha txtCodigo(Index)
            End If
         
        Case 18, 19, 2, 3, 41, 42, 49, 50, 54, 55, 62, 63, 64, 65, 68, 69 'TRABAJADORES
            PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "straba", "nomtraba", "codtraba", "N")
            
        Case 6, 7, 9 'VARIEDADES
            PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If Index = 9 And txtCodigo(Index).Text <> "" Then
                Select Case OpcionListado
                    Case 20
                        CalculoDestajoETT False
                    Case 21
                        CalculoPenalizacionETT False
                    Case 30
                        CalculoDestajo False
                    Case 31
                        CalculoPenalizacion False
                End Select
            End If
             
        Case 36, 28, 47 'VARIEDADES
            PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            
            
        Case 12 'CAPATAZ
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rcapataz", "nomcapat", "codcapat", "N")
            If txtCodigo(Index).Text <> "" Then
                Select Case OpcionListado
                    Case 20
                        CalculoDestajoETT False
                    Case 21
                        CalculoPenalizacionETT False
                        PonerFoco txtCodigo(21)
                    Case 30
                        CalculoDestajo False
                    Case 31
                        CalculoPenalizacion False
                        PonerFoco txtCodigo(21)
                    Case 22
                        PonerFoco txtCodigo(23)
                End Select
            End If
            
        Case 31, 32, 34, 45, 38, 43 'CAPATAZ
            PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rcapataz", "nomcapat", "codcapat", "N")
            
        Case 25 ' porcentaje
            If txtCodigo(Index).Text <> "" Then
                 PonerFormatoDecimal txtCodigo(Index), 9
            End If

        Case 24 'ALMACEN
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "salmpr", "nomalmac", "codalmac", "N")
            
            
        Case 21 ' porcentaje de penalizacion
            If PonerFormatoDecimal(txtCodigo(21), 4) Then
                If OpcionListado = 21 Then
                    CalculoPenalizacionETT False
                Else
                    CalculoPenalizacion False
                End If
                CmdAcepCalculoETT.SetFocus
            End If
            
        Case 23 ' bonificacion
            If PonerFormatoDecimal(txtCodigo(23), 4) Then
                CmdAcepCalculoETT.SetFocus
            End If
        
        Case 39, 40, 51 ' Importe
            If txtCodigo(Index).Text <> "" Then
                 PonerFormatoDecimal txtCodigo(Index), 3
            End If
        
        Case 58 'BANCO
            PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "banpropi", "nombanpr", "codbanpr", "N")
        
    End Select
End Sub


Private Sub FrameHorasTrabajadasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
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

Private Sub FrameHorasDestajoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameHorasDestajo.visible = visible
    If visible = True Then
        Me.FrameHorasDestajo.Top = -90
        Me.FrameHorasDestajo.Left = 0
        Me.FrameHorasDestajo.Height = 5565
        Me.FrameHorasDestajo.Width = 7425
        W = Me.FrameHorasDestajo.Width
        H = Me.FrameHorasDestajo.Height
    End If
End Sub


Private Sub FrameHorasETTVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameCalculoETT.visible = visible
    If visible = True Then
        Me.FrameCalculoETT.Top = -90
        Me.FrameCalculoETT.Left = 0
        Me.FrameCalculoETT.Height = 5055
        Me.FrameCalculoETT.Width = 6375
        W = Me.FrameCalculoETT.Width
        H = Me.FrameCalculoETT.Height
    End If
End Sub

Private Sub FrameAltaRapidaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameAltaRapida.visible = visible
    If visible = True Then
        Me.FrameAltaRapida.Top = -90
        Me.FrameAltaRapida.Left = 0
        Me.FrameAltaRapida.Height = 5055
        Me.FrameAltaRapida.Width = 6375
        W = Me.FrameAltaRapida.Width
        H = Me.FrameAltaRapida.Height
    End If
End Sub

Private Sub FrameEventualesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameEventuales.visible = visible
    If visible = True Then
        Me.FrameEventuales.Top = -90
        Me.FrameEventuales.Left = 0
        Me.FrameEventuales.Height = 5535
        Me.FrameEventuales.Width = 6375
        W = Me.FrameEventuales.Width
        H = Me.FrameEventuales.Height
    End If
End Sub


Private Sub FrameTrabajadoresCapatazVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameTrabajadoresCapataz.visible = visible
    If visible = True Then
        Me.FrameTrabajadoresCapataz.Top = -90
        Me.FrameTrabajadoresCapataz.Left = 0
        Me.FrameTrabajadoresCapataz.Height = 5055
        Me.FrameTrabajadoresCapataz.Width = 6375
        W = Me.FrameTrabajadoresCapataz.Width
        H = Me.FrameTrabajadoresCapataz.Height
    End If
End Sub

Private Sub FrameInfComprobacionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameInfComprobacion.visible = visible
    If visible = True Then
        Me.FrameInfComprobacion.Top = -90
        Me.FrameInfComprobacion.Left = 0
        Me.FrameInfComprobacion.Height = 4425
        Me.FrameInfComprobacion.Width = 6915
        W = Me.FrameInfComprobacion.Width
        H = Me.FrameInfComprobacion.Height
    End If
End Sub

Private Sub FrameEntradasCapatazVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de entradas capataz
    Me.FrameEntradasCapataz.visible = visible
    If visible = True Then
        Me.FrameEntradasCapataz.Top = -90
        Me.FrameEntradasCapataz.Left = 0
        Me.FrameEntradasCapataz.Height = 4425
        Me.FrameEntradasCapataz.Width = 6915
        W = Me.FrameEntradasCapataz.Width
        H = Me.FrameEntradasCapataz.Height
    End If
End Sub


Private Sub FrameBorradoAsesoriaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de entradas capataz
    Me.FrameBorradoAsesoria.visible = visible
    If visible = True Then
        Me.FrameBorradoAsesoria.Top = -90
        Me.FrameBorradoAsesoria.Left = 0
        Me.FrameBorradoAsesoria.Height = 4215
        Me.FrameBorradoAsesoria.Width = 6705
        W = Me.FrameBorradoAsesoria.Width
        H = Me.FrameBorradoAsesoria.Height
    End If
End Sub

Private Sub FrameBorradoMasivoETTVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameBorradoMasivoETT.visible = visible
    If visible = True Then
        Me.FrameBorradoMasivoETT.Top = -90
        Me.FrameBorradoMasivoETT.Left = 0
        Me.FrameBorradoMasivoETT.Height = 3885
        Me.FrameBorradoMasivoETT.Width = 6285
        W = Me.FrameBorradoMasivoETT.Width
        H = Me.FrameBorradoMasivoETT.Height
    End If
End Sub


Private Sub FrameCalculoHorasProductivasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el calculo de horas productivas
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

Private Sub FramePagoPartesCampoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el calculo de horas productivas
    Me.FramePagoPartesCampo.visible = visible
    If visible = True Then
        Me.FramePagoPartesCampo.Top = -90
        Me.FramePagoPartesCampo.Left = 0
        Me.FramePagoPartesCampo.Height = 4455
        Me.FramePagoPartesCampo.Width = 6345
        W = Me.FramePagoPartesCampo.Width
        H = Me.FramePagoPartesCampo.Height
    End If
End Sub


Private Sub FramePaseaBancoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el pase a banco
    Me.FramePaseABanco.visible = visible
    If visible = True Then
        Me.FramePaseABanco.Top = -90
        Me.FramePaseABanco.Left = 0
        Me.FramePaseABanco.Height = 5490 '5130
        Me.FramePaseABanco.Width = 6435
        W = Me.FramePaseABanco.Width
        H = Me.FramePaseABanco.Height
    End If
End Sub


Private Sub FrameListMensAsesoriaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el pase a banco
    Me.FrameListMensAsesoria.visible = visible
    If visible = True Then
        Me.FrameListMensAsesoria.Top = -90
        Me.FrameListMensAsesoria.Left = 0
        Me.FrameListMensAsesoria.Height = 4275
        Me.FrameListMensAsesoria.Width = 6375
        W = Me.FrameListMensAsesoria.Width
        H = Me.FrameListMensAsesoria.Height
    End If
End Sub


Private Sub FrameInfDiasTrabajadosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el pase a banco
    Me.FrameInfDiasTrabajados.visible = visible
    If visible = True Then
        Me.FrameInfDiasTrabajados.Top = -90
        Me.FrameInfDiasTrabajados.Left = 0
        Me.FrameInfDiasTrabajados.Height = 4275
        Me.FrameInfDiasTrabajados.Width = 6375
        W = Me.FrameInfDiasTrabajados.Width
        H = Me.FrameInfDiasTrabajados.Height
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
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .ConSubInforme = ConSubInforme
        .Opcion = OpcionListado
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmManTraba(indice As Integer)
    indCodigo = indice
    Set frmTra = New frmManTraba
    frmTra.DatosADevolverBusqueda = "0|2|"
    frmTra.Show vbModal
    Set frmTra = Nothing
End Sub

Private Sub AbrirFrmManCapataz(indice As Integer)
    indCodigo = indice
    Set frmCap = New frmManCapataz
    frmCap.DatosADevolverBusqueda = "0|1|"
    frmCap.Show vbModal
    Set frmCap = Nothing
End Sub

Private Sub AbrirFrmManBanco(indice As Integer)
    indCodigo = indice
    
    Set frmBan = New frmComercial
    
    AyudaBancosCom frmBan, txtCodigo(indCodigo)
    
    Set frmBan = Nothing
    
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub AbrirFrmManAlmac(indice As Integer)
    indCodigo = indice + 4
    
    Set frmAlm = New frmComercial
    
    AyudaAlmacenCom frmAlm, txtCodigo(indCodigo).Text
    
    Set frmAlm = Nothing
    
    PonerFoco txtCodigo(indCodigo)

End Sub


Private Function CargarTablaTemporal() As Boolean
Dim Sql As String
Dim Sql1 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarTablaTemporal
    
    CargarTablaTemporal = False

    Sql = "delete from tmpenvasesret where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql

'select albaran_envase.codartic, albaran_envase.fechamov
'from (albaran_envase inner join sartic on albaran_envase.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar
'Where stipar.esretornable = 1
'Union
'select smoval.codartic, smoval.fechamov
'from (smoval inner join  sartic on smoval.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar
'Where stipar.esretornable = 1

    Sql = "select " & vUsu.Codigo & ", albaran_envase.codartic, albaran_envase.fechamov, albaran_envase.cantidad, albaran_envase.tipomovi, albaran_envase.numalbar, "
    Sql = Sql & " albaran_envase.codclien, clientes.nomclien, " & DBSet("ALV", "T")
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
    
    Sql = Sql & " union "
    
    Sql = Sql & "select " & vUsu.Codigo & ", smoval.codartic, smoval.fechamov, smoval.cantidad, smoval.tipomovi, smoval.document, "
    Sql = Sql & " smoval.codigope, proveedor.nomprove, " & DBSet("ALC", "T")
    Sql = Sql & " from ((smoval inner join sartic on smoval.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar) "
    Sql = Sql & " inner join proveedor on smoval.codigope = proveedor.codprove "
    Sql = Sql & " where stipar.esretornable = 1 "
    
    If txtCodigo(12).Text <> "" Then Sql = Sql & " and stipar.codtipar >= " & DBSet(txtCodigo(12).Text, "N")
    If txtCodigo(13).Text <> "" Then Sql = Sql & " and stipar.codtipar <= " & DBSet(txtCodigo(13).Text, "N")
    
    If txtCodigo(20).Text <> "" Then Sql = Sql & " and smoval.codartic >= " & DBSet(txtCodigo(20).Text, "T")
    If txtCodigo(21).Text <> "" Then Sql = Sql & " and smoval.codartic <= " & DBSet(txtCodigo(21).Text, "T")
    
    If txtCodigo(14).Text <> "" Then Sql = Sql & " and smoval.fechamov >= " & DBSet(txtCodigo(14).Text, "F")
    If txtCodigo(15).Text <> "" Then Sql = Sql & " and smoval.fechamov <= " & DBSet(txtCodigo(15).Text, "F")

    Sql1 = "insert into tmpenvasesret " & Sql
    conn.Execute Sql1
    
    CargarTablaTemporal = True
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Carga Tabla Temporal"
End Function

Private Function CalculoHorasProductivas() As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String

    On Error GoTo eCalculoHorasProductivas

    CalculoHorasProductivas = False

    Sql = "fechahora = " & DBSet(txtCodigo(27).Text, "F") & " and codalmac = " & DBSet(txtCodigo(24), "N")
    Sql = Sql & " and codtraba in (select codtraba from straba where codsecci = 1)"


    If BloqueaRegistro("horas", Sql) Then
        Sql1 = "update horas set horasproduc = round(horasdia * (1 + (" & DBSet(txtCodigo(25), "N") & "/ 100)),2) "
        Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(27).Text, "F")
        Sql1 = Sql1 & " and codalmac = " & DBSet(txtCodigo(24), "N")
        Sql1 = Sql1 & " and codtraba in (select codtraba from straba where codsecci = 1) "
        
        conn.Execute Sql1
    
        CalculoHorasProductivas = True
    End If

    TerminaBloquear
    Exit Function

eCalculoHorasProductivas:
    MuestraError Err.Number, "Calculo Horas Productivas", Err.Description
    TerminaBloquear
End Function


Private Function ProcesoCargaHoras(cTabla As String, cWhere As String, EsPrevision As Boolean) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim ImpBruto As Currency
Dim IRPF As Currency
Dim SegSoc As Currency
Dim Neto As Currency

    On Error GoTo eProcesoCargaHoras
    
    Screen.MousePointer = vbHourglass
    
    If Not EsPrevision Then
        Sql = "CARNOM" 'carga de nominas
        'Bloquear para que nadie mas pueda contabilizar
        DesBloqueoManual (Sql)
        If Not BloqueoManual(Sql, "1") Then
            MsgBox "No se puede realizar el proceso de Carga de Nóminas. Hay otro usuario realizándolo.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    ProcesoCargaHoras = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    If Not EsPrevision Then
        Sql = "Select rpartes_trabajador.nroparte, rpartes.fechapar, rpartes_trabajador.codtraba, sum(if(rpartes_trabajador.importe is null,0,rpartes_trabajador.importe)) FROM " & QuitarCaracterACadena(cTabla, "_1")
    Else
        Sql = "Select rpartes_trabajador.codtraba, rpartes.fechapar, sum(if(rpartes_trabajador.importe is null,0,rpartes_trabajador.importe)) FROM " & QuitarCaracterACadena(cTabla, "_1")
    End If
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    If Not EsPrevision Then
        Sql = Sql & " group by 1, 2, 3"
        Sql = Sql & " order by 1, 2, 3"
    Else
        Sql = Sql & " group by 1, 2"
        Sql = Sql & " order by 1, 2"
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    
    If EsPrevision Then
        Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
        conn.Execute Sql
        
        '                                       codtraba,fecha,  importe
        Sql = "insert into tmpinformes (codusu, codigo1, fecha1, importe1) values "
    Else
        Sql = "insert into horas (codtraba, fechahora, horasdia, horasproduc, compleme,"
        Sql = Sql & "intconta, pasaridoc, codalmac, nroparte) values "
    End If
        
        
    Sql3 = ""
    While Not Rs.EOF
        If Not EsPrevision Then
            Sql2 = "select count(*) from horas where fechahora = " & DBSet(Rs.Fields(1).Value, "F")
            Sql2 = Sql2 & " and codtraba = " & DBSet(Rs.Fields(2).Value, "N")
            Sql2 = Sql2 & " and codalmac = " & DBSet(vParamAplic.AlmacenNOMI, "N")
            
            If TotalRegistros(Sql2) = 0 Then
                Sql3 = Sql3 & "(" & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "F") & ",0,0,"
                Sql3 = Sql3 & DBSet(Rs.Fields(3).Value, "N") & ",0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ","
                Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "),"
            End If
        Else
            Sql2 = "select count(*) from tmpinformes where codigo1 = " & DBSet(Rs.Fields(0).Value, "N")
            Sql2 = Sql2 & " and fecha1 = " & DBSet(Rs.Fields(1).Value, "F")
            Sql2 = Sql2 & " and codusu = " & vUsu.Codigo
            
            If TotalRegistros(Sql2) = 0 Then
                Sql2 = "select count(*) from horas where fechahora = " & DBSet(Rs.Fields(1).Value, "F")
                Sql2 = Sql2 & " and codtraba = " & DBSet(Rs.Fields(0).Value, "N")
            
                If TotalRegistros(Sql2) = 0 Then
                    Sql3 = Sql3 & "(" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & ","
                    Sql3 = Sql3 & DBSet(Rs.Fields(1).Value, "F") & ","
                    Sql3 = Sql3 & DBSet(Rs.Fields(2).Value, "N") & "),"
                End If
            End If
        End If
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If Sql3 <> "" Then
        ' quitamos la ultima coma
        Sql3 = Mid(Sql3, 1, Len(Sql3) - 1)
        Sql = Sql & Sql3
        
        conn.Execute Sql
    End If
    
    If Not EsPrevision Then
        DesBloqueoManual ("CARNOM") 'carga de nominas
        
    Else
        
        Sql = "select codigo1, sum(importe1) from tmpinformes where codusu = " & vUsu.Codigo
        Sql = Sql & " group by 1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            
            Sql2 = "select salarios.impsalar, salarios.imphorae, straba.dtosirpf, straba.dtosegso, straba.porc_antig from salarios, straba where straba.codtraba = " & DBSet(Rs!Codigo1, "N")
            Sql2 = Sql2 & " and salarios.codcateg = straba.codcateg "
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            ImpBruto = Round2(DBLet(Rs.Fields(1).Value, "N"), 2)
            
    '        [Monica]23/03/2010: incrementamos el bruto el porcentaje de antigüedad si lo tiene, si no 0
            ImpBruto = ImpBruto + Round2(ImpBruto * DBLet(Rs2!porc_antig, "N") / 100, 2)
            
            IRPF = Round2(ImpBruto * DBLet(Rs2!dtosirpf, "N") / 100, 2)
            SegSoc = Round2(ImpBruto * DBLet(Rs2!dtosegso, "N") / 100, 2)
            
            Neto = Round2(ImpBruto - IRPF - SegSoc, 2)
            
            Sql3 = "update tmpinformes set importe2 = " & DBSet(ImpBruto, "N")
            Sql3 = Sql3 & ", importe3 = " & DBSet(IRPF, "N")
            Sql3 = Sql3 & ", importe4 = " & DBSet(SegSoc, "N")
            Sql3 = Sql3 & ", importe5 = " & DBSet(Neto, "N")
            Sql3 = Sql3 & " where codusu = " & vUsu.Codigo
            Sql3 = Sql3 & " and codigo1 = " & DBSet(Rs!Codigo1, "N")
            
            conn.Execute Sql3
            Set Rs2 = Nothing
                
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
    End If
    
    Screen.MousePointer = vbDefault
    
    ProcesoCargaHoras = True
    Exit Function
    
eProcesoCargaHoras:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Carga de Horas", Err.Description
End Function



Private Function ProcesoCargaHorasPicassent(cTabla As String, cWhere As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoCargaHorasPicassent
    
    Screen.MousePointer = vbHourglass
    
    Sql = "CARNOM" 'carga de nominas
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Carga de Nóminas. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    ProcesoCargaHorasPicassent = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = cTabla & " INNER JOIN rcuadrilla ON rpartes.codcuadrilla = rcuadrilla.codcuadrilla "
    Sql = "Select rpartes_trabajador.nroparte, rpartes.fechapar, rpartes_trabajador.codtraba, rpartes_trabajador.codvarie, rcuadrilla.codcapat, sum(rpartes_trabajador.importe), sum(rpartes_trabajador.kilosrec) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1, 2, 3, 4, 5"
    Sql = Sql & " order by 1, 2, 3, 4, 5"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    Sql = "insert into horas (codtraba, fechahora, horasdia, horasproduc, importe,"
    Sql = Sql & "intconta, pasaridoc, codalmac, nroparte, codvarie, codcapat, kilos) values "
        
    Sql3 = ""
    While Not Rs.EOF
        Sql2 = "select count(*) from horas where fechahora = " & DBSet(Rs.Fields(1).Value, "F")
        Sql2 = Sql2 & " and codtraba = " & DBSet(Rs.Fields(2).Value, "N")
        Sql2 = Sql2 & " and codalmac = " & DBSet(vParamAplic.AlmacenNOMI, "N")
        Sql2 = Sql2 & " and codvarie = " & DBSet(Rs.Fields(3).Value, "N")
        Sql2 = Sql2 & " and codcapat = " & DBSet(Rs.Fields(4).Value, "N")
        
        
        If TotalRegistros(Sql2) = 0 Then
            Sql3 = Sql3 & "(" & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "F") & ",0,0,"
            Sql3 = Sql3 & DBSet(Rs.Fields(5).Value, "N") & ",0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(3).Value, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(4).Value, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(6).Value, "N") & "),"
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If Sql3 <> "" Then
        ' quitamos la ultima coma
        Sql3 = Mid(Sql3, 1, Len(Sql3) - 1)
        Sql = Sql & Sql3
        
        conn.Execute Sql
    End If
    
    DesBloqueoManual ("CARNOM") 'carga de nominas
    
    Screen.MousePointer = vbDefault
    
    ProcesoCargaHorasPicassent = True
    Exit Function
    
eProcesoCargaHorasPicassent:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Carga de Horas", Err.Description
End Function


Private Function ProcesoCargaHorasCatadau(cTabla As String, cWhere As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Almacen As Integer
Dim Sql5 As String
Dim Nregs As Long

Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoCargaHorasCatadau
    
    Screen.MousePointer = vbHourglass
    
    Sql = "CARNOM" 'carga de nominas
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Carga de Nóminas. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    ProcesoCargaHorasCatadau = False

    Sql5 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql5


    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = cTabla & " INNER JOIN rcuadrilla ON rpartes.codcuadrilla = rcuadrilla.codcuadrilla "
    Sql = "Select rpartes_trabajador.nroparte, rpartes.fechapar, rpartes_trabajador.codtraba, rpartes_trabajador.codvarie, rcuadrilla.codcapat, sum(rpartes_trabajador.importe), sum(rpartes_trabajador.kilosrec), sum(rpartes_trabajador.horastra) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1, 2, 3, 4, 5"
    Sql = Sql & " order by 1, 2, 3, 4, 5"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    Sql = "insert into horas (codtraba, fechahora, horasdia, horasproduc, importe,"
    Sql = Sql & "intconta, pasaridoc, codalmac, nroparte, codvarie, codcapat, kilos) values "
        
    Sql3 = ""
    While Not Rs.EOF
        Sql2 = "select count(*) from horas where fechahora = " & DBSet(Rs.Fields(1).Value, "F")
        Sql2 = Sql2 & " and codtraba = " & DBSet(Rs.Fields(2).Value, "N")
        Sql2 = Sql2 & " and codalmac = " & DBSet(vParamAplic.AlmacenNOMI, "N")
        Sql2 = Sql2 & " and codvarie = " & DBSet(Rs.Fields(3).Value, "N")
'        Sql2 = Sql2 & " and codcapat = " & DBSet(Rs.Fields(4).Value, "N")
        
        Nregs = TotalRegistros(Sql2)
            
        Sql5 = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
        Sql5 = Sql5 & " and importe1 = " & DBSet(Rs.Fields(2).Value, "N")
        Sql5 = Sql5 & " and fecha1 = " & DBSet(Rs.Fields(1).Value, "F")
        Sql5 = Sql5 & " and importe2 = " & DBSet(Rs.Fields(3).Value, "N")
        Sql5 = Sql5 & " and importe3 = " & DBSet(vParamAplic.AlmacenNOMI, "N")
        
        Nregs = Nregs + TotalRegistros(Sql5)
            
        If Nregs = 0 Then
            Sql3 = Sql3 & "(" & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "F") & "," & DBSet(Rs.Fields(7).Value, "N") & "," & DBSet(Rs.Fields(7).Value, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(5).Value, "N") & ",0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(3).Value, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(4).Value, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(6).Value, "N") & "),"
        
            Sql5 = "insert into tmpinformes (codusu, importe1, fecha1, importe2, importe3) values ("
            Sql5 = Sql5 & vUsu.Codigo & "," & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "F") & ","
            Sql5 = Sql5 & DBSet(Rs.Fields(3).Value, "N") & "," & DBSet(vParamAplic.AlmacenNOMI, "N") & ")"
            
            conn.Execute Sql5
        
        Else
            '[Monica]18/06/2013: solo voy a dejar que el trabajador trabaje mañana y tarde
            '                    con lo cual en Catadau, almacen 2 significa tarde, y he de crearlo como tal.
            '                    suponemos que es un trabajador que trabaja por la tarde con el mismo capataz misma variedad
            Sql4 = "select max(codalmac) + 1 codalmac from horas where fechahora = " & DBSet(Rs.Fields(1).Value, "F")
            Sql4 = Sql4 & " and codtraba = " & DBSet(Rs.Fields(2).Value, "N")
            Sql4 = Sql4 & " and codvarie = " & DBSet(Rs.Fields(3).Value, "N")
            Sql4 = Sql4 & " union "
            Sql4 = Sql4 & " select max(importe3) + 1 codalmac from tmpinformes where codusu = " & vUsu.Codigo
            Sql4 = Sql4 & " and fecha1 = " & DBSet(Rs.Fields(1).Value, "F")
            Sql4 = Sql4 & " and importe1 = " & DBSet(Rs.Fields(2).Value, "N")
            Sql4 = Sql4 & " and importe2 = " & DBSet(Rs.Fields(3).Value, "N")
                        
            Sql4 = "select max(codalmac) from (" & Sql4 & ") aaaaa"
        
            Almacen = DevuelveValor(Sql4)
            
            Sql3 = Sql3 & "(" & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "F") & ",0,0,"
            Sql3 = Sql3 & DBSet(Rs.Fields(5).Value, "N") & ",0,0," & DBSet(Almacen, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(3).Value, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(4).Value, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(6).Value, "N") & "),"
            
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If Sql3 <> "" Then
        ' quitamos la ultima coma
        Sql3 = Mid(Sql3, 1, Len(Sql3) - 1)
        Sql = Sql & Sql3
        
        conn.Execute Sql
    End If
    
    DesBloqueoManual ("CARNOM") 'carga de nominas
    
    Screen.MousePointer = vbDefault
    
    ProcesoCargaHorasCatadau = True
    Exit Function
    
eProcesoCargaHorasCatadau:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Carga de Horas", Err.Description
End Function


Private Function ActualizarTabla(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim cadMen As String
Dim i As Long
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim numalbar As Long
Dim devuelve As String
Dim Existe As Boolean
Dim NumRegis As Long

Dim cTabla2 As String
Dim cWhere2 As String
Dim RS1 As ADODB.Recordset

    On Error GoTo eActualizarTabla
    
    ActualizarTabla = False

    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("ACTDES") 'RECtificativas FACturas
    If Not BloqueoManual("ACTDES", "1") Then
        MsgBox "No se puede actualizar. Hay otro usuario actualizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    b = True
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    Sql2 = " select codtraba, fechahora, horas, horas, "
    Sql2 = Sql2 & ValorNulo & "," ' complemento
    Sql2 = Sql2 & ValorNulo & "," ' horasini
    Sql2 = Sql2 & ValorNulo & "," ' horasfin
    Sql2 = Sql2 & ValorNulo & "," ' anticipo
    Sql2 = Sql2 & ValorNulo & "," ' horas extra
    Sql2 = Sql2 & ValorNulo & "," ' fecha recepcion
    Sql2 = Sql2 & "0,0," ' integracion contable / integracion aridoc
    Sql2 = Sql2 & vParamAplic.AlmacenNOMI & "," ' almacen por defecto
    Sql2 = Sql2 & ValorNulo & "," ' nro de parte
    Sql2 = Sql2 & "codvarie, " ' variedad
    Sql2 = Sql2 & "codforfait, " ' forfait
    Sql2 = Sql2 & "numcajon, " ' cajones
    Sql2 = Sql2 & "Kilos " ' kilos
    Sql2 = Sql2 & " from " & cTabla
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql2 = Sql2 & " WHERE " & cWhere
    End If
    
    conn.BeginTrans
    
    ' insertamos en horas
    Sql = "insert into horas (codtraba, fechahora, horasdia, horasproduc, compleme, horasini, horasfin, "
    Sql = Sql & "anticipo, horasext, fecharec, intconta, pasaridoc, codalmac, nroparte, codvarie, codforfait, "
    Sql = Sql & " numcajon, kilos) "
    Sql = Sql & Sql2
    
    conn.Execute Sql
    
    ' borramos de horasdestajo
    Sql = "delete from horasdestajo "
    If cWhere <> "" Then
        Sql = Sql & " WHERE " & cWhere
    End If
    
    conn.Execute Sql
    
eActualizarTabla:
    If Err.Number Then
        b = False
        MuestraError Err.Number, "Actualizando Horas Destajo", Err.Description & cadMen
    End If
    If b Then
        conn.CommitTrans
        ActualizarTabla = True
    Else
        conn.RollbackTrans
        ActualizarTabla = False
    End If
End Function


Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmCapataces(indice As Integer)
    indCodigo = indice
    Set frmCap = New frmManCapataz
    frmCap.DatosADevolverBusqueda = "0|1|"
    frmCap.CodigoActual = txtCodigo(indCodigo)
    frmCap.Show vbModal
    Set frmCap = Nothing
End Sub



Private Function CalculoDestajoETT(actualiza As Boolean) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Kilos As Long
Dim CodigoETT As Long
Dim Categoria As Long

Dim Precio As Currency
Dim Importe As Currency


    On Error GoTo eCalculoDestajoETT

    CalculoDestajoETT = False

    Sql = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    CodigoETT = DevuelveValor(Sql)

    Sql = "select codcateg from rcapataz left join straba on rcapataz.codtraba = straba.codtraba where rcapataz.codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    Categoria = DevuelveValor(Sql)


    Sql = "select sum(kilostra) from rclasifica where fechaent = " & DBSet(txtCodigo(11).Text, "F") & " and codvarie = " & DBSet(txtCodigo(9), "N")
    Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")

    Kilos = DevuelveValor(Sql)
    
    Sql = "select precio from rtarifaett where codvarie = " & DBSet(txtCodigo(9).Text, "N")
    Sql = Sql & " and codigoett = " & DBSet(CodigoETT, "N")
    
    Precio = DevuelveValor(Sql)
    
    Importe = Round2(Kilos * Precio, 2)
    
    txtCodigo(10).Text = Format(Kilos, "###,###,##0")
    txtCodigo(8).Text = Format(Precio, "###,##0.0000")
    txtCodigo(13).Text = Format(Importe, "###,###,##0.00")

    If Not actualiza Then
        CalculoDestajoETT = True
        Exit Function
    Else
        Sql = "select count(*) from horasett where fechahora = " & DBSet(txtCodigo(11).Text, "F")
        Sql = Sql & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
        Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
        Sql = Sql & " and codigoett = " & DBSet(CodigoETT, "N")
        
        If TotalRegistros(Sql) = 0 Then
            Sql1 = "insert into horasett (fechahora,codvarie,codigoett,codcapat,complemento,codcateg,importe,penaliza,"
            Sql1 = Sql1 & "complcapataz , kilosalicatados, kilostiron, fecharec, intconta, pasaridoc) values ("
            Sql1 = Sql1 & DBSet(txtCodigo(11).Text, "F") & ","
            Sql1 = Sql1 & DBSet(txtCodigo(9).Text, "N") & ","
            Sql1 = Sql1 & DBSet(CodigoETT, "N") & ","
            Sql1 = Sql1 & DBSet(txtCodigo(12).Text, "N") & ","
            Sql1 = Sql1 & "0,"
            Sql1 = Sql1 & DBSet(Categoria, "N") & ","
            Sql1 = Sql1 & DBSet(Importe, "N") & ","
            Sql1 = Sql1 & "0,0,"
            Sql1 = Sql1 & DBSet(Kilos, "N") & ","
            Sql1 = Sql1 & "0,null,0,0) "
            
            conn.Execute Sql1
        Else
            Sql1 = "update horasett set importe = " & DBSet(Importe, "N")
            Sql1 = Sql1 & ", kilosalicatados = " & DBSet(Kilos, "N")
            Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            Sql1 = Sql1 & " and codigoett = " & DBSet(CodigoETT, "N")
            
            conn.Execute Sql1
        End If
        
        CalculoDestajoETT = True
        Exit Function
    End If
    
eCalculoDestajoETT:
    MuestraError Err.Number, "Calculo Destajo ETT", Err.Description
    TerminaBloquear
End Function



Private Function CalculoPenalizacionETT(actualiza As Boolean) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Kilos As Long
Dim CodigoETT As Long
Dim KilosTiron As Long

Dim Penalizacion As Currency

Dim Precio As Currency

Dim ImporteTotal As Currency
Dim ImporteAlicatado As Currency
Dim Porcentaje As Currency



    On Error GoTo eCalculoPenalizacionETT

    CalculoPenalizacionETT = False

    Sql = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    CodigoETT = DevuelveValor(Sql)

    Sql = "select sum(kilostra) from rclasifica where fechaent = " & DBSet(txtCodigo(11).Text, "F") & " and codvarie = " & DBSet(txtCodigo(9), "N")
    Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")

    Porcentaje = 0
    If txtCodigo(21).Text <> "" Then Porcentaje = CCur(ImporteSinFormato(txtCodigo(21).Text))


    Kilos = DevuelveValor(Sql)
    KilosTiron = Round2(Kilos * Porcentaje * 0.01, 0)
    
    Sql = "select precio from rtarifaett where codvarie = " & DBSet(txtCodigo(9).Text, "N")
    Sql = Sql & " and codigoett = " & DBSet(CodigoETT, "N")
    
    Precio = DevuelveValor(Sql)
    
    ImporteAlicatado = Round2((Kilos - KilosTiron) * Precio, 2)
    ImporteTotal = Round2(Kilos * Precio, 2)
    Penalizacion = ImporteTotal - ImporteAlicatado
    
    txtCodigo(22).Text = Format(Kilos, "###,###,##0")
    txtCodigo(20).Text = Format(Penalizacion, "###,###,##0.00")

    If Not actualiza Then
        CalculoPenalizacionETT = True
        Exit Function
    Else
        
        Sql1 = "update horasett set  penaliza = " & DBSet(Penalizacion, "N")
        Sql1 = Sql1 & ", kilosalicatados = " & DBSet(Kilos - KilosTiron, "N")
        Sql1 = Sql1 & ", kilostiron = " & DBSet(KilosTiron, "N")
        Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
        Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
        Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
        Sql1 = Sql1 & " and codigoett = " & DBSet(CodigoETT, "N")
        
        conn.Execute Sql1
        
        CalculoPenalizacionETT = True
        Exit Function
    End If
    
eCalculoPenalizacionETT:
    MuestraError Err.Number, "Calculo Penalizacion ETT", Err.Description
End Function
                               

Private Function CalculoBonificacionETT(actualiza As Boolean) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim CodigoETT As Long

    On Error GoTo eCalculoBonificacionETT

    CalculoBonificacionETT = False

    Sql = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    CodigoETT = DevuelveValor(Sql)

    Sql1 = "update horasett set  complemento = " & DBSet(txtCodigo(23).Text, "N")
    Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
    Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
    Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
    Sql1 = Sql1 & " and codigoett = " & DBSet(CodigoETT, "N")
    
    conn.Execute Sql1
        
    CalculoBonificacionETT = True
    Exit Function
    
eCalculoBonificacionETT:
    MuestraError Err.Number, "Calculo Bonificacion ETT", Err.Description
End Function
                               


Private Function ProcesoBorradoMasivo(cTabla As String, cWhere As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoBorradoMasivo
    
    Screen.MousePointer = vbHourglass
    
    Sql = "BORMAS" 'BORrado MASivo
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Borrado Masivo. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    ProcesoBorradoMasivo = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "delete FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    conn.Execute Sql
        
    DesBloqueoManual ("BORMAS") 'BORrado MASivo"
    
    Screen.MousePointer = vbDefault
    
    ProcesoBorradoMasivo = True
    Exit Function
    
eProcesoBorradoMasivo:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Borrado Masivo", Err.Description
End Function



Private Function CalculoAltaRapida() As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Kilos As Long
Dim CodigoETT As Long
Dim Categoria As Long

Dim Precio As Currency
Dim Importe As Currency
Dim i As Integer

Dim Fdesde As Date
Dim Fhasta As Date
Dim Fecha As Date

Dim trabajador As Long
Dim Dias As Long

    On Error GoTo eCalculoAltaRapida

    CalculoAltaRapida = False

    Sql = "select codtraba from rcapataz where rcapataz.codcapat = " & DBSet(txtCodigo(34).Text, "N")
    
    trabajador = DevuelveValor(Sql)

    Sql = "select codcateg from straba where codtraba = " & DBSet(trabajador, "N")

    Categoria = DevuelveValor(Sql)

    Fdesde = CDate(txtCodigo(35).Text)
    Fhasta = CDate(txtCodigo(26).Text)

    Dias = Fhasta - Fdesde

    Importe = 0
    If txtCodigo(40).Text <> "" Then
        Importe = ImporteSinFormato(txtCodigo(40).Text)
    End If

    For i = 0 To Dias
        Fecha = DateAdd("y", i, Fdesde)

        Sql = "select count(*) from horas where fechahora = " & DBSet(Fecha, "F")
        Sql = Sql & " and codvarie = " & DBSet(txtCodigo(36).Text, "N")
        Sql = Sql & " and codcapat = " & DBSet(txtCodigo(34).Text, "N")
        Sql = Sql & " and codtraba = " & DBSet(trabajador, "N")
        
        If TotalRegistros(Sql) = 0 Then
            Sql1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,importe,fecharec,intconta,pasaridoc,codalmac) values ("
            Sql1 = Sql1 & DBSet(Fecha, "F") & ","
            Sql1 = Sql1 & DBSet(txtCodigo(36).Text, "N") & ","
            Sql1 = Sql1 & DBSet(trabajador, "N") & ","
            Sql1 = Sql1 & DBSet(txtCodigo(34).Text, "N") & ","
            Sql1 = Sql1 & DBSet(Importe, "N") & ","
            Sql1 = Sql1 & "null,0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ") "
            
            conn.Execute Sql1
        End If
        
    Next i
    
    CalculoAltaRapida = True
    Exit Function
    
eCalculoAltaRapida:
    MuestraError Err.Number, "Calculo Alta Rápida", Err.Description
    TerminaBloquear
End Function



Private Function CalculoEventuales() As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Kilos As Long
Dim CodigoETT As Long
Dim Categoria As Long

Dim Precio As Currency
Dim Importe As Currency

Dim i As Integer
Dim J As Integer

Dim Fdesde As Date
Dim Fhasta As Date
Dim Fecha As Date

Dim TrabaDesde As Long
Dim Trabahasta As Long
Dim Dias As Long

    On Error GoTo eCalculoEventuales

    CalculoEventuales = False

    TrabaDesde = CLng(txtCodigo(41).Text)
    Trabahasta = CLng(txtCodigo(42).Text)

    Fdesde = CDate(txtCodigo(37).Text)
    Fhasta = CDate(txtCodigo(33).Text)

    Dias = Fhasta - Fdesde
        
    Importe = 0
    If txtCodigo(39).Text <> "" Then
        Importe = ImporteSinFormato(txtCodigo(39).Text)
    End If

    For J = TrabaDesde To Trabahasta
        '[Monica]29/10/2014: añadimos la condicion de que el trabajador que vamos a introducir no tenga fecha de baja
        If TotalRegistros("select count(*) from straba where codtraba = " & J & " and (fechabaja is null or fechabaja = '')") <> 0 Then
    
            For i = 0 To Dias
                Fecha = DateAdd("y", i, Fdesde)
        
                Sql = "select count(*) from horas where fechahora = " & DBSet(Fecha, "F")
                Sql = Sql & " and codvarie = " & DBSet(txtCodigo(28).Text, "N")
                Sql = Sql & " and codcapat = " & DBSet(0, "N")
                Sql = Sql & " and codtraba = " & DBSet(J, "N")
                
                If TotalRegistros(Sql) = 0 Then
                    Sql1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,importe,fecharec,intconta,pasaridoc,codalmac) values ("
                    Sql1 = Sql1 & DBSet(Fecha, "F") & ","
                    Sql1 = Sql1 & DBSet(txtCodigo(28).Text, "N") & ","
                    Sql1 = Sql1 & DBSet(J, "N") & ","
                    Sql1 = Sql1 & "0,"
                    Sql1 = Sql1 & DBSet(Importe, "N") & ","
                    Sql1 = Sql1 & "null,0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ") "
                    
                    conn.Execute Sql1
                End If
                
            Next i
        End If
    Next J
    
    CalculoEventuales = True
    Exit Function
    
eCalculoEventuales:
    MuestraError Err.Number, "Calculo Eventuales", Err.Description
    TerminaBloquear
End Function




Private Function CalculoTrabajCapataz() As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Importe As Currency

    On Error GoTo eCalculoTrabajCapataz

    CalculoTrabajCapataz = False
        
    conn.BeginTrans
        
    Importe = 0
    If txtCodigo(51).Text <> "" Then
        Importe = ImporteSinFormato(txtCodigo(51).Text)
    End If

    Sql = "select * from rcuadrilla INNER JOIN rcuadrilla_trabajador ON rcuadrilla.codcuadrilla = rcuadrilla_trabajador.codcuadrilla "
    Sql = Sql & " where rcuadrilla.codcapat = " & DBSet(txtCodigo(45).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(46).Text, "F")
        Sql = Sql & " and codvarie = " & DBSet(txtCodigo(47).Text, "N")
        Sql = Sql & " and codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql = Sql & " and codcapat = " & DBSet(txtCodigo(45).Text, "N")
        
        If TotalRegistros(Sql) = 0 Then
            Sql1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,importe,compleme, fecharec,intconta,pasaridoc,codalmac) values ("
            Sql1 = Sql1 & DBSet(txtCodigo(46).Text, "F") & ","
            Sql1 = Sql1 & DBSet(txtCodigo(47).Text, "N") & ","
            Sql1 = Sql1 & DBSet(Rs!CodTraba, "N") & ","
            Sql1 = Sql1 & DBSet(txtCodigo(45).Text, "N") & ",null, "
            Sql1 = Sql1 & DBSet(Importe, "N") & ","
            Sql1 = Sql1 & "null,0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ") "
            
            conn.Execute Sql1
        Else
            Sql1 = "update horas set compleme = if(compleme is null,0,compleme) + " & DBSet(Importe, "N")
            Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(46).Text, "F")
            Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(47).Text, "N")
            Sql1 = Sql1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
            Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(45).Text, "N")
        
            conn.Execute Sql1
        End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    conn.CommitTrans
                
    CalculoTrabajCapataz = True
    Exit Function
    
eCalculoTrabajCapataz:
    MuestraError Err.Number, "Calculo Trabajadores para un Capataz", Err.Description
    conn.RollbackTrans
    TerminaBloquear
End Function


Private Function CalculoTrabajCapatazNew() As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Importe As Currency

    On Error GoTo eCalculoTrabajCapatazNew

    CalculoTrabajCapatazNew = False
        
    conn.BeginTrans
        
    Importe = 0
    If txtCodigo(51).Text <> "" Then
        Importe = ImporteSinFormato(txtCodigo(51).Text)
    End If

    Sql = "select * from horas "
    Sql = Sql & " where horas.codcapat = " & DBSet(txtCodigo(45).Text, "N")
    Sql = Sql & " and horas.fechahora = " & DBSet(txtCodigo(46).Text, "F")
    Sql = Sql & " and horas.codvarie = " & DBSet(txtCodigo(47).Text, "N")
    
    If TotalRegistrosConsulta(Sql) <> 0 Then
    
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            Sql = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(46).Text, "F")
            Sql = Sql & " and codvarie = " & DBSet(txtCodigo(47).Text, "N")
            Sql = Sql & " and codtraba = " & DBSet(Rs!CodTraba, "N")
            Sql = Sql & " and codcapat = " & DBSet(txtCodigo(45).Text, "N")
            
            If TotalRegistros(Sql) = 0 Then
                Sql1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,importe,compleme, fecharec,intconta,pasaridoc,codalmac) values ("
                Sql1 = Sql1 & DBSet(txtCodigo(46).Text, "F") & ","
                Sql1 = Sql1 & DBSet(txtCodigo(47).Text, "N") & ","
                Sql1 = Sql1 & DBSet(Rs!CodTraba, "N") & ","
                Sql1 = Sql1 & DBSet(txtCodigo(45).Text, "N") & ",null, "
                Sql1 = Sql1 & DBSet(Importe, "N") & ","
                Sql1 = Sql1 & "null,0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ") "
                
                conn.Execute Sql1
            Else
                Sql1 = "update horas set compleme = if(compleme is null,0,compleme) + " & DBSet(Importe, "N")
                Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(46).Text, "F")
                Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(47).Text, "N")
                Sql1 = Sql1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
                Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(45).Text, "N")
            
                conn.Execute Sql1
            
                Sql1 = "update horas set compleme = if(compleme=0,null,compleme) "
                Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(46).Text, "F")
                Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(47).Text, "N")
                Sql1 = Sql1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
                Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(45).Text, "N")
            
                conn.Execute Sql1
            
            End If
            
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    
    Else
    
        MsgBox "No hay entradas de horas para esa variedad, fecha y capataz. Revise.", vbExclamation
        conn.CommitTrans
        Exit Function
    End If
    
    
    conn.CommitTrans
                
    CalculoTrabajCapatazNew = True
    Exit Function
    
eCalculoTrabajCapatazNew:
    MuestraError Err.Number, "Calculo Trabajadores para un Capataz", Err.Description
    conn.RollbackTrans
    TerminaBloquear
End Function





Private Function CalculoDestajo(actualiza As Boolean) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Kilos As Long
Dim CodigoETT As Long
Dim Categoria As Long

Dim Precio As Currency
Dim Importe As Currency

Dim KilosTrab As Long
Dim ImporteTrab As Currency
Dim Cuadrilla As Long
Dim Nregs As Long

    On Error GoTo eCalculoDestajo

    CalculoDestajo = False

    Sql = "select codcuadrilla from rcuadrilla where codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    Cuadrilla = DevuelveValor(Sql)

    Sql = "select count(*) from rcuadrilla_trabajador , rcuadrilla where rcuadrilla.codcapat = " & DBSet(txtCodigo(12).Text, "N")
    Sql = Sql & " and rcuadrilla.codcuadrilla = rcuadrilla_trabajador.codcuadrilla"
    
    Nregs = DevuelveValor(Sql)
    
    If Nregs <> 0 Then
        Sql = "select sum(kilostra) from rclasifica where fechaent = " & DBSet(txtCodigo(11).Text, "F") & " and codvarie = " & DBSet(txtCodigo(9), "N")
        Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
        Kilos = DevuelveValor(Sql)
        
        Sql = "select eurdesta from variedades where codvarie = " & DBSet(txtCodigo(9).Text, "N")
        
        Precio = DevuelveValor(Sql)
        
        Importe = Round2(Kilos * Precio, 2)
        
        txtCodigo(10).Text = Format(Kilos, "###,###,##0")
        txtCodigo(8).Text = Format(Precio, "###,##0.0000")
        txtCodigo(13).Text = Format(Importe, "###,###,##0.00")
        If Not actualiza Then
            CalculoDestajo = True
            Exit Function
        Else
            KilosTrab = Round(Kilos / Nregs, 0)
            ImporteTrab = Round2(Importe / Nregs, 2)
            
            Sql = "select codtraba from rcuadrilla_trabajador , rcuadrilla where rcuadrilla.codcapat = " & DBSet(txtCodigo(12).Text, "N")
            Sql = Sql & " and rcuadrilla.codcuadrilla = rcuadrilla_trabajador.codcuadrilla"
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            While Not Rs.EOF
                Sql = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(11).Text, "F")
                Sql = Sql & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
                Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
                Sql = Sql & " and codtraba = " & DBSet(Rs!CodTraba, "N")
                
                If TotalRegistros(Sql) = 0 Then
                    Sql1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,compleme,importe,penaliza,"
                    Sql1 = Sql1 & "kilos, fecharec, intconta, pasaridoc,codalmac) values ("
                    Sql1 = Sql1 & DBSet(txtCodigo(11).Text, "F") & ","
                    Sql1 = Sql1 & DBSet(txtCodigo(9).Text, "N") & ","
                    Sql1 = Sql1 & DBSet(Rs!CodTraba, "N") & ","
                    Sql1 = Sql1 & DBSet(txtCodigo(12).Text, "N") & ","
                    Sql1 = Sql1 & "0,"
                    Sql1 = Sql1 & DBSet(ImporteTrab, "N") & ","
                    Sql1 = Sql1 & "0,"
                    Sql1 = Sql1 & DBSet(KilosTrab, "N") & ","
                    Sql1 = Sql1 & "null,0,0, "
                    Sql1 = Sql1 & vParamAplic.AlmacenNOMI & ") "
                    
                    conn.Execute Sql1
                Else
                    Sql1 = "update horas set importe = " & DBSet(ImporteTrab, "N")
                    Sql1 = Sql1 & ", kilos = " & DBSet(KilosTrab, "N")
                    Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
                    Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
                    Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
                    Sql1 = Sql1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
                    
                    conn.Execute Sql1
                End If
                
                Rs.MoveNext
            Wend
        End If
    End If
    CalculoDestajo = True
    Exit Function
    
eCalculoDestajo:
    MuestraError Err.Number, "Calculo Destajo", Err.Description
    TerminaBloquear
End Function




Private Function CalculoPenalizacion(actualiza As Boolean) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Kilos As Long
Dim CodigoETT As Long
Dim KilosTiron As Long

Dim Penalizacion As Currency
Dim PenalizacionTrab As Currency
Dim PenalizacionDif As Currency
Dim NumTrab As Long

Dim Precio As Currency

Dim ImporteTotal As Currency
Dim ImporteAlicatado As Currency
Dim Porcentaje As Currency

Dim KilosTrab As Long
Dim KilosTironTrab As Long

Dim KilosDif As Long
Dim KilosTironDif As Long

Dim TrabCapataz As Long


    On Error GoTo eCalculoPenalizacion

    CalculoPenalizacion = False

    Sql = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    CodigoETT = DevuelveValor(Sql)

    Sql = "select sum(kilostra) from rclasifica where fechaent = " & DBSet(txtCodigo(11).Text, "F") & " and codvarie = " & DBSet(txtCodigo(9), "N")
    Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")

    Porcentaje = 0
    If txtCodigo(21).Text <> "" Then Porcentaje = CCur(ImporteSinFormato(txtCodigo(21).Text))


    Kilos = DevuelveValor(Sql)
    KilosTiron = Round2(Kilos * Porcentaje * 0.01, 0)
    
    '[Monica]06/10/2011: antes era eurhaneg
    Sql = "select eurdesta from variedades where codvarie = " & DBSet(txtCodigo(9).Text, "N")
    
    Precio = DevuelveValor(Sql)
    
    ImporteAlicatado = Round2((Kilos - KilosTiron) * Precio, 2)
    ImporteTotal = Round2(Kilos * Precio, 2)
    Penalizacion = ImporteTotal - ImporteAlicatado
    
    Sql = "select codtraba from horas where fechahora = " & DBSet(txtCodigo(11).Text, "F")
    Sql = Sql & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
    Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    NumTrab = TotalRegistrosConsulta(Sql)
    PenalizacionTrab = 0
    If NumTrab <> 0 Then PenalizacionTrab = Round2(Penalizacion / NumTrab, 2)
    PenalizacionDif = Round2(Penalizacion - (PenalizacionTrab * NumTrab), 2)
    KilosTrab = 0
    KilosTironTrab = 0
    If NumTrab <> 0 Then
        KilosTrab = Round2(Kilos / NumTrab, 0)
        KilosTironTrab = Round2(KilosTiron / NumTrab, 0)
    End If
    KilosDif = Kilos - Round2(KilosTrab * NumTrab, 0)
    KilosTironDif = KilosTiron - Round2(KilosTironTrab * NumTrab, 0)
    
    txtCodigo(22).Text = Format(Kilos, "###,###,##0")
    txtCodigo(20).Text = Format(Penalizacion, "###,###,##0.00")

    If Not actualiza Then
        CalculoPenalizacion = True
        Exit Function
    Else
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
        
            Sql1 = "update horas set  penaliza = " & DBSet(PenalizacionTrab, "N")
            Sql1 = Sql1 & ", kilos = " & DBSet(KilosTrab, "N")
            Sql1 = Sql1 & ", kilostiron = " & DBSet(KilosTironTrab, "N")
            Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            Sql1 = Sql1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
            
            conn.Execute Sql1
        
            Rs.MoveNext
        
        Wend
        
        If PenalizacionDif <> 0 Or KilosDif <> 0 Or KilosTironDif <> 0 Then
            TrabCapataz = DevuelveValor("select codtraba from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N"))
            
            Sql1 = "update horas set penaliza = penaliza + " & DBSet(PenalizacionDif, "N")
            Sql1 = Sql1 & ", kilos = kilos + " & DBSet(KilosDif, "N")
            Sql1 = Sql1 & ", kilostiron = kilostiron + " & DBSet(KilosTironDif, "N")
            Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            Sql1 = Sql1 & " and codtraba = " & DBSet(TrabCapataz, "N")
            
            conn.Execute Sql1
        End If
        
        Set Rs = Nothing
        
        CalculoPenalizacion = True
        Exit Function
    End If
    
eCalculoPenalizacion:
    MuestraError Err.Number, "Calculo Penalizacion", Err.Description
End Function


Private Function CalculoBonificacion(actualiza As Boolean) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Bonif As Currency
Dim NumTrab As Long

Dim BonifTrab As Currency
Dim BonifDif As Currency
Dim TrabCapataz As Long

    On Error GoTo eCalculoBonificacion

    CalculoBonificacion = False

    Sql = "select codtraba from horas where fechahora = " & DBSet(txtCodigo(11).Text, "F")
    Sql = Sql & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
    Sql = Sql & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    NumTrab = TotalRegistrosConsulta(Sql)
    
    Bonif = CCur(ImporteSinFormato(txtCodigo(23).Text))
    BonifTrab = 0
    If NumTrab <> 0 Then BonifTrab = Round2(Bonif / NumTrab, 2)
    
    BonifDif = Bonif - Round2(BonifTrab * NumTrab, 2)
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql1 = "update horas set  compleme = " & DBSet(BonifTrab, "N")
        Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
        Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
        Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
        Sql1 = Sql1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
        
        conn.Execute Sql1
        
        Rs.MoveNext
    Wend
    
    If BonifDif <> 0 Then
        TrabCapataz = DevuelveValor("select codtraba from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N"))
    
        Sql1 = "update horas set  complemen = " & DBSet(BonifDif, "N")
        Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
        Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
        Sql1 = Sql1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
        Sql1 = Sql1 & " and codtraba = " & DBSet(TrabCapataz, "N")
        
        conn.Execute Sql1
    
    End If
        
    Set Rs = Nothing
    
    CalculoBonificacion = True
    Exit Function
    
eCalculoBonificacion:
    MuestraError Err.Number, "Cálculo Bonificacion", Err.Description
End Function


Private Function ProcesoEntradasCapataz(cTabla As String, cWhere As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

Dim VarieAnt As Long
Dim FechaAnt As Date
Dim CapatAnt As Long

Dim TotCajon As Long
Dim TotKilos As Long

Dim Importe As Currency
Dim ImporteTot As Currency

Dim CodigoETT As Long
Dim Nregs As Integer

    On Error GoTo eProcesoEntradasCapataz
    
    Screen.MousePointer = vbHourglass
    
    ProcesoEntradasCapataz = False

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "select rentradas.codcapat, rentradas.fechaent, rentradas.codvarie, sum(rentradas.numcajo1) as cajon, sum(rentradas.kilostra) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rentradas")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & Replace(Replace(cWhere, "horas", "rentradas"), "fechahora", "fechaent")
    End If
    Sql = Sql & " group by 1,2,3 "
    Sql = Sql & " union "
    
    Sql = Sql & "select rclasifica.codcapat, rclasifica.fechaent, rclasifica.codvarie, sum(rclasifica.numcajon) as cajon, sum(rclasifica.kilostra) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rclasifica")
    If cWhere <> "" Then
        Sql = Sql & " WHERE " & Replace(Replace(cWhere, "horas", "rclasifica"), "fechahora", "fechaent")
    End If
    Sql = Sql & " group by 1, 2, 3 "
'    Sql = Sql & " union "
'
'    Sql = Sql & "select rhisfruta_entradas.codcapat, rhisfruta_entradas.fechaent, rhisfruta.codvarie, sum(rhisfruta_entradas.numcajon) as cajon, sum(rhisfruta_entradas.kilostra) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rhisfruta_entradas")
'    Sql = Sql & " INNER JOIN rhisfruta ON rhisfruta_entradas.numalbar = rhisfruta.numalbar "
'    If cWhere <> "" Then
'        Sql = Sql & " WHERE " & Replace(Replace(cWhere, "horas", "rhisfruta_entradas"), "fechahora", "fechaent")
'    End If
'    Sql = Sql & " group by 1, 2, 3 "
    
    
    Sql = Sql & " order by 1, 2, 3 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        VarieAnt = DBLet(Rs!codvarie, "N")
        CapatAnt = DBLet(Rs!codcapat, "N")
        FechaAnt = DBLet(Rs!FechaEnt, "F")
        
        TotCajon = 0
        TotKilos = 0
    End If
    Sql2 = ""
    Nregs = 0
                                        '   capataz,fecha,  variedad, numcajon, kilos
    Sql = "insert into tmpinformes (codusu, campo1, fecha1, importe1, importe2, importe3) values  "
    While Not Rs.EOF
        If DBLet(Rs!codcapat, "N") <> CapatAnt Or DBLet(Rs!FechaEnt, "F") <> FechaAnt Or DBLet(Rs!codvarie, "N") <> VarieAnt Then
            Sql2 = Sql2 & "( " & vUsu.Codigo & "," & DBSet(CapatAnt, "N") & "," & DBSet(FechaAnt, "F") & "," & DBSet(VarieAnt, "N") & ","
            Sql2 = Sql2 & DBSet(TotCajon, "N") & "," & DBSet(TotKilos, "N") & "),"
        
            VarieAnt = DBLet(Rs!codvarie, "N")
            CapatAnt = DBLet(Rs!codcapat, "N")
            FechaAnt = DBLet(Rs!FechaEnt, "F")
        
            TotCajon = 0
            TotKilos = 0
        
        End If
        
        TotCajon = TotCajon + DBLet(Rs!cajon, "N")
        TotKilos = TotKilos + DBLet(Rs!Kilos, "N")
        Nregs = 1
        Rs.MoveNext
    Wend
    
    ' ultimo registro
    If Nregs <> 0 Then
        Sql2 = Sql2 & "( " & vUsu.Codigo & "," & DBSet(CapatAnt, "N") & "," & DBSet(FechaAnt, "F") & "," & DBSet(VarieAnt, "N") & ","
        Sql2 = Sql2 & DBSet(TotCajon, "N") & "," & DBSet(TotKilos, "N") & "),"
    End If
    
    Set Rs = Nothing
    
    If Sql2 <> "" Then ' quitamos la ultima coma
        Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
    
        conn.Execute Sql & Sql2
    End If
    
  
                'capataz, fecha,  variedad
    Sql = "select campo1, fecha1, importe1 from tmpinformes where codusu = " & vUsu.Codigo & " order by 1,2,3"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql = "select if(isnull(sum(importe)),0,sum(importe)) + if(isnull(sum(compleme)),0,sum(compleme)) - if(isnull(sum(penaliza)),0,sum(penaliza)) as importe "
        Sql = Sql & " from horas where codcapat = " & DBSet(Rs!campo1, "N")
        Sql = Sql & " and fechahora = " & DBSet(Rs!fecha1, "F")
        Sql = Sql & " and codvarie = " & DBSet(Rs!importe1, "N")
    
        Importe = DevuelveValor(Sql)
        ImporteTot = Importe
        
        CodigoETT = DevuelveValor("select codigoett from rcapataz where codcapat = " & DBSet(Rs!campo1, "N"))
         
        ' si es ett tendrá registros en horasett
        Sql = "select if(isnull(sum(importe)),0,sum(importe)) + if(isnull(sum(complemento)),0,sum(complemento)) - if(isnull(sum(penaliza)),0,sum(penaliza)) "
        Sql = Sql & " from horasett where codcapat = " & DBSet(Rs!campo1, "N")
        Sql = Sql & " and fechahora = " & DBSet(Rs!fecha1, "F")
        Sql = Sql & " and codvarie = " & DBSet(Rs!importe1, "N")
        Sql = Sql & " and codigoett = " & DBSet(CodigoETT, "N")
        
        Importe = DevuelveValor(Sql)
        ImporteTot = ImporteTot + Importe
    
        Sql2 = "update tmpinformes set importe4 = " & DBSet(ImporteTot, "N")
        Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
        Sql2 = Sql2 & " and campo1 = " & DBSet(Rs!campo1, "N")
        Sql2 = Sql2 & " and fecha1 = " & DBSet(Rs!fecha1, "F")
        Sql2 = Sql2 & " and importe1 = " & DBSet(Rs!importe1, "N")
    
        conn.Execute Sql2
    
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    Screen.MousePointer = vbDefault
    
    ProcesoEntradasCapataz = True
    Exit Function
    
eProcesoEntradasCapataz:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Entradas Capataz", Err.Description
End Function


Private Function ProcesoEntradasCapatazRdto(cTabla As String, cWhere As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

Dim VarieAnt As Long
Dim FechaAnt As Date
Dim CapatAnt As Long

Dim TotCajon As Long
Dim TotKilos As Long

Dim Importe As Currency
Dim ImporteTot As Currency

Dim CodigoETT As Long
Dim Nregs As Integer

    On Error GoTo eProcesoEntradasCapatazRdto
    
    Screen.MousePointer = vbHourglass
    
    ProcesoEntradasCapatazRdto = False

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    '[Monica]05/02/2014: solo lo cambio para Picassent, para el resto lo dejo como estaba
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        Sql = "select rentradas.codcapat, rentradas.fechaent, rentradas.codvarie, sum("
        If vParamAplic.EsCaja1 Then Sql = Sql & "+coalesce(rentradas.numcajo1,0)"
        If vParamAplic.EsCaja2 Then Sql = Sql & "+coalesce(rentradas.numcajo2,0)"
        If vParamAplic.EsCaja3 Then Sql = Sql & "+coalesce(rentradas.numcajo3,0)"
        If vParamAplic.EsCaja4 Then Sql = Sql & "+coalesce(rentradas.numcajo4,0)"
        If vParamAplic.EsCaja5 Then Sql = Sql & "+coalesce(rentradas.numcajo5,0)"
        
        Sql = Sql & ") as cajon, sum(rentradas.kilosnet) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rentradas")
    Else
        Sql = "select rentradas.codcapat, rentradas.fechaent, rentradas.codvarie, sum(rentradas.numcajo1) as cajon, sum(rentradas.kilosnet) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rentradas")
    End If
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & Replace(Replace(cWhere, "horas", "rentradas"), "fechahora", "fechaent")
    End If
    Sql = Sql & " group by 1,2,3 "
    Sql = Sql & " union "
    
    Sql = Sql & "select rclasifica.codcapat, rclasifica.fechaent, rclasifica.codvarie, sum(rclasifica.numcajon) as cajon, sum(rclasifica.kilosnet) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rclasifica")
    If cWhere <> "" Then
        Sql = Sql & " WHERE " & Replace(Replace(cWhere, "horas", "rclasifica"), "fechahora", "fechaent")
    End If
    Sql = Sql & " group by 1, 2, 3 "
    Sql = Sql & " union "

    Sql = Sql & "select rhisfruta_entradas.codcapat, rhisfruta_entradas.fechaent, rhisfruta.codvarie, sum(rhisfruta_entradas.numcajon) as cajon, sum(rhisfruta_entradas.kilosnet) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rhisfruta_entradas")
    Sql = Sql & " INNER JOIN rhisfruta ON rhisfruta_entradas.numalbar = rhisfruta.numalbar "
    If cWhere <> "" Then
        Sql = Sql & " WHERE " & Replace(Replace(cWhere, "horas", "rhisfruta_entradas"), "fechahora", "fechaent")
    End If
    Sql = Sql & " group by 1, 2, 3 "
    
    Sql = Sql & " order by 1, 2, 3 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        VarieAnt = DBLet(Rs!codvarie, "N")
        CapatAnt = DBLet(Rs!codcapat, "N")
        FechaAnt = DBLet(Rs!FechaEnt, "F")
        
        TotCajon = 0
        TotKilos = 0
    End If
    Sql2 = ""
    Nregs = 0
                                        '   capataz,fecha,  variedad, numcajon, kilos
    Sql = "insert into tmpinformes (codusu, campo1, fecha1, importe1, importe2, importe3) values  "
    While Not Rs.EOF
        If DBLet(Rs!codcapat, "N") <> CapatAnt Or DBLet(Rs!FechaEnt, "F") <> FechaAnt Or DBLet(Rs!codvarie, "N") <> VarieAnt Then
            Sql2 = Sql2 & "( " & vUsu.Codigo & "," & DBSet(CapatAnt, "N") & "," & DBSet(FechaAnt, "F") & "," & DBSet(VarieAnt, "N") & ","
            Sql2 = Sql2 & DBSet(TotCajon, "N") & "," & DBSet(TotKilos, "N") & "),"
        
            VarieAnt = DBLet(Rs!codvarie, "N")
            CapatAnt = DBLet(Rs!codcapat, "N")
            FechaAnt = DBLet(Rs!FechaEnt, "F")
        
            TotCajon = 0
            TotKilos = 0
        
        End If
        
        TotCajon = TotCajon + DBLet(Rs!cajon, "N")
        TotKilos = TotKilos + DBLet(Rs!Kilos, "N")
        Nregs = 1
        Rs.MoveNext
    Wend
    
    ' ultimo registro
    If Nregs <> 0 Then
        Sql2 = Sql2 & "( " & vUsu.Codigo & "," & DBSet(CapatAnt, "N") & "," & DBSet(FechaAnt, "F") & "," & DBSet(VarieAnt, "N") & ","
        Sql2 = Sql2 & DBSet(TotCajon, "N") & "," & DBSet(TotKilos, "N") & "),"
    End If
    
    Set Rs = Nothing
    
    If Sql2 <> "" Then ' quitamos la ultima coma
        Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
    
        conn.Execute Sql & Sql2
    End If
    
  
                'capataz, fecha,  variedad
    Sql = "select campo1, fecha1, importe1 from tmpinformes where codusu = " & vUsu.Codigo & " order by 1,2,3"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql = "select if(isnull(sum(importe)),0,sum(importe)) + if(isnull(sum(compleme)),0,sum(compleme)) - if(isnull(sum(penaliza)),0,sum(penaliza)) as importe "
        Sql = Sql & " from horas where codcapat = " & DBSet(Rs!campo1, "N")
        Sql = Sql & " and fechahora = " & DBSet(Rs!fecha1, "F")
        Sql = Sql & " and codvarie = " & DBSet(Rs!importe1, "N")
    
        Importe = DevuelveValor(Sql)
        ImporteTot = Importe
        
'        CodigoETT = DevuelveValor("select codigoett from rcapataz where codcapat = " & DBSet(Rs!campo1, "N"))
'
'        ' si es ett tendrá registros en horasett
'        SQL = "select if(isnull(sum(importe)),0,sum(importe)) + if(isnull(sum(complemento)),0,sum(complemento)) - if(isnull(sum(penaliza)),0,sum(penaliza)) "
'        SQL = SQL & " from horasett where codcapat = " & DBSet(Rs!campo1, "N")
'        SQL = SQL & " and fechahora = " & DBSet(Rs!Fecha1, "F")
'        SQL = SQL & " and codvarie = " & DBSet(Rs!importe1, "N")
'        SQL = SQL & " and codigoett = " & DBSet(CodigoETT, "N")
'
'        Importe = DevuelveValor(SQL)
'        ImporteTot = ImporteTot + Importe
    
        Sql2 = "update tmpinformes set importe4 = " & DBSet(ImporteTot, "N")
        Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
        Sql2 = Sql2 & " and campo1 = " & DBSet(Rs!campo1, "N")
        Sql2 = Sql2 & " and fecha1 = " & DBSet(Rs!fecha1, "F")
        Sql2 = Sql2 & " and importe1 = " & DBSet(Rs!importe1, "N")
    
        conn.Execute Sql2
    
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    Screen.MousePointer = vbDefault
    
    ProcesoEntradasCapatazRdto = True
    Exit Function
    
eProcesoEntradasCapatazRdto:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Rendimiento Entradas Capataz", Err.Description
End Function


Private Sub ProcesoPaseABanco(cadWHERE As String)
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Cad As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
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
Dim Extra As String

Dim AntOpcion As Integer

On Error GoTo eProcesoPaseABanco
    
    BorrarTMPs
    CrearTMPs

    conn.BeginTrans
    
    Sql = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql
    
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    Sql = "select count(distinct rrecasesoria.codtraba) from (rrecasesoria inner join straba on rrecasesoria.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb1.visible = True
    CargarProgres Pb1, Rs.Fields(0).Value
    
    Rs.Close
    
    Sql = "select rrecasesoria.codtraba, sum(importe) importe from (rrecasesoria inner join straba on rrecasesoria.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    Sql = Sql & " group by rrecasesoria.codtraba "
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        
        '[Monica]23/03/2016: si el importe es negativo no entra
        If DBLet(Rs!Importe) >= 0 Then
            
            Sql3 = "insert into tmpImpor (codtraba, importe) values ("
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(ImporteSinFormato(CStr(Rs!Importe)), "N") & ")"
            
            conn.Execute Sql3
            
            
            Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, importe1) values (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & ","
            Sql3 = Sql3 & DBSet(txtCodigo(59).Text, "F") & "," & DBSet(ImporteSinFormato(CStr(Rs!Importe)), "N") & ")"
                
            conn.Execute Sql3
            
        End If
        
        Set Rs2 = Nothing
            
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    '[Monica]22/11/2013: iban
    Sql = "select codbanco, codsucur, digcontr, cuentaba, codorden34, sufijoem, iban from banpropi where codbanpr = " & DBSet(txtCodigo(58).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    Extra = ""
    If Rs.EOF Then
        Cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            Cad = ""
        Else
            '[Monica]22/11/2013: iban
            Cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
        Extra = DBLet(Rs!sufijoem, "T") & "|" & vParam.NombreEmpresa & "|"
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = Cad
    
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
            b = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", "Pago Nómina", Combo1(0).ListIndex, CodigoOrden34)
        Else
            b = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", "Pago Nómina", Combo1(0).ListIndex, CodigoOrden34)
        End If
    Else
        b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(0).ListIndex)
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
     
'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(0).ListIndex)
    If b Then
        b = CopiarFichero
        If b Then
            CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            CadParam = CadParam & "pFechaRecibo=""" & txtCodigo(59).Text & """|pFechaPago=""" & txtCodigo(60).Text & """|"
            numParam = 3
            cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
            cadNombreRPT = "rListadoPagos.rpt"
            cadTitulo = "Impresion de Pagos"
            ConSubInforme = False
            
            AntOpcion = OpcionListado
            OpcionListado = 0

            LlamarImprimir
            
            OpcionListado = AntOpcion
            
            If MsgBox("¿Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                Sql = "update rrecasesoria, straba, forpago set rrecasesoria.idconta = 1 where rrecasesoria.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWHERE
                conn.Execute Sql
            End If
        End If
    End If

eProcesoPaseABanco:
    If Err.Number <> 0 Then
        Mens = Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (0)
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub

Private Sub BorrarTMPs()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpImpor;"
    conn.Execute " DROP TABLE IF EXISTS tmpImporNeg;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function CrearTMPs() As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPs = False
    
    Sql = "CREATE TEMPORARY TABLE tmpImpor ( "
    Sql = Sql & "codtraba int(6) unsigned NOT NULL default '0',"
    Sql = Sql & "importe decimal(12,2)  NOT NULL default '0')"
    
    conn.Execute Sql
    
    Sql = "CREATE TEMPORARY TABLE tmpImporNeg ( "
    Sql = Sql & "codtraba int(6) unsigned NOT NULL default '0',"
    Sql = Sql & "concepto varchar(30),"
    Sql = Sql & "importe decimal(12,2)  NOT NULL default '0')"
    
    conn.Execute Sql
     
    CrearTMPs = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPs = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpImpor;"
        conn.Execute Sql
        Sql = " DROP TABLE IF EXISTS tmpImporNeg;"
        conn.Execute Sql
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


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = True
    
    If txtCodigo(59).Text = "" Or txtCodigo(60).Text = "" Then
        Sql = "Debe introducir obligatoriamente un valor en los campos de fecha. Reintroduzca. " & vbCrLf & vbCrLf
        MsgBox Sql, vbExclamation
        b = False
        PonerFoco txtCodigo(59)
    End If
    If b Then
        If txtCodigo(58).Text = "" Then
            Sql = "Debe introducir obligatoriamente un valor en el banco. Reintroduzca. " & vbCrLf & vbCrLf
            MsgBox Sql, vbExclamation
            b = False
            PonerFoco txtCodigo(58)
        End If
    End If
    '[Monica]18/09/2013: debe introducir el concepto
    If b And vParamAplic.Cooperativa = 9 Then
        If txtCodigo(66).Text = "" Then
            Sql = "Debe introducir obligatoriamente una descripción. Reintroduzca. " & vbCrLf & vbCrLf
            MsgBox Sql, vbExclamation
            b = False
            PonerFoco txtCodigo(66)
        End If
    End If
        
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I
    Combo1(0).Clear
    
    Combo1(0).AddItem "Nómina"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Pensión"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Otros Conceptos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    Combo1(1).Clear
    
    Combo1(1).AddItem "Enero"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Febrere"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    Combo1(1).AddItem "Marzo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    Combo1(1).AddItem "Abril"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 4
    Combo1(1).AddItem "Mayo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 5
    Combo1(1).AddItem "Junio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 6
    Combo1(1).AddItem "Julio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 7
    Combo1(1).AddItem "Agosto"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 8
    Combo1(1).AddItem "Septiembre"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 9
    Combo1(1).AddItem "Octubre"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 10
    Combo1(1).AddItem "Noviembre"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 11
    Combo1(1).AddItem "Diciembre"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 12
    
    
    Combo1(2).Clear
    
    Combo1(2).AddItem "Enero"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Febrero"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
    Combo1(2).AddItem "Marzo"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 3
    Combo1(2).AddItem "Abril"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 4
    Combo1(2).AddItem "Mayo"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 5
    Combo1(2).AddItem "Junio"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 6
    Combo1(2).AddItem "Julio"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 7
    Combo1(2).AddItem "Agosto"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 8
    Combo1(2).AddItem "Septiembre"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 9
    Combo1(2).AddItem "Octubre"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 10
    Combo1(2).AddItem "Noviembre"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 11
    Combo1(2).AddItem "Diciembre"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 12
    
    
    
    
End Sub


Private Function CargarTemporalListAsesoria(cadWHERE As String, Fdesde As Date, Fhasta As Date) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Cad As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
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

Dim ActTraba As String
Dim AntTraba As String

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

Dim Anticipado As Currency
Dim v_cadena As String
Dim Dias As String
Dim HayEmbargo As Boolean
Dim ImpEmbargo As Currency

On Error GoTo eCargarTemporalListAsesoria
    
    CargarTemporalListAsesoria = False
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "select distinct horas.codtraba, fechahora, sum(if(importe is null,0,importe) + if(compleme is null,0,compleme) - if(penaliza is null,0,penaliza)) importe from horas where " & cadWHERE
    Sql = Sql & " group by 1, 2 "
    Sql = Sql & " having sum(if(importe is null,0,importe) + if(compleme is null,0,compleme) - if(penaliza is null,0,penaliza)) <> 0 "
    Sql = Sql & " order by 1, 2 "
        
    Set Rs = New ADODB.Recordset
        
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        ActTraba = DBLet(Rs!CodTraba, "N")
        AntTraba = DBLet(Rs!CodTraba, "N")
    End If
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        v_cadena = String(Day(Fhasta), "N")
    Else
        v_cadena = String(31, "N") ' para Alzira
    End If
    Anticipado = 0
    Dias = 0
    HayReg = 0
    
    While Not Rs.EOF
        HayReg = 1
        Mens = "Calculando Dias" & vbCrLf & vbCrLf & "Trabajador: " & ActTraba & vbCrLf
        ActTraba = DBLet(Rs!CodTraba, "N")
        If ActTraba <> AntTraba Then
                                                
            ' calculamos el importe anticipado de lo que tenemos guardado en rrecibosnomina
            Sql = "select sum(neto34) from rrecibosnomina where codtraba = " & DBSet(AntTraba, "N")
            Sql = Sql & " and fechahora >= " & DBSet(Fdesde, "F")
            Sql = Sql & " and fechahora <= " & DBSet(Fhasta, "F")
            
            '[Monica]04/11/2016: y que no haya sido embargado
            Sql = Sql & " and hayembargo = 0 "
                                                
            Anticipado = DevuelveValor(Sql)
                                                
            Sql = "select sum(importe) from rrecibosnomina where codtraba = " & DBSet(AntTraba, "N")
            Sql = Sql & " and fechahora >= " & DBSet(Fdesde, "F")
            Sql = Sql & " and fechahora <= " & DBSet(Fhasta, "F")
                                                
            Bruto = DevuelveValor(Sql)
                                                
            ImpEmbargo = 0
'            Sql = "select hayembargo from straba where codtraba = " & DBSet(AntTraba, "N")
'            HayEmbargo = (DevuelveValor(Sql) = "1")
'            If HayEmbargo Then
'                Sql = "select impembargo from straba where codtraba = " & DBSet(AntTraba, "N")
'                ImpEmbargo = DevuelveValor(Sql)
'            End If
            
            Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, nombre1, importe1, importe2, importe3, importe4) values ("
            Sql3 = Sql3 & vUsu.Codigo & ","
            Sql3 = Sql3 & DBSet(AntTraba, "N") & ","
            Sql3 = Sql3 & DBSet(Fhasta, "F") & ","
            Sql3 = Sql3 & DBSet(v_cadena, "T") & ","
            Sql3 = Sql3 & DBSet(Anticipado, "N") & ","
            Sql3 = Sql3 & DBSet(Dias, "N") & ","
            Sql3 = Sql3 & DBSet(Bruto, "N") & ","
            Sql3 = Sql3 & DBSet(ImpEmbargo, "N") & ")"
            
            conn.Execute Sql3

            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                v_cadena = String(Day(Fhasta), "N")
            Else
                v_cadena = String(31, "N") ' para Alzira
            End If
            
            AntTraba = ActTraba
            Anticipado = 0
            Dias = 0
        End If
        
        i = Day(DBLet(Rs.Fields(1).Value, "N"))
        If i = 1 Then
            v_cadena = "S" & Mid(v_cadena, 2, Len(v_cadena)) ' Replace(v_cadena, "N", "S", I, 1)
        Else
            v_cadena = Mid(v_cadena, 1, i - 1) & Replace(v_cadena, "N", "S", i, 1)
        End If
        Dias = Dias + 1
        
        Anticipado = Anticipado + DBLet(Rs!Importe, "N")
        
        Rs.MoveNext
    Wend
    If HayReg = 1 Then
        ' calculamos el importe anticipado de lo que tenemos guardado en rrecibosnomina
        Sql = "select sum(neto34) from rrecibosnomina where codtraba = " & DBSet(ActTraba, "N")
        Sql = Sql & " and fechahora >= " & DBSet(Fdesde, "F")
        Sql = Sql & " and fechahora <= " & DBSet(Fhasta, "F")
        '[Monica]04/11/2016: y que no haya sido embargado
        Sql = Sql & " and hayembargo = 0 "
                                            
        Anticipado = DevuelveValor(Sql)
                                            
        Sql = "select sum(importe) from rrecibosnomina where codtraba = " & DBSet(ActTraba, "N")
        Sql = Sql & " and fechahora >= " & DBSet(Fdesde, "F")
        Sql = Sql & " and fechahora <= " & DBSet(Fhasta, "F")
                                            
        Bruto = DevuelveValor(Sql)
        
        ImpEmbargo = 0
'        Sql = "select hayembargo from straba where codtraba = " & DBSet(ActTraba, "N")
'        HayEmbargo = (DevuelveValor(Sql) = "1")
'        If HayEmbargo Then
'            Sql = "select impembargo from straba where codtraba = " & DBSet(ActTraba, "N")
'            ImpEmbargo = DevuelveValor(Sql)
'        End If
        
        
        Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, nombre1, importe1, importe2, importe3, importe4) values ("
        Sql3 = Sql3 & vUsu.Codigo & ","
        Sql3 = Sql3 & DBSet(ActTraba, "N") & ","
        Sql3 = Sql3 & DBSet(Fhasta, "F") & ","
        Sql3 = Sql3 & DBSet(v_cadena, "T") & ","
        Sql3 = Sql3 & DBSet(Anticipado, "N") & ","
        Sql3 = Sql3 & DBSet(Dias, "N") & ","
        Sql3 = Sql3 & DBSet(Bruto, "N") & ","
        Sql3 = Sql3 & DBSet(ImpEmbargo, "N") & ")"
        
        conn.Execute Sql3
    End If
    Set Rs = Nothing
    
    CargarTemporalListAsesoria = True
    Exit Function
    
eCargarTemporalListAsesoria:
    If Err.Number <> 0 Then
        Mens = Err.Description
        MsgBox "Error " & Mens, vbExclamation
    End If
End Function




Private Sub ProcesoPaseABancoAnticipos(cadWHERE As String)
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Cad As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
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
Dim Extra As String

Dim AntOpcion As Integer

On Error GoTo eProcesoPaseABanco
    
    BorrarTMPs
    CrearTMPs

    conn.BeginTrans
    
    Sql = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql
    
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    Sql = "select count(distinct horasanticipos.codtraba) from (horasanticipos inner join straba on horasanticipos.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb1.visible = True
    CargarProgres Pb1, Rs.Fields(0).Value
    
    Rs.Close
    
    Sql = "select horasanticipos.codtraba, sum(importe) importe from (horasanticipos inner join straba on horasanticipos.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    Sql = Sql & " group by horasanticipos.codtraba "
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        
        '[Monica]23/03/2016: si el importe es negativo no entra
        If DBLet(Rs!Importe) >= 0 Then
        
            Sql3 = "insert into tmpImpor (codtraba, importe) values ("
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(ImporteSinFormato(CStr(Rs!Importe)), "N") & ")"
            
            conn.Execute Sql3
            
            
            Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, importe1) values (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & ","
            Sql3 = Sql3 & DBSet(txtCodigo(59).Text, "F") & "," & DBSet(ImporteSinFormato(CStr(Rs!Importe)), "N") & ")"
                
            conn.Execute Sql3
            
        End If
        
        Set Rs2 = Nothing
            
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    '[Monica]22/11/2013: iban
    Sql = "select codbanco, codsucur, digcontr, cuentaba, codorden34, sufijoem, iban from banpropi where codbanpr = " & DBSet(txtCodigo(58).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    Extra = ""
    If Rs.EOF Then
        Cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            Cad = ""
        Else
            '[Monica]22/11/2013: iban
            Cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
        Extra = DBLet(Rs!sufijoem, "T") & "|" & vParam.NombreEmpresa & "|"
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = Cad
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
            b = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", txtCodigo(66).Text, Combo1(0).ListIndex, CodigoOrden34)
        Else
            b = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", txtCodigo(66).Text, Combo1(0).ListIndex, CodigoOrden34)
        End If
    Else
        b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, txtCodigo(66).Text, CodigoOrden34, Combo1(0).ListIndex)
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
     
'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, txtCodigo(66).Text, CodigoOrden34, Combo1(0).ListIndex)
    If b Then
        b = CopiarFichero
        If b Then
            CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            CadParam = CadParam & "pFechaRecibo=""" & txtCodigo(59).Text & """|pFechaPago=""" & txtCodigo(60).Text & """|"
            numParam = 3
            cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
            cadNombreRPT = "rListadoPagos.rpt"
            cadTitulo = "Impresion de Pagos"
            ConSubInforme = False
            
            AntOpcion = OpcionListado
            OpcionListado = 0

            LlamarImprimir
            
            OpcionListado = AntOpcion
            
            If MsgBox("¿Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                Sql = "update horasanticipos, straba, forpago set horasanticipos.fechapago = " & DBSet(txtCodigo(60).Text, "F")
                Sql = Sql & ", concepto = " & DBSet(Trim(txtCodigo(66).Text), "T")
                Sql = Sql & " where horasanticipos.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWHERE
                conn.Execute Sql
            Else
                conn.RollbackTrans
                cmdCancel_Click (0)
                Exit Sub
            End If
        End If
    End If

eProcesoPaseABanco:
    If Err.Number <> 0 Then
        Mens = Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (0)
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub



Private Function CargarTemporalListDiasTrabajados(cadWHERE As String, Fdesde As Date, Fhasta As Date, cadWHERE2 As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Cad As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
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

Dim ActTraba As String
Dim AntTraba As String

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

Dim Anticipado As Currency
Dim v_cadena As String
Dim Dias As String
Dim HayEmbargo As Boolean
Dim ImpEmbargo As Currency

On Error GoTo eCargarTemporalListAsesoria
    
    CargarTemporalListDiasTrabajados = False
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
    
    If cadWHERE2 <> "" Then
        cadWHERE2 = QuitarCaracterACadena(cadWHERE2, "{")
        cadWHERE2 = QuitarCaracterACadena(cadWHERE2, "}")
        cadWHERE2 = QuitarCaracterACadena(cadWHERE2, "_1")
    End If
    
    
        
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "select codtraba, fecha from ("
    
    Sql = Sql & "select distinct rpartes_trabajador.codtraba, rpartes.fecentrada fecha, sum(coalesce(rpartes_trabajador.importe,0)) from rpartes inner join rpartes_trabajador on rpartes.nroparte = rpartes_trabajador.nroparte where " & cadWHERE
    Sql = Sql & " group by 1, 2 "
    Sql = Sql & " having sum(coalesce(rpartes_trabajador.importe,0)) <> 0 "
    
    Sql = Sql & " union "
    Sql = Sql & "select distinct horas.codtraba, horas.fechahora fecha , sum(if(importe is null,0,importe) + if(compleme is null,0,compleme) - if(penaliza is null,0,penaliza)) importe from  horas where " & cadWHERE2
    Sql = Sql & " group by 1, 2 "
    Sql = Sql & " having  sum(if(importe is null,0,importe) + if(compleme is null,0,compleme) - if(penaliza is null,0,penaliza)) <> 0 "
    Sql = Sql & " order by 1, 2 "
    
    Sql = Sql & ") aaaaaa "
    
    Sql = Sql & " order by 1, 2 "
        
    Set Rs = New ADODB.Recordset
        
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        ActTraba = DBLet(Rs!CodTraba, "N")
        AntTraba = DBLet(Rs!CodTraba, "N")
    End If
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        v_cadena = String(Day(Fhasta), "N")
    Else
        v_cadena = String(31, "N") ' para Alzira
    End If
    Anticipado = 0
    Dias = 0
    HayReg = 0
    
    While Not Rs.EOF
        HayReg = 1
        Mens = "Calculando Dias" & vbCrLf & vbCrLf & "Trabajador: " & ActTraba & vbCrLf
        ActTraba = DBLet(Rs!CodTraba, "N")
        If ActTraba <> AntTraba Then
                                                
            ' calculamos el importe anticipado de lo que tenemos guardado en rrecibosnomina
            Sql = "select sum(neto34) from rrecibosnomina where codtraba = " & DBSet(AntTraba, "N")
            Sql = Sql & " and fechahora >= " & DBSet(Fdesde, "F")
            Sql = Sql & " and fechahora <= " & DBSet(Fhasta, "F")
                                                
            Anticipado = 0 'DevuelveValor(Sql)
                                                
            Sql = "select sum(importe) from rrecibosnomina where codtraba = " & DBSet(AntTraba, "N")
            Sql = Sql & " and fechahora >= " & DBSet(Fdesde, "F")
            Sql = Sql & " and fechahora <= " & DBSet(Fhasta, "F")
                                                
            Bruto = 0 'DevuelveValor(Sql)
                                                
            ImpEmbargo = 0
'            Sql = "select hayembargo from straba where codtraba = " & DBSet(AntTraba, "N")
'            HayEmbargo = (DevuelveValor(Sql) = "1")
'            If HayEmbargo Then
'                Sql = "select impembargo from straba where codtraba = " & DBSet(AntTraba, "N")
'                ImpEmbargo = DevuelveValor(Sql)
'            End If
            
            Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, nombre1, importe1, importe2, importe3, importe4) values ("
            Sql3 = Sql3 & vUsu.Codigo & ","
            Sql3 = Sql3 & DBSet(AntTraba, "N") & ","
            Sql3 = Sql3 & DBSet(Fhasta, "F") & ","
            Sql3 = Sql3 & DBSet(v_cadena, "T") & ","
            Sql3 = Sql3 & DBSet(Anticipado, "N") & ","
            Sql3 = Sql3 & DBSet(Dias, "N") & ","
            Sql3 = Sql3 & DBSet(Bruto, "N") & ","
            Sql3 = Sql3 & DBSet(ImpEmbargo, "N") & ")"
            
            conn.Execute Sql3

            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                v_cadena = String(Day(Fhasta), "N")
            Else
                v_cadena = String(31, "N") ' para Alzira
            End If
            
            AntTraba = ActTraba
            Anticipado = 0
            Dias = 0
        End If
        
        i = Day(DBLet(Rs.Fields(1).Value, "N"))
        If i = 1 Then
            v_cadena = "S" & Mid(v_cadena, 2, Len(v_cadena)) ' Replace(v_cadena, "N", "S", I, 1)
        Else
            v_cadena = Mid(v_cadena, 1, i - 1) & Replace(v_cadena, "N", "S", i, 1)
        End If
        Dias = Dias + 1
        
'        Anticipado = Anticipado + DBLet(Rs!Importe, "N")
        
        Rs.MoveNext
    Wend
    If HayReg = 1 Then
        ' calculamos el importe anticipado de lo que tenemos guardado en rrecibosnomina
        Sql = "select sum(neto34) from rrecibosnomina where codtraba = " & DBSet(ActTraba, "N")
        Sql = Sql & " and fechahora >= " & DBSet(Fdesde, "F")
        Sql = Sql & " and fechahora <= " & DBSet(Fhasta, "F")
                                            
        Anticipado = 0 'DevuelveValor(Sql)
                                            
        Sql = "select sum(importe) from rrecibosnomina where codtraba = " & DBSet(ActTraba, "N")
        Sql = Sql & " and fechahora >= " & DBSet(Fdesde, "F")
        Sql = Sql & " and fechahora <= " & DBSet(Fhasta, "F")
                                            
        Bruto = 0 'DevuelveValor(Sql)
        
        ImpEmbargo = 0
'        Sql = "select hayembargo from straba where codtraba = " & DBSet(ActTraba, "N")
'        HayEmbargo = (DevuelveValor(Sql) = "1")
'        If HayEmbargo Then
'            Sql = "select impembargo from straba where codtraba = " & DBSet(ActTraba, "N")
'            ImpEmbargo = DevuelveValor(Sql)
'        End If
        
        
        Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, nombre1, importe1, importe2, importe3, importe4) values ("
        Sql3 = Sql3 & vUsu.Codigo & ","
        Sql3 = Sql3 & DBSet(ActTraba, "N") & ","
        Sql3 = Sql3 & DBSet(Fhasta, "F") & ","
        Sql3 = Sql3 & DBSet(v_cadena, "T") & ","
        Sql3 = Sql3 & DBSet(Anticipado, "N") & ","
        Sql3 = Sql3 & DBSet(Dias, "N") & ","
        Sql3 = Sql3 & DBSet(Bruto, "N") & ","
        Sql3 = Sql3 & DBSet(ImpEmbargo, "N") & ")"
        
        conn.Execute Sql3
    End If
    Set Rs = Nothing
    
    CargarTemporalListDiasTrabajados = True
    Exit Function
    
eCargarTemporalListAsesoria:
    If Err.Number <> 0 Then
        Mens = Err.Description
        MsgBox "Error " & Mens, vbExclamation
    End If
End Function


