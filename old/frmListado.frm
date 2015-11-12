VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7080
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTraspasoCalibrador 
      Height          =   4665
      Left            =   30
      TabIndex        =   194
      Top             =   60
      Width           =   6555
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   285
         Left            =   240
         TabIndex        =   202
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
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   200
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   1620
         Width           =   2295
      End
      Begin VB.CommandButton cmdAcepTras 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   196
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelTras 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   195
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
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   51
         Left            =   690
         TabIndex        =   201
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Proceso que realiza el Traspaso desde el Calibrador seleccionado de la clasificación de entradas."
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
         TabIndex        =   199
         Top             =   630
         Width           =   5820
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   198
         Top             =   3480
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   197
         Top             =   3120
         Width           =   6195
      End
   End
   Begin VB.Frame FrameKilosProducto 
      Height          =   6480
      Left            =   0
      TabIndex        =   158
      Top             =   30
      Width           =   6615
      Begin VB.CheckBox Check3 
         Caption         =   "Salta página Producto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4140
         TabIndex        =   193
         Top             =   5070
         Width           =   2085
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   171
         Top             =   3210
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   172
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
         TabIndex        =   189
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
         TabIndex        =   188
         Text            =   "Text5"
         Top             =   3570
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   175
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   173
         Top             =   4050
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelInf 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5190
         TabIndex        =   181
         Top             =   5625
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepInf 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   179
         Top             =   5625
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   167
         Top             =   1275
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   168
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
         TabIndex        =   166
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
         TabIndex        =   165
         Text            =   "Text5"
         Top             =   1635
         Width           =   3375
      End
      Begin VB.CommandButton Command10 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command9 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   169
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   170
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
         TabIndex        =   162
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
         TabIndex        =   161
         Text            =   "Text5"
         Top             =   2580
         Width           =   3375
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   585
         TabIndex        =   159
         Top             =   4860
         Width           =   3480
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   5
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
            Top             =   135
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Hectáreas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   38
            Left            =   90
            TabIndex        =   160
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   45
         Left            =   675
         TabIndex        =   192
         Top             =   3015
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   44
         Left            =   1005
         TabIndex        =   191
         Top             =   3255
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   43
         Left            =   1005
         TabIndex        =   190
         Top             =   3645
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   37
         Left            =   1620
         MouseIcon       =   "frmListado.frx":0620
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   38
         Left            =   1620
         MouseIcon       =   "frmListado.frx":0772
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   50
         Left            =   675
         TabIndex        =   187
         Top             =   3870
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   1005
         TabIndex        =   186
         Top             =   4110
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   48
         Left            =   1005
         TabIndex        =   185
         Top             =   4455
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   960
         TabIndex        =   184
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   960
         TabIndex        =   183
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
         TabIndex        =   182
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   42
         Left            =   675
         TabIndex        =   180
         Top             =   1080
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   33
         Left            =   1620
         MouseIcon       =   "frmListado.frx":08C4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   34
         Left            =   1620
         MouseIcon       =   "frmListado.frx":0A16
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1620
         Picture         =   "frmListado.frx":0B68
         ToolTipText     =   "Buscar fecha"
         Top             =   4455
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1620
         Picture         =   "frmListado.frx":0BF3
         ToolTipText     =   "Buscar fecha"
         Top             =   4050
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   41
         Left            =   675
         TabIndex        =   178
         Top             =   2025
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   40
         Left            =   1005
         TabIndex        =   176
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   39
         Left            =   1005
         TabIndex        =   174
         Top             =   2655
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   35
         Left            =   1620
         MouseIcon       =   "frmListado.frx":0C7E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   36
         Left            =   1620
         MouseIcon       =   "frmListado.frx":0DD0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2610
         Width           =   240
      End
   End
   Begin VB.Frame FrameEntradasCampo 
      Height          =   6480
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check9 
         Caption         =   "Detallar Notas"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   309
         Top             =   4490
         Width           =   1815
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Salta página por socio"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   251
         Top             =   5190
         Width           =   2205
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Clasificado por Socio"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   250
         Top             =   4840
         Width           =   1815
      End
      Begin VB.Frame FrameTipo 
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   585
         TabIndex        =   134
         Top             =   4860
         Width           =   3480
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
            Top             =   585
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   94
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
            TabIndex        =   136
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
            TabIndex        =   135
            Top             =   180
            Width           =   1110
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   130
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
         TabIndex        =   129
         Text            =   "Text5"
         Top             =   2220
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   89
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   88
         Top             =   2220
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":0F22
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":122C
         Style           =   1  'Graphical
         TabIndex        =   117
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
         TabIndex        =   116
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
         TabIndex        =   115
         Text            =   "Text5"
         Top             =   3210
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   91
         Top             =   3570
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   90
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
         TabIndex        =   114
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
         TabIndex        =   99
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   87
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   86
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   4110
         TabIndex        =   96
         Top             =   5625
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5190
         TabIndex        =   97
         Top             =   5625
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   93
         Top             =   4545
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   92
         Top             =   4140
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Imprimir Resumen"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   98
         Top             =   4140
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1536
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1688
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
         TabIndex        =   133
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   1005
         TabIndex        =   132
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   675
         TabIndex        =   131
         Top             =   2025
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1620
         Picture         =   "frmListado.frx":17DA
         ToolTipText     =   "Buscar fecha"
         Top             =   4140
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1620
         Picture         =   "frmListado.frx":1865
         ToolTipText     =   "Buscar fecha"
         Top             =   4545
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1620
         MouseIcon       =   "frmListado.frx":18F0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1A42
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1B94
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1620
         MouseIcon       =   "frmListado.frx":1CE6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   675
         TabIndex        =   128
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   1005
         TabIndex        =   127
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   1005
         TabIndex        =   126
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
         TabIndex        =   125
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   675
         TabIndex        =   124
         Top             =   3015
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   960
         TabIndex        =   123
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   960
         TabIndex        =   122
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   1005
         TabIndex        =   121
         Top             =   4545
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   1005
         TabIndex        =   120
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   675
         TabIndex        =   119
         Top             =   3960
         Width           =   450
      End
   End
   Begin VB.Frame FrameGrabacionAgriweb 
      Height          =   6735
      Left            =   0
      TabIndex        =   137
      Top             =   0
      Width           =   6675
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   32
         Left            =   2610
         MaxLength       =   5
         TabIndex        =   111
         Tag             =   "Campol|N|S|0|99.99|clientes|codposta|00.00||"
         Top             =   5400
         Width           =   1200
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   102
         Top             =   1830
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   103
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
         TabIndex        =   153
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
         TabIndex        =   152
         Text            =   "Text5"
         Top             =   2205
         Width           =   3675
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   4
         Left            =   2610
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   28
         Left            =   2610
         MaxLength       =   9
         TabIndex        =   107
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3870
         Width           =   1200
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   29
         Left            =   2610
         MaxLength       =   13
         TabIndex        =   108
         Tag             =   "Campol|N|S|||clientes|codposta|#,###,###,###||"
         Top             =   4260
         Width           =   1200
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   104
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
         TabIndex        =   147
         Text            =   "Text5"
         Top             =   2595
         Width           =   3735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   30
         Left            =   2610
         MaxLength       =   10
         TabIndex        =   109
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4650
         Width           =   1200
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   31
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   110
         Tag             =   "Campol|N|S|||clientes|codposta|#,##0.00||"
         Top             =   5025
         Width           =   1200
      End
      Begin VB.CommandButton CmdCancelAgri 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5250
         TabIndex        =   113
         Top             =   6060
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepAgri 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3990
         TabIndex        =   112
         Top             =   6060
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   100
         Top             =   975
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   101
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
         TabIndex        =   139
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
         TabIndex        =   138
         Text            =   "Text5"
         Top             =   1350
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   27
         Left            =   2610
         MaxLength       =   4
         TabIndex        =   105
         Tag             =   "Campol|N|S|||clientes|codposta|0000||"
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "Precio Estipulado Compra"
         ForeColor       =   &H00972E0B&
         Height          =   285
         Index           =   39
         Left            =   390
         TabIndex        =   157
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   780
         TabIndex        =   156
         Top             =   1860
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   780
         TabIndex        =   155
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
         TabIndex        =   154
         Top             =   1620
         Width           =   390
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1425
         MouseIcon       =   "frmListado.frx":1E38
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1830
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   1425
         MouseIcon       =   "frmListado.frx":1F8A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2205
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Superficie Total Contrato"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   151
         Top             =   5055
         Width           =   2025
      End
      Begin VB.Label Label4 
         Caption         =   "CIF Industria transformadora"
         ForeColor       =   &H00972E0B&
         Height          =   285
         Index           =   29
         Left            =   390
         TabIndex        =   150
         Top             =   3870
         Width           =   2595
      End
      Begin VB.Label Label4 
         Caption         =   "Kgs. Contratados"
         ForeColor       =   &H00972E0B&
         Height          =   285
         Index           =   36
         Left            =   390
         TabIndex        =   149
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
         TabIndex        =   148
         Top             =   2610
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   22
         Left            =   1410
         MouseIcon       =   "frmListado.frx":20DC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   2595
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Grabación Fichero Agriweb"
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
         TabIndex        =   146
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Formalización"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   30
         Left            =   390
         TabIndex        =   145
         Top             =   4680
         Width           =   1485
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   11
         Left            =   2250
         Picture         =   "frmListado.frx":222E
         ToolTipText     =   "Buscar fecha"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   33
         Left            =   795
         TabIndex        =   144
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   34
         Left            =   795
         TabIndex        =   143
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
         TabIndex        =   142
         Top             =   765
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   1440
         MouseIcon       =   "frmListado.frx":22B9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   975
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1440
         MouseIcon       =   "frmListado.frx":240B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Ejercicio"
         ForeColor       =   &H00972E0B&
         Height          =   285
         Index           =   27
         Left            =   390
         TabIndex        =   141
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Producto según tabla"
         ForeColor       =   &H00972E0B&
         Height          =   315
         Index           =   28
         Left            =   390
         TabIndex        =   140
         Top             =   3480
         Width           =   1665
      End
   End
   Begin VB.Frame FrameTraspasoROPAS 
      Height          =   4890
      Left            =   0
      TabIndex        =   286
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   1830
         MaxLength       =   3
         TabIndex        =   289
         Top             =   2265
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   1830
         MaxLength       =   3
         TabIndex        =   290
         Top             =   2625
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   60
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   303
         Text            =   "Text5"
         Top             =   2265
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   61
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   302
         Text            =   "Text5"
         Top             =   2625
         Width           =   3375
      End
      Begin VB.CommandButton Command19 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":255D
         Style           =   1  'Graphical
         TabIndex        =   297
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command16 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":2867
         Style           =   1  'Graphical
         TabIndex        =   296
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   59
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   295
         Text            =   "Text5"
         Top             =   1695
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   58
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   294
         Text            =   "Text5"
         Top             =   1335
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   59
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   288
         Top             =   1695
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   287
         ToolTipText     =   " "
         Top             =   1335
         Width           =   750
      End
      Begin VB.CommandButton cmdAcepROPAS 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   292
         Top             =   4095
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelROPAS 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5280
         TabIndex        =   293
         Top             =   4095
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   62
         Left            =   1830
         MaxLength       =   4
         TabIndex        =   291
         Top             =   3105
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ejercicio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   79
         Left            =   570
         TabIndex        =   307
         Top             =   3150
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   87
         Left            =   570
         TabIndex        =   306
         Top             =   2070
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   86
         Left            =   900
         TabIndex        =   305
         Top             =   2310
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   52
         Left            =   900
         TabIndex        =   304
         Top             =   2700
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   43
         Left            =   1530
         MouseIcon       =   "frmListado.frx":2B71
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1740
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   44
         Left            =   1515
         MouseIcon       =   "frmListado.frx":2CC3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   42
         Left            =   1530
         MouseIcon       =   "frmListado.frx":2E15
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   45
         Left            =   1530
         MouseIcon       =   "frmListado.frx":2F67
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   85
         Left            =   585
         TabIndex        =   301
         Top             =   1140
         Width           =   405
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
         TabIndex        =   300
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   84
         Left            =   870
         TabIndex        =   299
         Top             =   1740
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   83
         Left            =   870
         TabIndex        =   298
         Top             =   1380
         Width           =   465
      End
   End
   Begin VB.Frame FrameTraspasoFactCoop 
      Height          =   5490
      Left            =   0
      TabIndex        =   222
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   45
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   239
         Text            =   "Text5"
         Top             =   1095
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   45
         Left            =   1890
         MaxLength       =   2
         TabIndex        =   238
         Top             =   1095
         Width           =   750
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   7
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   247
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   4380
         Width           =   2115
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   243
         Top             =   2580
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   244
         Top             =   2985
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelTrasCoop 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5280
         TabIndex        =   249
         Top             =   4695
         Width           =   975
      End
      Begin VB.CommandButton cmdAcepTrasCoop 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   4
         Left            =   4200
         TabIndex        =   248
         Top             =   4695
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   48
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   240
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   242
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
         TabIndex        =   226
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
         TabIndex        =   225
         Text            =   "Text5"
         Top             =   2025
         Width           =   3375
      End
      Begin VB.CommandButton Command14 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":30B9
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command13 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":33C3
         Style           =   1  'Graphical
         TabIndex        =   223
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   246
         Top             =   3930
         Width           =   1065
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   245
         Top             =   3540
         Width           =   1065
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1590
         MouseIcon       =   "frmListado.frx":36CD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cooperativa"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   60
         Left            =   630
         TabIndex        =   241
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   54
         Left            =   630
         TabIndex        =   237
         Top             =   4425
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   68
         Left            =   645
         TabIndex        =   236
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   67
         Left            =   975
         TabIndex        =   235
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   65
         Left            =   975
         TabIndex        =   234
         Top             =   2985
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   64
         Left            =   930
         TabIndex        =   233
         Top             =   1710
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   63
         Left            =   930
         TabIndex        =   232
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
         TabIndex        =   231
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   59
         Left            =   645
         TabIndex        =   230
         Top             =   1470
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   30
         Left            =   1590
         MouseIcon       =   "frmListado.frx":381F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   29
         Left            =   1590
         MouseIcon       =   "frmListado.frx":3971
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1590
         Picture         =   "frmListado.frx":3AC3
         ToolTipText     =   "Buscar fecha"
         Top             =   2985
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1590
         Picture         =   "frmListado.frx":3B4E
         ToolTipText     =   "Buscar fecha"
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   57
         Left            =   675
         TabIndex        =   229
         Top             =   3375
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   56
         Left            =   1005
         TabIndex        =   228
         Top             =   3615
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   55
         Left            =   1005
         TabIndex        =   227
         Top             =   4005
         Width           =   420
      End
   End
   Begin VB.Frame FrameSociosSeccion 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   7020
      Begin VB.CheckBox Check8 
         Caption         =   "Imprimir Socios de baja"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3450
         TabIndex        =   308
         Top             =   3300
         Width           =   2355
      End
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Ordenar por"
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
            Caption         =   "Sección"
            Height          =   345
            Index           =   0
            Left            =   495
            TabIndex        =   22
            Top             =   225
            Width           =   1290
         End
      End
      Begin VB.CommandButton Command6 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":3BD9
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command5 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":3EE3
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
         Width           =   3015
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
         Width           =   3015
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   8
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
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
         Width           =   3015
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
         Width           =   3015
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   3510
         TabIndex        =   2
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   1
         Top             =   3720
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
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1575
         MouseIcon       =   "frmListado.frx":41ED
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1560
         MouseIcon       =   "frmListado.frx":433F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1560
         MouseIcon       =   "frmListado.frx":4491
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar sección"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1560
         MouseIcon       =   "frmListado.frx":45E3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar sección"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   17
         Left            =   600
         TabIndex        =   20
         Top             =   2160
         Width           =   405
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
         Caption         =   "Informe de Socios por Sección"
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
         Caption         =   "Sección"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   14
         Left            =   600
         TabIndex        =   16
         Top             =   1080
         Width           =   585
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
   Begin VB.Frame FrameKilosRecolect 
      Height          =   6480
      Left            =   30
      TabIndex        =   252
      Top             =   30
      Width           =   6615
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   57
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   269
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
         TabIndex        =   267
         Text            =   "Text5"
         Top             =   2220
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   57
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   258
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   56
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   257
         Top             =   2220
         Width           =   735
      End
      Begin VB.CommandButton Command18 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":4735
         Style           =   1  'Graphical
         TabIndex        =   265
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command17 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":4A3F
         Style           =   1  'Graphical
         TabIndex        =   263
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
         TabIndex        =   261
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
         TabIndex        =   259
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   55
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   256
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   54
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   255
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepKilosSoc 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   270
         Top             =   5745
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelKilosSoc 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5160
         TabIndex        =   272
         Top             =   5745
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   53
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   266
         Top             =   4470
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   52
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   264
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
         TabIndex        =   254
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
         TabIndex        =   253
         Text            =   "Text5"
         Top             =   3210
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   51
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   262
         Top             =   3570
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   50
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   260
         Top             =   3210
         Width           =   735
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Incluir pendiente de clasificar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   690
         TabIndex        =   268
         Top             =   4920
         Width           =   2565
      End
      Begin MSComctlLib.ProgressBar Pb3 
         Height          =   255
         Left            =   450
         TabIndex        =   285
         Top             =   5310
         Visible         =   0   'False
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   41
         Left            =   1620
         MouseIcon       =   "frmListado.frx":4D49
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   32
         Left            =   1620
         MouseIcon       =   "frmListado.frx":4E9B
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
         TabIndex        =   284
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   77
         Left            =   1005
         TabIndex        =   283
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   76
         Left            =   675
         TabIndex        =   282
         Top             =   2025
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1620
         Picture         =   "frmListado.frx":4FED
         ToolTipText     =   "Buscar fecha"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1620
         Picture         =   "frmListado.frx":5078
         ToolTipText     =   "Buscar fecha"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   31
         Left            =   1620
         MouseIcon       =   "frmListado.frx":5103
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1620
         MouseIcon       =   "frmListado.frx":5255
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   75
         Left            =   675
         TabIndex        =   281
         Top             =   1080
         Width           =   405
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
         TabIndex        =   280
         Top             =   420
         Width           =   5595
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   74
         Left            =   960
         TabIndex        =   279
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   73
         Left            =   960
         TabIndex        =   278
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   72
         Left            =   1005
         TabIndex        =   277
         Top             =   4455
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   71
         Left            =   1005
         TabIndex        =   276
         Top             =   4110
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   70
         Left            =   675
         TabIndex        =   275
         Top             =   3870
         Width           =   450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1620
         MouseIcon       =   "frmListado.frx":53A7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   40
         Left            =   1620
         MouseIcon       =   "frmListado.frx":54F9
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
         TabIndex        =   274
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   62
         Left            =   1005
         TabIndex        =   273
         Top             =   3255
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   61
         Left            =   675
         TabIndex        =   271
         Top             =   3015
         Width           =   645
      End
   End
   Begin VB.Frame FrameCampos 
      Height          =   6525
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check4 
         Caption         =   "Imprimir Cabecera Cooperativa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3150
         TabIndex        =   221
         Top             =   5280
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   4725
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   4365
         Width           =   1440
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3150
         TabIndex        =   82
         Top             =   4890
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   3150
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   4365
         Width           =   1440
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   63
         Top             =   3525
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1935
         MaxLength       =   2
         TabIndex        =   62
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text5"
         Top             =   3525
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text5"
         Top             =   3120
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5100
         TabIndex        =   66
         Top             =   5850
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4020
         TabIndex        =   64
         Top             =   5850
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   59
         Top             =   1680
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   58
         Top             =   1275
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
         Top             =   1680
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
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   61
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   60
         Top             =   2175
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
         Top             =   2580
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
         Top             =   2175
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":564B
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command1 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":5955
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Frame Frame3 
         Caption         =   "Clasificado por"
         ForeColor       =   &H00972E0B&
         Height          =   1740
         Left            =   495
         TabIndex        =   49
         Top             =   4005
         Width           =   2460
         Begin VB.OptionButton Opcion1 
            Caption         =   "Zona"
            Height          =   255
            Index           =   3
            Left            =   495
            TabIndex        =   74
            Top             =   1260
            Width           =   1470
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Termino Municipal"
            Height          =   255
            Index           =   2
            Left            =   495
            TabIndex        =   73
            Top             =   945
            Width           =   1605
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Socio"
            Height          =   345
            Index           =   0
            Left            =   495
            TabIndex        =   51
            Top             =   225
            Width           =   1290
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Clase/Variedad"
            Height          =   255
            Index           =   1
            Left            =   495
            TabIndex        =   50
            Top             =   585
            Width           =   1470
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Producción"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   10
         Left            =   4770
         TabIndex        =   84
         Top             =   4050
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Hectáreas"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   9
         Left            =   3150
         TabIndex        =   81
         Top             =   4050
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Situación"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   8
         Left            =   630
         TabIndex        =   79
         Top             =   2970
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   78
         Top             =   3165
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   77
         Top             =   3555
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1560
         MouseIcon       =   "frmListado.frx":5C5F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar situación"
         Top             =   3525
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1575
         MouseIcon       =   "frmListado.frx":5DB1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar situación"
         Top             =   3150
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   72
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   71
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   630
         TabIndex        =   70
         Top             =   2025
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
         TabIndex        =   69
         Top             =   405
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   68
         Top             =   2220
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   67
         Top             =   2610
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   65
         Top             =   1080
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1560
         MouseIcon       =   "frmListado.frx":5F03
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1560
         MouseIcon       =   "frmListado.frx":6055
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1560
         MouseIcon       =   "frmListado.frx":61A7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1575
         MouseIcon       =   "frmListado.frx":62F9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2205
         Width           =   240
      End
   End
   Begin VB.Frame FrameCalidades 
      Height          =   4455
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7020
      Begin VB.CommandButton cmdCancel 
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
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   34
         Top             =   1275
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
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
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   36
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
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
         Picture         =   "frmListado.frx":644B
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command7 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":6755
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ordenar por"
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
         MouseIcon       =   "frmListado.frx":6A5F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1575
         MouseIcon       =   "frmListado.frx":6BB1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1575
         MouseIcon       =   "frmListado.frx":6D03
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar calidad"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1575
         MouseIcon       =   "frmListado.frx":6E55
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar calidad"
         Top             =   2430
         Width           =   240
      End
   End
   Begin VB.Frame FrameBajaSocios 
      Height          =   3150
      Left            =   0
      TabIndex        =   210
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   47
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   215
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelBajaSocio 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5130
         TabIndex        =   217
         Top             =   2325
         Width           =   975
      End
      Begin VB.CommandButton cmdAcepBajaSocio 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4020
         TabIndex        =   216
         Top             =   2325
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   46
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   214
         Top             =   1275
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   46
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   213
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.CommandButton Command12 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":6FA7
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command11 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":72B1
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   66
         Left            =   660
         TabIndex        =   220
         Top             =   1800
         Width           =   450
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
         TabIndex        =   219
         Top             =   420
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Situación"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   58
         Left            =   660
         TabIndex        =   218
         Top             =   1290
         Width           =   660
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   46
         Left            =   1620
         MouseIcon       =   "frmListado.frx":75BB
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar situación"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   1620
         Picture         =   "frmListado.frx":770D
         ToolTipText     =   "Buscar fecha"
         Top             =   1800
         Width           =   240
      End
   End
   Begin VB.Frame FrameTrazabilidad 
      Height          =   4665
      Left            =   30
      TabIndex        =   203
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton CmdCancelTraza 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   206
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepTraza 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   205
         Top             =   3780
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   285
         Left            =   240
         TabIndex        =   204
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
         TabIndex        =   209
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Caption         =   "aa"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   208
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
         TabIndex        =   207
         Top             =   870
         Width           =   5820
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6960
      Top             =   5220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
    ' 12 .- Listado de Calidades
    ' 13 .- Listado de Socios por Sección
    ' 14 .- Listado de Entradas en Bascula
    ' 15 .- Listado de Campos
    ' 16 .- Listado de Entradas clasificacion
    ' 17 .- Reimpresion de albaranes de Clasificacion
    ' 18 .- Informe de Kilos/Gastos (rhisfruta)
    ' 19 .- Grabación de Fichero Agriweb
    ' 20 .- Informe de Kilos Por Producto
    ' 21 .- Traspaso desde el calibrador
    ' 22 .- Traspaso TRAZABILIDAD
    
    
    ' 23 .- Baja de Socios (dentro del mantenimiento socios)
    
    ' 24 .- Traspaso de Facturas Cooperativa ( traspaso liquidacion )
    ' 25 .- Listado de Kilos recolectados socio / cooperativa
    ' 26 .- Traspaso de ROPAS solo para Catadau
    
    
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
Private WithEvents frmSitu As frmManSituacion 'Situacion de socio
Attribute frmSitu.VB_VarHelpID = -1
Private WithEvents frmCoop As frmManCoope 'Cooperativa
Attribute frmCoop.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Tabla1 As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim indice As Integer

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

'[Monica] 01/10/2009 añadido el poder detallar las notas
Private Sub Check2_Click()
    If OpcionListado = 18 Then
        Check9.Enabled = (Check2.Value = 0)
        If Not Check9.Enabled Then Check9.Value = 0
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
Dim b As Boolean
Dim vSQL As String

    If Not DatosOk Then Exit Sub


    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
     'D/H Socio
     cDesde = Trim(txtcodigo(23).Text)
     cHasta = Trim(txtcodigo(24).Text)
     nDesde = txtNombre(23).Text
     nHasta = txtNombre(24).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
     End If

     'D/H Clase
     cDesde = Trim(txtcodigo(25).Text)
     cHasta = Trim(txtcodigo(26).Text)
     nDesde = txtNombre(25).Text
     nHasta = txtNombre(26).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codclase}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
     End If
     
    ' PRODUCTO
    If txtcodigo(22).Text <> "" Then
        If Not AnyadirAFormula(cadSelect, "{variedades.codprodu} = " & DBSet(txtcodigo(22).Text, "N")) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{variedades.codprodu}  = " & DBSet(txtcodigo(22).Text, "N")) Then Exit Sub
    End If
     
     If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
     If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub

     Tabla = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
     Tabla = "(" & Tabla & ") INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
     
     vSQL = ""
     If txtcodigo(25).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(25).Text, "N")
     If txtcodigo(26).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(26).Text, "N")
     If txtcodigo(22).Text <> "" Then vSQL = vSQL & " and variedades.codprodu = " & DBSet(txtcodigo(22).Text, "N")
     Set frmMens = New frmMensajes
     
     frmMens.OpcionMensaje = 16
     frmMens.cadWHERE = vSQL
     frmMens.Show vbModal
     
     Set frmMens = Nothing
     
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If HayRegParaInforme(Tabla, cadSelect) Then
        b = GeneraFicheroAgriweb(Tabla, cadSelect)
        If b Then
            If CopiarFichero Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                CmdCancelAgri_Click
            End If
        End If
     End If

End Sub

Private Sub cmdAcepBajaSocio_Click()
Dim Sql As String

    On Error GoTo eErrores

    If txtcodigo(47).Text = "" Then
        MsgBox "Debe introducir la fecha de baja.", vbExclamation
        PonerFoco txtcodigo(47)
        Exit Sub
    End If
    If txtcodigo(46).Text = "" Then
        MsgBox "Debe introducir la nueva situación del socio.", vbExclamation
        PonerFoco txtcodigo(46)
        Exit Sub
    End If
    
    Sql = "update rsocios_seccion set fecbaja = " & DBSet(txtcodigo(47), "F")
    Sql = Sql & " where codsocio = " & DBSet(NumCod, "N")
    Sql = Sql & " and fecbaja is null"
    conn.Execute Sql
    
    Sql = "update rsocios set codsitua = " & DBSet(txtcodigo(46).Text, "N")
    Sql = Sql & " where codsocio = " & DBSet(NumCod, "N")
    conn.Execute Sql
    
    MsgBox "Proceso realizado correctamente.", vbExclamation
    cmdCancelBajaSocio_Click
    Exit Sub
eErrores:
    MuestraError Err.Number, "Baja de Socio", Err.Description
End Sub

Private Sub CmdAcepInf_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean
Dim vSQL As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
     'D/H Socio
     cDesde = Trim(txtcodigo(33).Text)
     cHasta = Trim(txtcodigo(34).Text)
     nDesde = txtNombre(33).Text
     nHasta = txtNombre(34).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rcampos.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
     End If

     'D/H Clase
     cDesde = Trim(txtcodigo(35).Text)
     cHasta = Trim(txtcodigo(36).Text)
     nDesde = txtNombre(35).Text
     nHasta = txtNombre(36).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codclase}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
     End If
        
     
    ' PRODUCTO
     cDesde = Trim(txtcodigo(37).Text)
     cHasta = Trim(txtcodigo(38).Text)
     nDesde = txtNombre(37).Text
     nHasta = txtNombre(38).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codprodu}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
     End If
        
     
     'D/H fecha
     cDesde = Trim(txtcodigo(39).Text)
     cHasta = Trim(txtcodigo(40).Text)
     nDesde = ""
     nHasta = ""
     devuelve = CadenaDesdeHasta(cDesde, cHasta, "fecalbar", "F")

'     If Not (cDesde = "" And cHasta = "") Then
'         'Cadena para seleccion Desde y Hasta
'         codigo = "{rhisfruta.fecalbar}"
'         TipCod = "F"
'         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
'     End If
     
'     If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
'     If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub

     Tabla = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
     Tabla = "(" & Tabla & ") INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
     
     
     vSQL = ""
     If txtcodigo(35).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(35).Text, "N")
     If txtcodigo(36).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(36).Text, "N")
     
     If txtcodigo(37).Text <> "" Then vSQL = vSQL & " and variedades.codprodu >= " & DBSet(txtcodigo(37).Text, "N")
     If txtcodigo(38).Text <> "" Then vSQL = vSQL & " and variedades.codprodu <= " & DBSet(txtcodigo(38).Text, "N")
     
     Set frmMens = New frmMensajes
     
     frmMens.OpcionMensaje = 16
     frmMens.cadWHERE = vSQL
     frmMens.Show vbModal
     
     Set frmMens = Nothing
            
     'combo1(5): tipo de has
     cadParam = cadParam & "pTipoHas=" & Combo1(5).ListIndex & "|"
     numParam = numParam + 1
     
     ' salto de pagina o no por producto
     cadParam = cadParam & "pSaltoProd=" & Check3.Value & "|"
     numParam = numParam + 1
     
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If HayRegParaInforme(Tabla, cadSelect) Then
        If CargarTemporal5(Tabla, cadSelect) Then
           If HayRegParaInforme("tmpinfkilos", "codusu = " & vUsu.Codigo) Then
               cadNombreRPT = "rInfKilosProd.rpt"
               cadTitulo = "Informe de Kilos por Producto"
               
               cadFormula = "{tmpinfkilos.codusu} = " & vUsu.Codigo
               
               LlamarImprimir
           End If
        End If
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
Dim b As Boolean
Dim vSQL As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
     'D/H Socio
     cDesde = Trim(txtcodigo(54).Text)
     cHasta = Trim(txtcodigo(55).Text)
     nDesde = txtNombre(54).Text
     nHasta = txtNombre(55).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rhisfruta.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
     End If

     'D/H Clase
     cDesde = Trim(txtcodigo(56).Text)
     cHasta = Trim(txtcodigo(57).Text)
     nDesde = txtNombre(56).Text
     nHasta = txtNombre(57).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codclase}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
     End If
        
    vSQL = ""
    If txtcodigo(56).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(56).Text, "N")
    If txtcodigo(57).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(57).Text, "N")
     
    ' PRODUCTO
     cDesde = Trim(txtcodigo(50).Text)
     cHasta = Trim(txtcodigo(51).Text)
     nDesde = txtNombre(50).Text
     nHasta = txtNombre(51).Text
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{variedades.codprodu}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
     End If
        
    If txtcodigo(50).Text <> "" Then vSQL = vSQL & " and variedades.codprodu >= " & DBSet(txtcodigo(50).Text, "N")
    If txtcodigo(51).Text <> "" Then vSQL = vSQL & " and variedades.codprodu <= " & DBSet(txtcodigo(51).Text, "N")
     
     'D/H fecha
     cDesde = Trim(txtcodigo(52).Text)
     cHasta = Trim(txtcodigo(53).Text)
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
            
     Tabla = "rcampos"
            
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If CargarTemporal6(Tabla, cadSelect) Then
         If HayRegParaInforme("tmpclasifica", "codusu = " & vUsu.Codigo) Then
             cadNombreRPT = "rInfKilosSocio.rpt"
             cadTitulo = "Informe de Kilos Socio/Cooperativa"
             
             cadFormula = "{tmpclasifica.codusu} = " & vUsu.Codigo
             
             LlamarImprimir
         End If
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
Dim b As Boolean
Dim vSQL As String


    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


    If txtcodigo(62).Text = "" Then
        MsgBox "Debe introducir un valor en el campo ejercicio. Revise.", vbExclamation
        PonerFoco txtcodigo(62)
        Exit Sub
    End If

     '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtcodigo(58).Text)
    cHasta = Trim(txtcodigo(59).Text)
    nDesde = txtNombre(58).Text
    nHasta = txtNombre(59).Text
    If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rsocios.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If

    cadSelect1 = cadSelect


    'D/H Producto
    cDesde = Trim(txtcodigo(60).Text)
    cHasta = Trim(txtcodigo(61).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codprodu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
    End If
    
    vSQL = ""
    If txtcodigo(60).Text <> "" Then vSQL = vSQL & " and variedades.codprodu >= " & DBSet(txtcodigo(60).Text, "N")
    If txtcodigo(61).Text <> "" Then vSQL = vSQL & " and variedades.codprodu <= " & DBSet(txtcodigo(61).Text, "N")
    
    Set frmMens1 = New frmMensajes
    
    frmMens1.OpcionMensaje = 4
    frmMens1.cadWHERE = vSQL
    frmMens1.Show vbModal
    
    Set frmMens1 = Nothing

    Tabla1 = "rsocios INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & vParamAplic.SeccionHorto
    Tabla1 = Tabla1 & " and rsocios_seccion.fecbaja is null "
    
    Tabla = "((" & Tabla1 & ") INNER JOIN rcampos ON rcampos.codsocio = rsocios.codsocio and rcampos.fecbajas is null and rcampos.supcoope <> 0) "
    Tabla = Tabla & " INNER JOIN variedades on rcampos.codvarie = variedades.codvarie "
     
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(Tabla1, cadSelect1) Then
        b = GeneraFicheroTraspasoROPAS(Tabla1, cadSelect1, Tabla, cadSelect)
        If b Then
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
    
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Select Case Index
       Case 0 'frame de campos
            '======== FORMULA  ====================================
            'D/H Socio
            cDesde = Trim(txtcodigo(2).Text)
            cHasta = Trim(txtcodigo(3).Text)
            nDesde = txtNombre(2).Text
            nHasta = txtNombre(3).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rcampos.codsocio}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
            End If
       
            'D/H Clase
            cDesde = Trim(txtcodigo(0).Text)
            cHasta = Trim(txtcodigo(1).Text)
            nDesde = txtNombre(0).Text
            nHasta = txtNombre(1).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{variedades.codclase}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
            End If
    
            vSQL = ""
            If txtcodigo(0).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(0).Text, "N")
            If txtcodigo(1).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(1).Text, "N")
            
    
    
            'D/H Situacion
            cDesde = Trim(txtcodigo(4).Text)
            cHasta = Trim(txtcodigo(5).Text)
            nDesde = txtNombre(4).Text
            nHasta = txtNombre(5).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rcampos.codsitua}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSituacion= """) Then Exit Sub
            End If
    
            Tabla = "(rcampos INNER JOIN rpartida ON rcampos.codparti = rpartida.codparti) "
            Tabla = Tabla & " INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie "
            
            If Opcion1(0).Value Then numOp = PonerGrupo(1, "Socios")
            If Opcion1(1).Value Then numOp = PonerGrupo(1, "Clases")
            If Opcion1(2).Value Then numOp = PonerGrupo(1, "Terminos")
            If Opcion1(3).Value Then numOp = PonerGrupo(1, "Zonas")
            
            If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub
            
            
            cadTitulo = "Informe de Campos"
            If Opcion1(0).Value Then cadTitulo = cadTitulo & " por Socios"
            If Opcion1(1).Value Then cadTitulo = cadTitulo & " por Clases"
            If Opcion1(2).Value Then cadTitulo = cadTitulo & " por Terminos"
            If Opcion1(3).Value Then cadTitulo = cadTitulo & " por Zonas"
            
            
            'combo1(0): tipo de has
            cadParam = cadParam & "pTipoHas=" & Combo1(0).ListIndex & "|"
            numParam = numParam + 1
            
            'combo1(1): tipo de kilos 0=aforo 1=real
            cadParam = cadParam & "pKilos=" & Combo1(1).ListIndex & "|"
            numParam = numParam + 1
            
            cadNombreRPT = "rInfCampos.rpt"
            
            ' resumen o no
            cadParam = cadParam & "pResumen=" & Format(Check1.Value, "0") & "|"
            numParam = numParam + 1
            
            ' Imprimir Cabecera o no
            cadParam = cadParam & "pCabecera=" & Format(Check4.Value, "0") & "|"
            numParam = numParam + 1
            
            Set frmMens = New frmMensajes
            
            frmMens.OpcionMensaje = 16
            frmMens.cadWHERE = vSQL
            frmMens.Show vbModal
            
            Set frmMens = Nothing
            
             'Comprobar si hay registros a Mostrar antes de abrir el Informe
            If HayRegParaInforme(Tabla, cadSelect) Then
                If CargarTemporal(Tabla, cadSelect) Then
'                   cadParam = cadParam & "pCampos=""" & ConcatenarCampos(tabla, cadselect) & """|"
'                   numParam = numParam + 1
                    cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
                    numParam = numParam + 1

                    With frmImprimir
                        .FormulaSeleccion = cadFormula
                        .OtrosParametros = cadParam
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
      
       Case 1 'Frame Informe de socios por seccion
            '======== FORMULA  ====================================
            'D/H Seccion
            cDesde = Trim(txtcodigo(8).Text)
            cHasta = Trim(txtcodigo(9).Text)
            nDesde = txtNombre(8).Text
            nHasta = txtNombre(9).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & Tabla & ".codsecci}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSeccion= """) Then Exit Sub
            End If
            
            'D/H Socio
            cDesde = Trim(txtcodigo(10).Text)
            cHasta = Trim(txtcodigo(11).Text)
            nDesde = txtNombre(10).Text
            nHasta = txtNombre(11).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & Tabla & ".codsocio}"
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
            
            
            
            'Obtener el parametro con el ORDEN del Informe
            '---------------------------------------------
        '    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
        '    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
            If Opcion(0).Value Then numOp = PonerGrupo(1, "Seccion")
            If Opcion(1).Value Then numOp = PonerGrupo(1, "Socio")
            
            cadNombreRPT = "rManSocSeccion.rpt"
            
            If Opcion(0).Value Then cadTitulo = "Listado de Socios por Sección"
            If Opcion(1).Value Then cadTitulo = "Listado de Socios"
            
            'Comprobar si hay registros a Mostrar antes de abrir el Informe
            If HayRegParaInforme(Tabla, cadSelect) Then
                LlamarImprimir
            End If
        
        Case 2 ' informe de calidades
            '======== FORMULA  ====================================
            'D/H Variedad
            cDesde = Trim(txtcodigo(18).Text)
            cHasta = Trim(txtcodigo(19).Text)
            nDesde = txtNombre(18).Text
            nHasta = txtNombre(19).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & Tabla & ".codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
            End If
            
            'D/H Calidad
            cDesde = Trim(txtcodigo(16).Text)
            cHasta = Trim(txtcodigo(17).Text)
            nDesde = txtNombre(16).Text
            nHasta = txtNombre(17).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & Tabla & ".codcalid}"
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
            If HayRegParaInforme(Tabla, cadSelect) Then
                LlamarImprimir
            End If
            
            
        Case 3 ' informe de entradas de bascula
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
            
            'D/H VARIEDAD
            cDesde = Trim(txtcodigo(14).Text)
            cHasta = Trim(txtcodigo(15).Text)
            nDesde = txtNombre(14).Text
            nHasta = txtNombre(15).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & Tabla & ".codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
            End If

            'D/H fecha
            cDesde = Trim(txtcodigo(6).Text)
            cHasta = Trim(txtcodigo(7).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                Select Case OpcionListado
                    Case 14, 16
                        'Cadena para seleccion Desde y Hasta
                        Codigo = "{" & Tabla & ".fechaent}"
                    Case 17, 18
                        'Cadena para seleccion Desde y Hasta
                        Codigo = "{" & Tabla & ".fecalbar}"
                End Select
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
                
            Select Case OpcionListado
              Case 14 ' listado de entradas (rentradas)
                ' resumen o no
                cadParam = cadParam & "pResumen=" & Format(Check2.Value, "0") & "|"
                numParam = numParam + 1
                
                nTabla = "(rentradas INNER JOIN variedades ON rentradas.codvarie = variedades.codvarie) "
    
                cadNombreRPT = "rInfEntradas.rpt"
                cadTitulo = "Informe de Entradas Báscula"
                
                ConSubInforme = True
            
            
                'Comprobar si hay registros a Mostrar antes de abrir el Informe
                If HayRegParaInforme(nTabla, cadSelect) Then
                    LlamarImprimir
                End If
            
              Case 16 ' listado de entradas clasificadas (rclasifica)
                nTabla = "(rclasifica INNER JOIN variedades ON rclasifica.codvarie = variedades.codvarie) "
                Select Case Combo1(3).ListIndex
                    Case 0 ' informe normal
                        If CargarTemporal2(nTabla, cadSelect) Then
                            cadNombreRPT = "rInfEntradasClas.rpt"
                            cadTitulo = "Informe de Entradas Clasificadas"
                            
                            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                            'Comprobar si hay registros a Mostrar antes de abrir el Informe
                            If HayRegParaInforme("tmpinformes", "{tmpinformes.codusu} = " & vUsu.Codigo) Then
                                LlamarImprimir
                            End If
                        End If
                    Case 1 ' informe detalle clasificacion
                        If CargarTemporal3(nTabla, cadSelect) Then
                            cadNombreRPT = "rInfEntradasClas1.rpt"
                            cadTitulo = "Informe de Entradas Clasificadas"
                            
                            cadFormula = "{tmpclasifica.codusu} = " & vUsu.Codigo
                            'Comprobar si hay registros a Mostrar antes de abrir el Informe
                            If HayRegParaInforme("tmpclasifica", "{tmpclasifica.codusu} = " & vUsu.Codigo) Then
                                With frmImprimir
                                    .FormulaSeleccion = cadFormula
                                    .OtrosParametros = cadParam
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
                cadParam = cadParam & "pDuplicado=0|"
                numParam = numParam + 1
                
                indRPT = 22 'Impresion de Albaran de clasificacion
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                  
                'Nombre fichero .rpt a Imprimir
'                frmImprimir.NombreRPT = nomDocu
                cadNombreRPT = nomDocu
'                cadNombreRPT = "rInfEntradas.rpt"
                cadTitulo = "Impresion de Albaranes"
'                OpcionListado = 22
                ConSubInforme = True
                
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.impreso} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.impreso} = 0") Then Exit Sub
            
                'Comprobar si hay registros a Mostrar antes de abrir el Informe
                If HayRegParaInforme(Tabla, cadSelect) Then
                    LlamarImprimir
                    
                    If MsgBox("¿ Impresión correcta para actualizar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        If ActualizarRegistros(Tabla, cadSelect) Then
                            MsgBox "Proceso realizado correctamente.", vbExclamation
                            cmdCancel_Click (0)
                        End If
                    End If
                End If
              
              Case 18 ' informe de kilos/gastos
                nTabla = "(rhisfruta INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
                If CargarTemporal4New(nTabla, cadSelect) Then
                    
                    If Check5.Value = 1 Then
                        ' imprimimos por socio
                        cadParam = cadParam & "pSaltar=" & Check6.Value & "|"
                        numParam = numParam + 1
                       '[Monica] 01/10/2009 añadido el poder detallar las notas
                        cadParam = cadParam & "pDetalleNota=" & Check9.Value & "|"
                        numParam = numParam + 1
                        
                        cadNombreRPT = "rInfHcoEntClas2.rpt"
                    Else
                        If Check2.Value = 0 Then
                            cadNombreRPT = "rInfHcoEntClas.rpt"
                            '[Monica] 01/10/2009 añadido el poder detallar las notas
                             cadParam = cadParam & "pDetalleNota=" & Check9.Value & "|"
                             numParam = numParam + 1
                        Else
                            ' imprimimos un resumen por variedad
                            cadNombreRPT = "rInfHcoEntClas1.rpt"
                        End If
                    End If
                    cadTitulo = "Informe de Kilos / Gastos"
                    
                    cadFormula = "{tmpclasifica2.codusu} = " & vUsu.Codigo
                    'Comprobar si hay registros a Mostrar antes de abrir el Informe
                    If HayRegParaInforme("tmpclasifica2", "{tmpclasifica2.codusu} = " & vUsu.Codigo) Then
                        With frmImprimir
                            .FormulaSeleccion = cadFormula
                            .OtrosParametros = cadParam
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
    End Select
End Sub

Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView1
End Sub

Private Sub cmdAcepTras_Click()
Dim Sql As String
Dim I As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim NomFic As String
Dim CADENA As String
Dim cadena1 As String
Dim Directorio As String

Dim File1 As FileSystemObject

On Error GoTo eError

    
    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist


    If Combo1(6).ListIndex = 2 And vParamAplic.Cooperativa = 4 Then
    ' solo para el calibrador de alzira de kaki la extension es diferente
        Me.CommonDialog1.DefaultExt = "pdt"
        CommonDialog1.Filter = "Archivos PTD|*.ptd|"
        CommonDialog1.FilterIndex = 1
        Me.CommonDialog1.FileName = "*.ptd"
    Else
        Me.CommonDialog1.DefaultExt = "txt"
        CommonDialog1.Filter = "Archivos TXT|*.txt|"
        CommonDialog1.FilterIndex = 1
        Me.CommonDialog1.FileName = "*.txt"
    End If
    
    
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowOpen
    Set File1 = New FileSystemObject
    
    Directorio = File1.GetParentFolderName(Me.CommonDialog1.FileName)

    Select Case vParamAplic.Cooperativa
        Case 0  '******* CATADAU *******
'             Directorio = GetFolder("Selecciona directorio")
            If Directorio <> "" Then
                Sql = "DROP TABLE IF EXISTS tmpCata; "
                conn.Execute Sql
                
                Sql = "CREATE TEMPORARY TABLE tmpCata ("
                Sql = Sql & " codcalid int, kilosnet decimal(10,2)) "
                conn.Execute Sql
                
                If Combo1(6).ListIndex = 1 Then ' si calibrador pequeño
                    'creamos la tabla temporal solo si estamos en calibrador pequeño
                    Sql = "DROP TABLE IF EXISTS tmpcalibrador; "
                    conn.Execute Sql
                    
                    Sql = "CREATE TEMPORARY TABLE `tmpcalibrador` ("
                    Sql = Sql & "`numnota` varchar(10) default NULL, "
                    Sql = Sql & "`fecnota` varchar(20) default NULL, "
                    Sql = Sql & "`albaran` varchar(20) default NULL, "
                    Sql = Sql & "`porcen1` varchar(10) default NULL, "
                    Sql = Sql & "`porcen2` varchar(10) default NULL, "
                    Sql = Sql & "`kilos1` varchar(30) default NULL, "
                    Sql = Sql & "`kilos2` varchar(30) default NULL, "
                    Sql = Sql & "`kilos3` varchar(30) default NULL, "
                    Sql = Sql & "`numnota2` varchar(10) default NULL, "
                    Sql = Sql & "`export` varchar(10) default NULL, "
                    Sql = Sql & "`nomcalid` varchar(30) default NULL, "
                    Sql = Sql & "`kilos4` varchar(30) default NULL, "
                    Sql = Sql & "`kilos5` varchar(30) default NULL "
                    Sql = Sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
                
                    conn.Execute Sql
                End If
                
                conn.BeginTrans

                b = ProcesarDirectorioCatadau(Directorio & "\", Combo1(6).ListIndex, Pb1, lblProgres(0), lblProgres(1))
            End If
        
        Case 1 '********* VALSUR *************
            CommonDialog1.FilterIndex = 1
            Me.CommonDialog1.ShowOpen
            
            If Me.CommonDialog1.FileName <> "" Then
                InicializarVbles
        '        InicializarTabla
                    '========= PARAMETROS  =============================
                'Añadir el parametro de Empresa
                cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
                numParam = numParam + 1
        
                If vParamAplic.Cooperativa = 0 Then
                    Sql = "DROP TABLE IF EXISTS tmpCata; "
                    conn.Execute Sql
                    
                    
                    Sql = "CREATE TEMPORARY TABLE tmpCata ("
                    Sql = Sql & " codcalid int, kilosnet decimal(10,2)) "
                    conn.Execute Sql
                End If
    
                conn.BeginTrans
                ' resto de casos (incluido catadau, calibrador grande)
                b = ProcesarFichero(Me.CommonDialog1.FileName, Combo1(6).ListIndex, Me.Pb1, Me.lblProgres(0), Me.lblProgres(1))
            Else
                MsgBox "No ha seleccionado ningún fichero", vbExclamation
                Exit Sub
            End If
    
        Case 4 ' ******** ALZIRA **********
            If Directorio <> "" Then

                Sql = "DROP TABLE IF EXISTS tmpCata; "
                conn.Execute Sql
                
                Sql = "CREATE TEMPORARY TABLE tmpCata ("
                Sql = Sql & " codcalid int, kilosnet decimal(10,2)) "
                conn.Execute Sql
                
                
                If Combo1(6).ListIndex = 0 Then
                    'creamos la tabla temporal solo si estamos en precalibrado
                    Sql = "DROP TABLE IF EXISTS tmpcalibrador; "
                    conn.Execute Sql
                    
                    Sql = "CREATE TEMPORARY TABLE `tmpcalibrador` ("
                    Sql = Sql & "`numnota` varchar(10) default NULL, "
                    Sql = Sql & "`fecnota` varchar(20) default NULL, "
                    Sql = Sql & "`nomcalid` varchar(30) default NULL, "
                    Sql = Sql & "`kilos1` varchar(30) default NULL, "
                    Sql = Sql & "`kilos2` varchar(30) default NULL, "
                    Sql = Sql & "`kilos3` varchar(30) default NULL, "
                    Sql = Sql & "`kilos4` varchar(30) default NULL "
                    Sql = Sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
                
                    conn.Execute Sql
                End If
            
                conn.BeginTrans

                b = ProcesarDirectorioAlzira(Directorio & "\", Combo1(6).ListIndex, Pb1, lblProgres(0), lblProgres(1))
            End If
    End Select
    
eError:
    If Err.Number = 32755 Then Exit Sub ' le han dado a cancelar

    If Err.Number <> 0 Or Not b Then
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

Private Sub cmdAcepTrasCoop_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean
Dim vSQL As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


     '======== FORMULA  ====================================
    ' Cooperativa
    If Not AnyadirAFormula(cadSelect, "{rsocios.codcoope} = " & DBSet(txtcodigo(45).Text, "N")) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "{rsocios.codcoope} = " & DBSet(txtcodigo(45).Text, "N")) Then Exit Sub
     
    'D/H Socio
    cDesde = Trim(txtcodigo(48).Text)
    cHasta = Trim(txtcodigo(49).Text)
    nDesde = txtNombre(48).Text
    nHasta = txtNombre(49).Text
    If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{rfactsoc.codsocio}"
         TipCod = "N"
         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If

    'D/H Fecha de Factura
    cDesde = Trim(txtcodigo(43).Text)
    cHasta = Trim(txtcodigo(44).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= """) Then Exit Sub
    End If

    'D/H Factura
    cDesde = Trim(txtcodigo(41).Text)
    cHasta = Trim(txtcodigo(42).Text)
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
     
    Tabla = "rfactsoc INNER JOIN rsocios ON rfactsoc.codsocio = rsocios.codsocio"
     
      'Comprobar si hay registros a Mostrar antes de abrir el Informe
     If HayRegParaInforme(Tabla, cadSelect) Then
        b = GeneraFicheroTraspasoCoop(Tabla, cadSelect)
        If b Then
            If CopiarFicheroCoop(txtcodigo(45).Text) Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                cmdCancelTrasCoop_Click
            End If
        End If
     End If


End Sub

Private Sub CmdAcepTraza_Click()
Dim Sql As String
Dim I As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim NomFic As String
Dim CADENA As String
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
    Sql = "DROP TABLE IF EXISTS tmpEntrada; "
    conn.Execute Sql
    
    Sql = "DROP TABLE IF EXISTS tmpClasific; "
    conn.Execute Sql
    
    
    Sql = "CREATE TEMPORARY TABLE tmpEntrada ("
    Sql = Sql & " codsocio int, codcampo int, numalbar int, codvarie int, fecalbar date, "
    Sql = Sql & " horalbar datetime, kilosbru int, kilosnet int, numcajon int) "
    conn.Execute Sql
    
    Sql = "CREATE TEMPORARY TABLE tmpClasific ("
    Sql = Sql & " numalbar int, codvarie int, codcalir int, porcenta decimal(5,2)) "
    conn.Execute Sql
'08052009
        
        
        conn.BeginTrans
    
        If CargarTablasTemporales(Fichero1, Fichero2) Then
            InicializarVbles
                
                '========= PARAMETROS  =============================
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
    
            If ComprobarErrores() Then
                    cadTabla = "tmpinformes"
                    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                    
                    Sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
                    
                    If TotalRegistros(Sql) <> 0 Then
                        MsgBox "Hay errores en el Traspaso de Trazabilidad. Debe corregirlos previamente.", vbExclamation
                        cadTitulo = "Errores de Traspaso de TRAZABILIDAD"
                        cadNombreRPT = "rErroresTraza.rpt"
                        LlamarImprimir
                        conn.RollbackTrans
                        Exit Sub
                    Else
                        b = CargarClasificacion()
                    End If
            Else
                b = False
            End If
        Else
            b = False
        End If
    Else
        CmdCancelTraza_Click
        Exit Sub
    End If
    
eError:
    If Err.Number <> 0 Or Not b Then
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

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim b As Boolean
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
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 12 ' Listado de Calidades
                PonerFoco txtcodigo(18)
        
            Case 13 ' Listado de Socios por seccion
                PonerFoco txtcodigo(8)
                
            Case 14, 16, 17, 18 '14 = Listado de entradas en bascula
                                '16 = Listado de Entradas clasificadas
                                '17 = Reimpresion de Albaranes de Clasificacion
                                '18 = Informe de Kilos/gastos
                PonerFoco txtcodigo(12)
                
            Case 15 ' Listado de campos huertos
                PonerFoco txtcodigo(2)
                
            Case 19 ' grabacion de fichero agriweb
                PonerFoco txtcodigo(23)
                txtcodigo(27).Text = Format(Year(Now), "0000")
                
            Case 20 ' informe de kilos por producto
                PonerFoco txtcodigo(33)
                
            Case 21 ' traspaso desde el calibrador
                Combo1(6).SetFocus
                
            Case 23 ' baja de socio
                PonerFoco txtcodigo(46)
            
            Case 24 ' traspaso factura cooperativa
                Combo1(7).ListIndex = 0
                PonerFoco txtcodigo(45)
                
            Case 25 ' informe de kilos recolectados por socio/cooperativa
                PonerFoco txtcodigo(54)
            
            Case 26 ' traspaso de ROPAS
                PonerFoco txtcodigo(58)
                txtcodigo(62).Text = Format(Year(Now), "0000")
            
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me
    
    ConSubInforme = False

    For H = 0 To 20
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 22 To 28
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 29 To 32
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 33 To 38
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 46 To 46
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 40 To 45
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    
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
    
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 12 ' Listado de Calidades
        FrameCalidadesVisible True, H, W
        CargarListViewOrden (2)
        indFrame = 2
        Tabla = "rcalidad"
    
    Case 13 ' Listado de Socios por Seccion
        FrameSociosSeccionVisible True, H, W
        Opcion(0).Value = True
        CargarListViewOrden (3)
        indFrame = 1
        Tabla = "rsocios_seccion"
        
    Case 14 ' Listado de entradas en bascula
        FrameEntradaBasculaVisible True, H, W
        Opcion(0).Value = True
        indFrame = 1
        Tabla = "rentradas"
        Check2.visible = True
        Check2.Enabled = True
        Check5.visible = False
        Check5.Enabled = False
        Check6.visible = False
        Check6.Enabled = False
        '[Monica] 01/10/2009 añadido el poder detallar las notas
        Check9.visible = False
        Check9.Enabled = False
        
        
        FrameTipo.Enabled = False
        FrameTipo.visible = False
        
    Case 15 ' Listado de Campos
        CargaCombo
        Combo1(0).ListIndex = 0
        Combo1(1).ListIndex = 0
        FrameCamposVisible True, H, W
        Opcion1(0).Value = True
        Tabla = "rcampos"
        
    Case 16, 17, 18 '16= Listado de entradas clasificacion
                    '17= reimpresion de albaranes de clasificacion
                    '18= informe de kilos/gastos
        CargaCombo
        Combo1(2).ListIndex = 0
        Combo1(3).ListIndex = 0
        FrameEntradaBasculaVisible True, H, W
        Opcion(0).Value = True
        indFrame = 1
        Select Case OpcionListado
            Case 16
                Tabla = "rclasifica"
                Check2.visible = False
                Check2.Enabled = False
                Check5.visible = False
                Check5.Enabled = False
                Check6.visible = False
                Check6.Enabled = False
               '[Monica] 01/10/2009 añadido el poder detallar las notas
                Check9.visible = False
                Check9.Enabled = False
                FrameTipo.Enabled = True
                FrameTipo.visible = True
                Label3.Caption = "Informe de Entradas"
            Case 17, 18
                Tabla = "rhisfruta"
                FrameTipo.Enabled = False
                FrameTipo.visible = False
                If OpcionListado = 17 Then
                    Check2.visible = False
                    Check2.Enabled = False
                    Check5.visible = False
                    Check5.Enabled = False
                    Check6.visible = False
                    Check6.Enabled = False
                    '[Monica] 01/10/2009 añadido el poder detallar las notas
                    Check9.visible = False
                    Check9.Enabled = False

                    Label3.Caption = "Reimpresión de Albaranes"
                Else
                    Check2.visible = True
                    Check2.Enabled = True
                    Check5.visible = True
                    Check5.Enabled = True
                    Check6.visible = True
                    Check6.Enabled = True
                    '[Monica] 01/10/2009 añadido el poder detallar las notas
                    Check9.visible = True
                    Check9.Enabled = True
                    Label3.Caption = "Informe de Kilos/Gastos"
                End If
        End Select
    
    Case 19 ' grabacion de fichero agriweb
        CargaCombo
        Combo1(4).ListIndex = 0
        FrameGrabacionFicheroVisible True, H, W
    
    Case 20 ' informe de kilos por producto
        CargaCombo
        Combo1(5).ListIndex = 0
        FrameKilosProductoVisible True, H, W
        
    Case 21 ' traspaso desde el calibrador
        CargaCombo
        Combo1(6).ListIndex = 0
        FrameTraspasoCalibradorVisible True, H, W
        Pb1.visible = False
        
    Case 22 ' traspaso de trazabilidad
        FrameTraspasoTrazaVisible True, H, W
        pb2.visible = False
        lblProgres(2).Caption = ""
        lblProgres(3).Caption = ""
        
    Case 23 ' baja de socios
        FrameBajaSociosVisible True, H, W
    
    Case 24 ' traspaso facturas cooperativa (VALSUR)
        CargaCombo
        FrameTraspasoFactCoopVisible True, H, W
    
    Case 25 ' informe de kilos recolectados por socio cooperativa
        FrameKilosRecolectVisible True, H, W
    
    
    Case 26 ' traspaso ROPAS
        FrameTraspasoROPASVisible True, H, W
    
    End Select
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub




Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
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

Private Sub frmCoop_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") ' codigo de cooperativa
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion de la cooperativa
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
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMens1_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
        
        If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    Else
        Sql = " {rsocios.codsocio} = -1 "
        
        If Not AnyadirAFormula(cadSelect1, Sql) Then Exit Sub
    End If
End Sub

Private Sub frmProd_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSitu_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
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


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1, 6, 7, 25, 26, 35, 36  ' Clase
            AbrirFrmClase (Index)
        
        Case 31, 32 'clase
            AbrirFrmClase (Index + 25)
        
        Case 4, 5 ' Situacion de campo
            AbrirFrmSituacionCampo (Index)
            
        Case 8, 9 'SECCION
            AbrirFrmSeccion (Index)
        
        Case 2, 3, 10, 11, 12, 13, 23, 24, 33, 34  'SOCIOS
            AbrirFrmSocios (Index)
        
        Case 29, 30 'socios
            AbrirFrmSocios (Index + 19)
        
        Case 27, 28 'socios
            AbrirFrmSocios (Index + 27)
        
        Case 42, 43 'socios
            AbrirFrmSocios (Index + 16)
        
        Case 14, 15, 18, 19 'VARIEDADES
            AbrirFrmVariedad (Index)
    
        Case 20 ' cooperativa
            AbrirFrmCooperativa (45)
            
        Case 16, 17 'CALIDADES
            AbrirFrmCalidad (Index)
            
        Case 22, 37, 38 'Producto
            AbrirFrmProducto (Index)
            
        Case 40, 41 'Producto
            AbrirFrmProducto (Index + 10)
            
        Case 44, 45 'Producto
            AbrirFrmProducto (Index + 16)
        
        Case 46 ' situacion de socio
            AbrirFrmSituacion (Index)
        
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
        Case 0, 1
            indice = Index + 6
        Case 11
            indice = 30
        Case 2, 3
            indice = Index + 37
        Case 5
            indice = 47
        Case 7, 8
            indice = Index + 45
    End Select


    imgFec(0).Tag = indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(indice).Text <> "" Then frmC.NovaData = txtcodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(indice) '<===
    ' ********************************************

End Sub


Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1, 20, 21, 25, 26, 35, 36, 56, 57 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
    
        Case 8, 9 'SECCIONES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rseccion", "nomsecci", "codsecci", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            
        Case 2, 3, 10, 11, 12, 13, 23, 24, 33, 34, 48, 49, 54, 55, 58, 59 'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            
        Case 4, 5 'SITUACION
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsituacioncampo", "nomsitua", "codsitua", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "00")
            
        Case 6, 7, 30, 39, 40, 47, 43, 44, 52, 53 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
        Case 16, 17 'CALIDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rcalidad", "nomcalid", "codcalid", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "00")
            
        Case 14, 15, 18, 19 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
        
        Case 22, 37, 38, 50, 51, 60, 61 'PRODUCTOS
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "productos", "nomprodu", "codprodu", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
            
        Case 27, 29 ' datos de agroweb
            txtcodigo(Index).Text = Format(txtcodigo(Index).Text, FormatoCampo(txtcodigo(Index)))
            
        Case 31
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index) = Format(TransformaPuntosComas(txtcodigo(Index).Text), "#,##0.00")
            
        Case 32 ' datos de agroweb
            PonerFormatoDecimal txtcodigo(Index), 4
    
        Case 45 ' cooperativa
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rcoope", "nomcoope", "codcoope", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "00")
        
        Case 46 'SITUACION de socio
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsituacion", "nomsitua", "codsitua", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "00")
    
        Case 62 ' Ejercicio
            txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000")
    
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

Private Sub FrameCalidadesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameCalidades.visible = visible
    If visible = True Then
        Me.FrameCalidades.Top = -90
        Me.FrameCalidades.Left = 0
        Me.FrameCalidades.Height = 4820
        Me.FrameCalidades.Width = 8600
        W = Me.FrameCalidades.Width
        H = Me.FrameCalidades.Height
    End If
End Sub

Private Sub FrameSociosSeccionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameSociosSeccion.visible = visible
    If visible = True Then
        Me.FrameSociosSeccion.Top = -90
        Me.FrameSociosSeccion.Left = 0
        Me.FrameSociosSeccion.Height = 4820
        Me.FrameSociosSeccion.Width = 6600
        W = Me.FrameSociosSeccion.Width
        H = Me.FrameSociosSeccion.Height
    End If
End Sub

Private Sub FrameEntradaBasculaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameEntradasCampo.visible = visible
    If visible = True Then
        Me.FrameEntradasCampo.Top = -90
        Me.FrameEntradasCampo.Left = 0
        Me.FrameEntradasCampo.Height = 6480
        Me.FrameEntradasCampo.Width = 6615
        W = Me.FrameEntradasCampo.Width
        H = Me.FrameEntradasCampo.Height
    End If
End Sub

Private Sub FrameCamposVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameCampos.visible = visible
    If visible = True Then
        Me.FrameCampos.Top = -90
        Me.FrameCampos.Left = 0
        Me.FrameCampos.Height = 6390
        Me.FrameCampos.Width = 6600
        W = Me.FrameCampos.Width
        H = Me.FrameCampos.Height
    End If
End Sub

Private Sub FrameGrabacionFicheroVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameGrabacionAgriweb.visible = visible
    If visible = True Then
        Me.FrameGrabacionAgriweb.Top = -90
        Me.FrameGrabacionAgriweb.Left = 0
        Me.FrameGrabacionAgriweb.Height = 6975
        Me.FrameGrabacionAgriweb.Width = 6675
        W = Me.FrameGrabacionAgriweb.Width
        H = Me.FrameGrabacionAgriweb.Height
    End If
End Sub

Private Sub FrameKilosProductoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameKilosProducto.visible = visible
    If visible = True Then
        Me.FrameKilosProducto.Top = -90
        Me.FrameKilosProducto.Left = 0
        Me.FrameKilosProducto.Height = 6480
        Me.FrameKilosProducto.Width = 6615
        W = Me.FrameKilosProducto.Width
        H = Me.FrameKilosProducto.Height
    End If
End Sub

Private Sub FrameTraspasoCalibradorVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameTraspasoCalibrador.visible = visible
    If visible = True Then
        Me.FrameTraspasoCalibrador.Top = -90
        Me.FrameTraspasoCalibrador.Left = 0
        Me.FrameTraspasoCalibrador.Height = 4665
        Me.FrameTraspasoCalibrador.Width = 6555
        W = Me.FrameTraspasoCalibrador.Width
        H = Me.FrameTraspasoCalibrador.Height
    End If
End Sub


Private Sub FrameTraspasoTrazaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el trapaso de trazabilidad
    Me.FrameTrazabilidad.visible = visible
    If visible = True Then
        Me.FrameTrazabilidad.Top = -90
        Me.FrameTrazabilidad.Left = 0
        Me.FrameTrazabilidad.Height = 4665
        Me.FrameTrazabilidad.Width = 6555
        W = Me.FrameTrazabilidad.Width
        H = Me.FrameTrazabilidad.Height
    End If
End Sub

Private Sub FrameBajaSociosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para baja de socios
    Me.FrameBajaSocios.visible = visible
    If visible = True Then
        Me.FrameBajaSocios.Top = -90
        Me.FrameBajaSocios.Left = 0
        Me.FrameBajaSocios.Height = 3150
        Me.FrameBajaSocios.Width = 6615
        W = Me.FrameBajaSocios.Width
        H = Me.FrameBajaSocios.Height
    End If
End Sub


Private Sub FrameTraspasoFactCoopVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para baja de socios
    Me.FrameTraspasoFactCoop.visible = visible
    If visible = True Then
        Me.FrameTraspasoFactCoop.Top = -90
        Me.FrameTraspasoFactCoop.Left = 0
        Me.FrameTraspasoFactCoop.Height = 5490
        Me.FrameTraspasoFactCoop.Width = 6615
        W = Me.FrameTraspasoFactCoop.Width
        H = Me.FrameTraspasoFactCoop.Height
    End If
End Sub


Private Sub FrameTraspasoROPASVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para baja de socios
    Me.FrameTraspasoROPAS.visible = visible
    If visible = True Then
        Me.FrameTraspasoROPAS.Top = -90
        Me.FrameTraspasoROPAS.Left = 0
        Me.FrameTraspasoROPAS.Height = 5490
        Me.FrameTraspasoROPAS.Width = 6615
        W = Me.FrameTraspasoROPAS.Width
        H = Me.FrameTraspasoROPAS.Height
    End If
End Sub



Private Sub FrameKilosRecolectVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para baja de socios
    Me.FrameKilosRecolect.visible = visible
    If visible = True Then
        Me.FrameKilosRecolect.Top = -90
        Me.FrameKilosRecolect.Left = 0
        Me.FrameKilosRecolect.Height = 6480
        Me.FrameKilosRecolect.Width = 6615
        W = Me.FrameKilosRecolect.Width
        H = Me.FrameKilosRecolect.Height
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

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadSelect1 = ""
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
            cadParam = cadParam & campo & "{" & Tabla & ".codclase}" & "|"
            cadParam = cadParam & nomCampo & " {" & "clases" & ".nomclase}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Producto""" & "|"
            numParam = numParam + 3
            
        Case "Producto"
            cadParam = cadParam & campo & "{" & Tabla & ".codprodu}" & "|"
            cadParam = cadParam & nomCampo & " {" & "productos" & ".nomprodu}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3

        'Informe de calibres
        Case "Seccion"
            cadParam = cadParam & campo & "{" & Tabla & ".codsecci}" & "|"
            cadParam = cadParam & nomCampo & "{rseccion.nomsecci}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Seccion""" & "|"
            numParam = numParam + 3
            
        Case "Socio"
            cadParam = cadParam & campo & "{" & Tabla & ".codsocio}" & "|"
            cadParam = cadParam & nomCampo & " {" & "rsocios" & ".nomsocio}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Socio""" & "|"
            numParam = numParam + 3
            
        'Informe de calidades
        Case "Variedad"
            cadParam = cadParam & campo & "{" & Tabla & ".codvarie}" & "|"
            cadParam = cadParam & nomCampo & "{variedades.nomvarie}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Variedad""" & "|"
            numParam = numParam + 3
            
        Case "Calidad"
            cadParam = cadParam & campo & "{" & Tabla & ".codcalid}" & "|"
            cadParam = cadParam & nomCampo & " {" & "rcalidad" & ".nomcalid}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Calidad""" & "|"
            numParam = numParam + 3
            
            
        'Informe de campos
        Case "Socios"
            cadParam = cadParam & campo & "{rcampos.codsocio}" & "|"
            cadParam = cadParam & nomCampo & "{rsocios.nomsocio}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Socio""" & "|"
            numParam = numParam + 3
            
        Case "Clases"
            cadParam = cadParam & campo & "{variedades.codclase}" & "|"
            cadParam = cadParam & nomCampo & " {clases.nomclase}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3
            
        Case "Terminos"
            cadParam = cadParam & campo & "{rpartida.codpobla}" & "|"
            cadParam = cadParam & nomCampo & " {" & "rpueblos" & ".despobla}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Termino Municipal""" & "|"
            numParam = numParam + 3
            
        Case "Zonas"
            cadParam = cadParam & campo & "{rpartida.codzonas}" & "|"
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
Dim campo As String
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
            Tipo = "Código"
        Case "Alfabético"
            cadParam = cadParam & "Orden" & "= {" & Tabla
            Select Case OpcionListado
                Case 10
                    cadParam = cadParam & ".nomclien}|"
                Case 11
                    cadParam = cadParam & ".nomprove}|"
            End Select
            Tipo = "Alfabético"
    End Select
    
    numParam = numParam + 1

End Function

Private Sub AbrirFrmProducto(indice As Integer)
    indCodigo = indice
    Set frmProd = New frmComercial
    
    AyudaProductosCom frmProd, txtcodigo(indice).Text
    
    Set frmProd = Nothing
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

Private Sub AbrirFrmSituacionCampo(indice As Integer)
    indCodigo = indice
    Set frmSit = New frmManSituCamp
    frmSit.DatosADevolverBusqueda = "0|1|"
    frmSit.Show vbModal
    Set frmSit = Nothing
End Sub

Private Sub AbrirFrmSituacion(indice As Integer)
    indCodigo = indice
    Set frmSitu = New frmManSituacion
    frmSitu.DatosADevolverBusqueda = "0|1|"
    frmSitu.Show vbModal
    Set frmSitu = Nothing
End Sub


Private Sub AbrirFrmClase(indice As Integer)
    If indice = 6 Or indice = 7 Then
        indCodigo = indice + 14
    Else
        indCodigo = indice
    End If
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtcodigo(indice).Text
    
    Set frmCla = Nothing
End Sub



Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub


Private Sub AbrirFrmCooperativa(indice As Integer)
    indCodigo = indice
    Set frmCoop = New frmManCoope
    frmCoop.DatosADevolverBusqueda = "0|1|"
    frmCoop.Show vbModal
    Set frmCoop = Nothing
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


'Private Function DatosOk() As Boolean
'Dim b As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim vClien As CSocio
'' añadido
'Dim Mens As String
'Dim numfactu As String
'Dim numser As String
'Dim Fecha As Date
'
'    b = True
'    If txtCodigo(9).Text = "" Or txtCodigo(10).Text = "" Or txtCodigo(11).Text = "" Then
'        MsgBox "Debe introducir la letra de serie, el número de factura y la fecha de factura para localizar la factura a rectificar", vbExclamation
'        b = False
'    End If
'    If b And vParamAplic.Cooperativa = 2 Then
'        If txtCodigo(8).Text = "" Then
'            MsgBox "Debe introducir el cliente. Reintroduzca.", vbExclamation
'            b = False
'        Else
'            ' obtenemos la cooperativa del anterior cliente y del nuevo pq tienen que coincidir
'            ' anterior cliente
'            Sql = ""
'            Sql = DevuelveDesdeBDNew(cAgro, "ssocio", "codcoope", "codsocio", txtCodigo(12).Text, "N")
'            ' nuevo cliente
'            Sql2 = ""
'            Sql2 = DevuelveDesdeBDNew(cAgro, "ssocio", "codcoope", "codsocio", txtCodigo(8).Text, "N")
'            If Sql <> Sql2 Then
'                MsgBox "El nuevo cliente debe pertenecer al mismo colectivo que el cliente de la factura a rectificar. Reintroduzca.", vbExclamation
'                b = False
'            End If
'        End If
'    End If
'
''    If b And Contabilizada = 1 And vParamAplic.NumeroConta <> 0 And txtCodigo(8).Text <> "" Then 'comprobamos que la cuenta contable del nuevo cliente existe
''        Set vClien = New CSocio
''        If vClien.LeerDatos(txtCodigo(8).Text) Then
''            sql = ""
''            sql = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", vClien.CuentaConta, "T")
''            If sql = "" Then
''                MsgBox "La cuenta contable del nuevo cliente no existe. Revise", vbExclamation
''                b = False
''            End If
''        End If
''    End If
'
'' añadido
''    b = True
'
'    If ConTarjetaProfesional(txtCodigo(9).Text, txtCodigo(10).Text, txtCodigo(11).Text) Then
'        MsgBox "Este Factura tiene alguna tarjeta profesional, no se permite hacer la factura rectificativa", vbExclamation
'        b = False
'    Else
'        If txtCodigo(13).Text = "" Then
'            MsgBox "Debe introducir obligatoriamente una Fecha de Facturación.", vbExclamation
'            b = False
'            PonerFoco txtCodigo(13)
'        Else
'                If Not FechaDentroPeriodoContable(CDate(txtCodigo(13).Text)) Then
'                    Mens = "La Fecha de Facturación no es del ejercicio actual ni siguiente. Reintroduzca."
'                    MsgBox Mens, vbExclamation
'                    b = False
'                    PonerFoco txtCodigo(13)
'                Else
'                    'VRS:2.0.1(0)
'                    If Not FechaSuperiorUltimaLiquidacion(CDate(txtCodigo(13).Text)) Then
'                        Mens = "  La Fecha de Facturación es inferior a la última liquidación de Iva. " & vbCrLf & vbCrLf
'                        ' unicamente si el usuario es root el proceso continuará
'                        If vSesion.Nivel > 0 Then
'                            Mens = Mens & "  El proceso no continuará."
'                            MsgBox Mens, vbExclamation
'                            b = False
'                            PonerFoco txtCodigo(13)
'                        Else
'                            Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
'                            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                                b = False
'                                PonerFoco txtCodigo(13)
'                            End If
'                        End If
'                    End If
'                    ' la fecha de factura no debe ser inferior a la ultima factura de la serie
'                    numser = "letraser"
'                    numfactu = ""
'                    numfactu = DevuelveDesdeBDNew(cAgro, "stipom", "contador", "codtipom", "FAG", "T", numser)
'                    If numfactu <> "" Then
'                        If FechaFacturaInferiorUltimaFacturaSerieHco(CDate(txtCodigo(13).Text), CLng(numfactu), numser, 0) Then
'                            Mens = "La Fecha de Factura es inferior a la última factura de la serie." & vbCrLf & vbCrLf
'                            Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
'                            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                                b = False
'                                PonerFoco txtCodigo(13)
'                            End If
'                        End If
'                    End If
'                End If
'        End If
'    End If
'
'    DatosOk = b
'
'
'' end añadido
'    If b And txtCodigo(87).Text = "" Then
'        MsgBox "Para rectificar una factura ha de introducir obligatoriamente un motivo. Reintroduzca", vbExclamation
'        b = False
'    End If
'    DatosOk = b
'
'End Function
'


' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer
Dim Rs As ADODB.Recordset
Dim Sql As String


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
    
    'tipo de produccion
    Combo1(1).AddItem "Esperada"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Real"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
  
    'tipo de informe de entradas clasificadas
    Combo1(2).AddItem "Todas"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Sólo Clasificadas"
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

    'tipo de calibrador
    Select Case vParamAplic.Cooperativa
        Case 0 ' 0=catadau
            Combo1(6).AddItem "Calibrador Grande"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 0
            Combo1(6).AddItem "Calibrador Pequeño"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 1
        Case 1 ' 1=valsur
            Combo1(6).AddItem "Calibrador 1"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 0
        Case 4 '4=alzira
            Combo1(6).AddItem "Precalibrador"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 0
            Combo1(6).AddItem "Escandalladora"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 1
            Combo1(6).AddItem "Calibrador Kaki"
            Combo1(6).ItemData(Combo1(6).NewIndex) = 2
    End Select
    
    ' tipo de factura a traspasar
    'tipo de factura
    Sql = "select codtipom, nomtipom from usuarios.stipom where tipodocu > 0 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    While Not Rs.EOF
'        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
        Sql = Rs.Fields(1).Value
        Sql = Rs.Fields(0).Value & " - " & Sql
        Combo1(7).AddItem Sql 'campo del codigo
        Combo1(7).ItemData(Combo1(7).NewIndex) = I
        I = I + 1
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    

End Sub

Private Function ConcatenarCampos(cTabla As String, cWhere As String) As String
Dim Rs As ADODB.Recordset
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
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not Rs.EOF
        Sql1 = Sql1 & DBLet(Rs.Fields(0).Value, "N") & ","
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    'quitamos el ultimo or
    ConcatenarCampos = Mid(Sql1, 1, Len(Sql1) - 1)
    
End Function

Private Function CargarTemporal(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String

    
    On Error GoTo eCargarTemporal
    
    CargarTemporal = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

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
    If Opcion1(0) Then ' socios
        Sql1 = "select " & vUsu.Codigo & ",rcampos.codsocio, sum(kilosnet) from rentradas right join rcampos on rentradas.codcampo = rcampos.codcampo "
        Sql1 = Sql1 & " where rcampos.codcampo in (" & Sql & ")"
        Sql1 = Sql1 & " group by 1,2"
        
        Sql2 = "insert into tmpinformes (codusu, codigo1, importe1) " & Sql1
        conn.Execute Sql2
    End If
    
    If Opcion1(1) Then ' clases
        Sql1 = "select " & vUsu.Codigo & ",variedades.codclase, sum(kilosnet) from rentradas right join (rcampos inner join variedades on rcampos.codvarie = variedades.codvarie) on rentradas.codcampo = rcampos.codcampo where rcampos.codcampo in (" & Sql & ")"
        Sql1 = Sql1 & " group by 1,2"
        
        Sql2 = "insert into tmpinformes (codusu, codigo1, importe1) " & Sql1
        conn.Execute Sql2
    End If
    
    If Opcion1(2) Then ' terminos
        Sql1 = "select " & vUsu.Codigo & ", rpartida.codpobla, sum(kilosnet) from rentradas right join (rcampos inner join rpartida on rcampos.codparti = rpartida.codparti)  on rentradas.codcampo = rcampos.codcampo where rcampos.codcampo in (" & Sql & ")"
        Sql1 = Sql1 & " group by 1,2"
    
        Sql2 = "insert into tmpinformes (codusu, nombre1, importe1) " & Sql1
        conn.Execute Sql2
    End If
    
    If Opcion1(3) Then ' zonas
        Sql1 = "select " & vUsu.Codigo & ", rpartida.codzonas, sum(kilosnet) from rentradas right join (rcampos inner join rpartida on rcampos.codparti = rpartida.codparti)  on rentradas.codcampo = rcampos.codcampo where rcampos.codcampo in (" & Sql & ")"
        Sql1 = Sql1 & " group by 1,2"
    
        Sql2 = "insert into tmpinformes (codusu, codigo1, importe1) " & Sql1
        conn.Execute Sql2
    End If
    
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporal2(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
    
    On Error GoTo eCargarTemporal
    
    CargarTemporal2 = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rclasifica.numnotac FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    If cWhere <> "" Then
        Sql = "select distinct rclasifica.numnotac  from " & cTabla & " where " & cWhere
    Else
        Sql = "select distinct rclasifica.numnotac  from " & cTabla
    End If
    
    Select Case Combo1(2).ListIndex
        Case 0: 'todas
            Sql1 = "select " & vUsu.Codigo & ", rclasifica.numnotac, 0 from rclasifica where numnotac in (" & Sql & ")"
        
        Case 1: ' solo clasificadas
            Sql1 = "select " & vUsu.Codigo & ",rclasifica_clasif.numnotac, sum(rclasifica_clasif.kilosnet) from rclasifica_clasif inner join rclasifica on rclasifica_clasif.numnotac = rclasifica.numnotac "
            Sql1 = Sql1 & " where rclasifica.numnotac in (" & Sql & ")"
            Sql1 = Sql1 & " group by 1,2 "
            Sql1 = Sql1 & " having not sum(rclasifica_clasif.kilosnet)  is null "
            
        
        Case 2: ' pendientes
            Sql1 = "select " & vUsu.Codigo & ",rclasifica_clasif.numnotac, sum(rclasifica_clasif.kilosnet) from rclasifica_clasif inner join rclasifica on rclasifica_clasif.numnotac = rclasifica.numnotac "
            Sql1 = Sql1 & " where rclasifica.numnotac in (" & Sql & ")"
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
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Sql As String
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
    Sql = "Select rclasifica.numnotac FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    If cWhere <> "" Then
        Sql = "select distinct rclasifica.numnotac  from " & cTabla & " where " & cWhere
    Else
        Sql = "select distinct rclasifica.numnotac  from " & cTabla
    End If
    
  ' solo clasificadas
    Sql1 = "select rclasifica.numnotac, rclasifica.codvarie, rclasifica.codcampo, rclasifica.codsocio,sum(rclasifica_clasif.kilosnet) from rclasifica inner join rclasifica_clasif on rclasifica.numnotac = rclasifica_clasif.numnotac "
    Sql1 = Sql1 & " where rclasifica.numnotac in (" & Sql & ")"
    Sql1 = Sql1 & " group by 1,2,3,4 "
    Sql1 = Sql1 & " having not sum(rclasifica_clasif.kilosnet) is null "
        
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not Rs.EOF
        Clase = DevuelveDesdeBDNew(cAgro, "variedades", "codclase", "codvarie", Rs!CodVarie, "N")
        Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")
        Sql2 = Sql2 & " order by 1 "
        
        Set RS1 = New ADODB.Recordset
        
        res = ""
        Res1 = ""
        I = 0
        RS1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS1.EOF
            I = I + 1
            vSQL = "select kilosnet from rclasifica_clasif where numnotac= " & DBSet(Rs!numnotac, "N")
            vSQL = vSQL & " and codcalid = " & DBSet(RS1!codcalid, "N")
            
            res = res & "cal" & I & "," 'Format(Rs1!codcalid, "00") & ","
            Res1 = Res1 & DBSet(TotalRegistros(vSQL), "N") & ","
            
            RS1.MoveNext
        Wend
        
        Set RS1 = Nothing
        
        
        Sql2 = "insert into tmpclasifica (codusu, codcampo, codsocio, numnotac, codvarie, codclase, "
        Sql2 = Sql2 & Mid(res, 1, Len(res) - 1) & ") values ("
        Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!CodVarie, "N") & "," & DBSet(Clase, "N") & ","
        Sql2 = Sql2 & Mid(Res1, 1, Len(Res1) - 1) & ")"
        
        conn.Execute Sql2
        
        Rs.MoveNext
        
    Wend
    
    Set Rs = Nothing
    
    CargarTemporal3 = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function

'
' carga temporal para sacar informe de kilos / gastos de la rhisfruta
'
Private Function CargarTemporal4(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Sql As String
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
    Sql = "select rhisfruta.numalbar, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.codsocio, rhisfruta.kilosnet "
    Sql = Sql & " from " & cTabla
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
        
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not Rs.EOF
        Clase = DevuelveDesdeBDNew(cAgro, "variedades", "codclase", "codvarie", Rs!CodVarie, "N")
        Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")
        Sql2 = Sql2 & " order by 1 "
        
        Set RS1 = New ADODB.Recordset
        
        res = ""
        Res1 = ""
        I = 0
        RS1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS1.EOF
            I = I + 1
            vSQL = "select kilosnet from rhisfruta_clasif where numalbar= " & DBSet(Rs!numalbar, "N")
            vSQL = vSQL & " and codvarie = " & DBSet(Rs!CodVarie, "N")
            vSQL = vSQL & " and codcalid = " & DBSet(RS1!codcalid, "N")
            
            res = res & "cal" & I & ","
            Res1 = Res1 & DBSet(TotalRegistros(vSQL), "N") & ","
            
            RS1.MoveNext
        Wend
        
        Set RS1 = Nothing
        
        
        Sql2 = "insert into tmpclasifica (codusu, codcampo, codsocio, numnotac, codvarie, codclase, "
        Sql2 = Sql2 & Mid(res, 1, Len(res) - 1) & ") values ("
        Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs!CodVarie, "N") & "," & DBSet(Clase, "N") & ","
        Sql2 = Sql2 & Mid(Res1, 1, Len(Res1) - 1) & ")"
        
        conn.Execute Sql2
        
        Rs.MoveNext
        
    Wend
    
    Set Rs = Nothing
    
    CargarTemporal4 = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function

'
' carga temporal para sacar informe de kilos / gastos de la rhisfruta
'
Private Function CargarTemporal4New(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Sql As String
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

    Sql = "DROP TABLE IF EXISTS tmp; "
    conn.Execute Sql
    
'    SQL = "CREATE TABLE `tmp` ("
''    SQL = SQL & " `codcampo` int(7) default NULL,"
''    SQL = SQL & " `codsocio` int(7) default NULL,"
''    SQL = SQL & " `numnotac` int(7) default NULL,"
'    SQL = SQL & " `codvarie` int(6) default NULL,"
''    SQL = SQL & " `codclase` smallint(3) default NULL,"
''    SQL = SQL & " `nom1` varchar(3),"
'    SQL = SQL & " `kilcal1` int(7) default NULL,"
''    SQL = SQL & " `nom2` varchar(3),"
'    SQL = SQL & " `kilcal2` int(7) default NULL,"
''    SQL = SQL & " `nom3` varchar(3),"
'    SQL = SQL & " `kilcal3` int(7) default NULL,"
''    SQL = SQL & " `nom4` varchar(3),"
'    SQL = SQL & " `kilcal4` int(7) default NULL,"
''    SQL = SQL & " `nom5` varchar(3),"
'    SQL = SQL & " `kilcal5` int(7) default NULL,"
''    SQL = SQL & " `nom6` varchar(3),"
'    SQL = SQL & " `kilcal6` int(7) default NULL,"
''    SQL = SQL & " `nom7` varchar(3),"
'    SQL = SQL & " `kilcal7` int(7) default NULL,"
''    SQL = SQL & " `nom8` varchar(3),"
'    SQL = SQL & " `kilcal8` int(7) default NULL,"
''    SQL = SQL & " `nom9` varchar(3),"
'    SQL = SQL & " `kilcal9` int(7) default NULL,"
''    SQL = SQL & " `nom10` varchar(3),"
'    SQL = SQL & " `kilcal10` int(7) default NULL,"
''    SQL = SQL & " `nom11` varchar(3),"
'    SQL = SQL & " `kilcal11` int(7) default NULL,"
''    SQL = SQL & " `nom12` varchar(3),"
'    SQL = SQL & " `kilcal12` int(7) default NULL,"
''    SQL = SQL & " `nom13` varchar(3),"
'    SQL = SQL & " `kilcal13` int(7) default NULL,"
''    SQL = SQL & " `nom14` varchar(3),"
'    SQL = SQL & " `kilcal14` int(7) default NULL,"
''    SQL = SQL & " `nom15` varchar(3),"
'    SQL = SQL & " `kilcal15` int(7) default NULL,"
''    SQL = SQL & " `nom16` varchar(3),"
'    SQL = SQL & " `kilcal16` int(7) default NULL,"
''    SQL = SQL & " `nom17` varchar(3),"
'    SQL = SQL & " `kilcal17` int(7) default NULL,"
''    SQL = SQL & " `nom18` varchar(3),"
'    SQL = SQL & " `kilcal18` int(7) default NULL,"
''    SQL = SQL & " `nom19` varchar(3),"
'    SQL = SQL & " `kilcal19` int(7) default NULL,"
''    SQL = SQL & " `nom20` varchar(3),"
'    SQL = SQL & " `kilcal20` int(7) default NULL)"
'
'    conn.Execute SQL
    
    Sql2 = "delete from tmpclasifica2 where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "select rhisfruta.numalbar, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.codsocio, rhisfruta.kilosnet "
    Sql = Sql & " from " & cTabla
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
        
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not Rs.EOF
        Clase = DevuelveDesdeBDNew(cAgro, "variedades", "codclase", "codvarie", Rs!CodVarie, "N")
        Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")
        Sql2 = Sql2 & " order by 1 "
        
        Set RS1 = New ADODB.Recordset
        
        res = ""
        Res1 = ""
        I = 0
        RS1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS1.EOF Then
            While Not RS1.EOF
                I = I + 1
                vSQL = "select kilosnet from rhisfruta_clasif where numalbar= " & DBSet(Rs!numalbar, "N")
                vSQL = vSQL & " and codvarie = " & DBSet(Rs!CodVarie, "N")
                vSQL = vSQL & " and codcalid = " & DBSet(RS1!codcalid, "N")
                
                res = res & "nomcal" & I & "," & "kilcal" & I & ","
                Res1 = Res1 & DBSet(NombreCalidad(CStr(Rs!CodVarie), CStr(RS1!codcalid)), "T") & "," & DBSet(TotalRegistros(vSQL), "N") & ","
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            
            
            Sql2 = "insert into tmpclasifica2 (codusu, codcampo, codsocio, numnotac, codvarie, codclase, "
            Sql2 = Sql2 & Mid(res, 1, Len(res) - 1) & ") values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
            Sql2 = Sql2 & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs!CodVarie, "N") & "," & DBSet(Clase, "N") & ","
            Sql2 = Sql2 & Mid(Res1, 1, Len(Res1) - 1) & ")"
            
            conn.Execute Sql2
        End If
        
        Rs.MoveNext
        
    Wend
    
    Set Rs = Nothing
    
    Sql = "select codvarie, sum(kilcal1), sum(kilcal2) as kilos2, sum(kilcal3) as kilos3, sum(kilcal4) as kilos4, sum(kilcal5), sum(kilcal6), sum(kilcal7), sum(kilcal8), "
    Sql = Sql & " sum(kilcal9), sum(kilcal10), sum(kilcal11), sum(kilcal12), sum(kilcal13), sum(kilcal14), sum(kilcal15), sum(kilcal16),"
    Sql = Sql & " sum(kilcal17), sum(kilcal18), sum(kilcal19), sum(kilcal20) from tmpclasifica2 "
    Sql = Sql & " where codusu = " & vUsu.Codigo
    Sql = Sql & " group by 1 "
    
    
    Set RS1 = New ADODB.Recordset
    
    RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS1.EOF
        m = 1 ' para evitar que sean todos ceros y haya un bucle infinito
        I = 1
        
        While I < 20 And m < 40
            Sql = "select codvarie, sum(kilcal1), sum(kilcal2) as kilos2, sum(kilcal3) as kilos3, sum(kilcal4) as kilos4, sum(kilcal5), sum(kilcal6), sum(kilcal7), sum(kilcal8), "
            Sql = Sql & " sum(kilcal9), sum(kilcal10), sum(kilcal11), sum(kilcal12), sum(kilcal13), sum(kilcal14), sum(kilcal15), sum(kilcal16),"
            Sql = Sql & " sum(kilcal17), sum(kilcal18), sum(kilcal19), sum(kilcal20) from tmpclasifica2 "
            Sql = Sql & " where codusu = " & vUsu.Codigo
            Sql = Sql & " and codvarie = " & DBSet(RS1!CodVarie, "N")
            Sql = Sql & " group by 1 "
        
            Set Rs2 = New ADODB.Recordset
            
            Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            If DBLet(Rs2.Fields(I).Value, "N") = 0 Then
                Sql = "update tmpclasifica2 set kilcal" & I & "=kilcal" & I + 1 & ","
                Sql = Sql & " nomcal" & I & "=nomcal" & I + 1
                
                For J = I + 1 To 19
                    Sql = Sql & ", kilcal" & J & "=kilcal" & J + 1
                    Sql = Sql & ", nomcal" & J & "=nomcal" & J + 1
                Next J
                
                Sql = Sql & ", kilcal20=" & ValorNulo
                Sql = Sql & ", nomcal20=" & ValorNulo
                Sql = Sql & " where codvarie = " & DBSet(RS1.Fields(0).Value, "N")
                Sql = Sql & " and codusu = " & vUsu.Codigo
                
                conn.Execute Sql
                
            Else
                I = I + 1
          
            End If
            
            m = m + 1
            
            Set Rs2 = Nothing
            
        Wend
    
        RS1.MoveNext
    Wend
    
    Set RS1 = Nothing
    
    CargarTemporal4New = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporal5(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Sql As String
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
    
    On Error GoTo eCargarTemporal
    
    CargarTemporal5 = False

    Sql2 = "delete from tmpinfkilos where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
    End If
    
    Sql = "select distinct rcampos.codsocio, rcampos.codcampo "
    Sql = Sql & " from " & cTabla
    Sql = Sql & " where rcampos.fecbajas is null "
    If cWhere <> "" Then
        Sql = Sql & " and " & cWhere
    End If
    Sql = Sql & " union "
    Sql = Sql & " select distinct rhisfruta.codsocio, rhisfruta.codcampo "
    Sql = Sql & " from (" & cTabla & ") inner join rhisfruta on rcampos.codcampo = rhisfruta.codcampo and rcampos.codsocio = rhisfruta.codsocio "
    If cWhere <> "" Then
        Sql = Sql & " where " & cWhere
    End If
    If txtcodigo(39).Text <> "" Then Sql = Sql & " and rhisfruta.fecalbar >= " & DBSet(txtcodigo(39).Text, "F")
    If txtcodigo(40).Text <> "" Then Sql = Sql & " and rhisfruta.fecalbar <= " & DBSet(txtcodigo(40).Text, "F")
    Sql = Sql & " order by 1, 2"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql2 = "insert into tmpinfkilos (codusu, codprodu, codsocio, codcampo, kilosnet, porcen,"
    Sql2 = Sql2 & "canaforo, hanegada, hectarea, rdtohane, rdtohecta, nroarbol) values "
    
    While Not Rs.EOF
        SocioAct = DBLet(Rs.Fields(0).Value, "N")
        CampoAct = DBLet(Rs.Fields(1).Value, "N")
            
        Producto = ProductoCampo(DBLet(Rs.Fields(1).Value, "N"))
            
        Sql3 = "(" & vUsu.Codigo & "," & DBSet(Producto, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & ","
        
        SQLaux = "select sum(kilosnet) from rhisfruta where codsocio = " & DBSet(Rs.Fields(0).Value, "N")
        SQLaux = SQLaux & " and codcampo = " & DBSet(Rs.Fields(1).Value, "N")
        If txtcodigo(39).Text <> "" Then SQLaux = SQLaux & " and rhisfruta.fecalbar >= " & DBSet(txtcodigo(39).Text, "F")
        If txtcodigo(40).Text <> "" Then SQLaux = SQLaux & " and rhisfruta.fecalbar <= " & DBSet(txtcodigo(40).Text, "F")
        
        Sql3 = Sql3 & DBSet(DevuelveValor(SQLaux), "N") & ",0," 'kilosnet
        
        SqlAux2 = "select canaforo, "
        Select Case Combo1(5).ListIndex
            Case 0
                SqlAux2 = SqlAux2 & " supcoope, nroarbol"
            Case 1
                SqlAux2 = SqlAux2 & " supsigpa, nroarbol"
            Case 2
                SqlAux2 = SqlAux2 & " supcatas, nroarbol"
        End Select
        SqlAux2 = SqlAux2 & " from rcampos where codcampo = " & DBSet(Rs.Fields(1).Value, "N")
        
        Set RS1 = New ADODB.Recordset
        RS1.Open SqlAux2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RS1.EOF Then
            Sql3 = Sql3 & DBSet(RS1.Fields(0).Value, "N") & "," 'canaforo
            Ha = Round2(DBLet(RS1.Fields(1).Value, "N") / cFaneca, 2)
            Sql3 = Sql3 & DBSet(Ha, "N") & "," 'hanegadas
            Sql3 = Sql3 & DBSet(RS1.Fields(1).Value, "N") & ",0,0," 'hectareas
            Sql3 = Sql3 & DBSet(RS1.Fields(2).Value, "N") 'arboles
            Sql3 = Sql3 & "),"
        Else
            Sql3 = Sql3 & "0,0,0,0,0,0),"
        End If
        
        Sql2 = Sql2 & Sql3
        
        Set RS1 = Nothing
        
        
        Rs.MoveNext
    Wend

    'quitamos la ultima coma
    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
    conn.Execute Sql2
    
    CargarTemporal5 = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function



Private Function ActualizarRegistros(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim Sql As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "update " & QuitarCaracterACadena(cTabla, "_1") & " set impreso = 1 "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    conn.Execute Sql
    
    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function


Private Function NombreCalidad(Var As String, Calid As String) As String
Dim Sql As String

    NombreCalidad = ""

    Sql = "select nomcalab from rcalidad where codvarie = " & DBSet(Var, "N")
    Sql = Sql & " and codcalid = " & DBSet(Calid, "N")
    
    NombreCalidad = DevuelveValor(Sql)
    
End Function



Private Function GeneraFicheroAgriweb(pTabla As String, pWhere As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim vSocio As CSocio
Dim b As Boolean
Dim Nregs As Long
Dim total As Variant

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
    
    Set Rs = Nothing
    
    'Imprimimos las lineas
    Aux = "select  rcampos.codsocio, sum(rcampos.supsigpa) "
    Aux = Aux & " from " & cTabla
    Aux = Aux & " where " & vWhere
    Aux = Aux & " group by 1 "
    Aux = Aux & " order by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        'No hayningun registro
    Else
        b = True
        Regs = 0
        While Not Rs.EOF And b
            Regs = Regs + 1
            Set vSocio = New CSocio
            
            If vSocio.LeerDatos(DBLet(Rs!Codsocio, "N")) Then
                LineaAgriweb NFic, vSocio, Rs
            Else
                b = False
            End If
            
            Set vSocio = Nothing
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
            
    Close (NFic)
    
    If Regs > 0 Then GeneraFicheroAgriweb = True
    Exit Function
    
EGen:
    Set Rs = Nothing
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
    cad = cad & Format(txtcodigo(27).Text, "0000")             'p.2 año ejercicio
    cad = cad & "17"                                           'p.6 comunidad autonoma
    cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 9)    'p.8 cif empresa
    cad = cad & "OP"                                           'p.17 tipo de vendedor
    cad = cad & RellenaABlancos(txtcodigo(28).Text, True, 9)   'p.19 cif industria transformadora
    cad = cad & RellenaABlancos(Combo1(4).Text, True, 2)       'p.28 producto segun tabla
    cad = cad & RellenaAceros(ImporteSinFormato(txtcodigo(29).Text), True, 10)    'p.30 kilos contratados
    cad = cad & Format(txtcodigo(30).Text, "ddmmyyyy")         'p.40 fecha de contratacion
    cad = cad & RellenaAceros(ImporteSinFormato(CCur(txtcodigo(31).Text) * 100), False, 6)    'p.48 superficie
    cad = cad & Format(txtcodigo(32).Text, "00.00")            'p.54 precio
    
    Print #NFich, cad
End Sub

Private Sub LineaAgriweb(NFich As Integer, vSocio As CSocio, ByRef Rs As ADODB.Recordset)
Dim cad As String
Dim Areas As Long

    cad = "P"                                                'p.1 tipo de registro
    cad = cad & Format(txtcodigo(27).Text, "0000")           'p.2 año ejercicio
    cad = cad & "17"                                         'p.6 comunidad autonoma
    cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 9)  'p.8 cif empresa
    cad = cad & "OP"                                         'p.17 tipo de vendedor
    cad = cad & RellenaABlancos(txtcodigo(28).Text, True, 9) 'p.19 cif de la empresa transformadora
    cad = cad & RellenaABlancos(Combo1(4).Text, True, 2)     'p.28 codigo del producto
    cad = cad & RellenaABlancos(vSocio.Nombre, True, 40)     'p.30 nombre socio
    cad = cad & RellenaABlancos(vSocio.nif, True, 9)         'p.70 nif socio
    cad = cad & "PA"                                         'p.79 tipo productor
    cad = cad & RellenaAceros(ImporteSinFormato(CStr(Round2(DBLet(Rs.Fields(1).Value, "N") * 100, 0))), False, 6)   'p.81 superficie amparada
    
    Print #NFich, cad
End Sub

Private Function ProductoCampo(campo As String) As String
Dim Sql As String

    ProductoCampo = ""
    
    Sql = "select variedades.codprodu from rcampos inner join variedades on rcampos.codvarie = variedades.codvarie "
    Sql = Sql & " where rcampos.codcampo = " & DBSet(campo, "N")
    
    ProductoCampo = DevuelveValor(Sql)

End Function


'Private Function ProcesarFichero(nomFich As String, TipoCal As Byte) As Boolean
'Dim NF As Long
'Dim cad As String
'Dim I As Integer
'Dim longitud As Long
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim NumReg As Long
'Dim Sql As String
'Dim Sql1 As String
'Dim total As Long
'Dim v_cant As Currency
'Dim v_impo As Currency
'Dim v_prec As Currency
'Dim b As Boolean
'Dim NomFic As String
'
'    On Error GoTo eProcesarFichero
'
'
'    ProcesarFichero = False
'    NF = FreeFile
'
'    Open nomFich For Input As #NF
'
'    Line Input #NF, cad
'    I = 0
'
'    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
'    longitud = FileLen(nomFich)
'
'    pb1.visible = True
'    Me.pb1.Max = longitud
'    Me.Refresh
'    Me.pb1.Value = 0
'
'
'    b = True
'    While Not EOF(NF)
'        I = I + 1
'
'        Me.pb1.Value = Me.pb1.Value + Len(cad)
'        lblProgres(1).Caption = "Linea " & I
'        Me.Refresh
'
'        If vParamAplic.Cooperativa = 1 Then ' si es valsur
'            b = ProcesarLineaValsur(cad, Combo1(6).ListIndex)
'        Else ' si es catadau
'            b = ProcesarLineaCatadau(NF, cad, Combo1(6).ListIndex)
'            If Combo1(6).ListIndex = 0 Then I = I + 6
'        End If
'
'        If b = False Then
'            ProcesarFichero = False
'            Exit Function
'        End If
'
'        If Not EOF(NF) Then Line Input #NF, cad
'    Wend
'    Close #NF
'
'    If cad <> "" And b Then
'        If vParamAplic.Cooperativa = 1 Then ' si es valsur
'            b = ProcesarLineaValsur(cad, Combo1(6).ListIndex)
''        Else
''            b = ProcesarLineaCatadau(NF, Cad, Combo1(6).ListIndex)
'        End If
'        If b = False Then
'            ProcesarFichero = False
'            Exit Function
'        End If
'    End If
'
'    ProcesarFichero = b
'
'    pb1.visible = False
'    lblProgres(0).Caption = ""
'    lblProgres(1).Caption = ""
'
'eProcesarFichero:
'    If Err.Number <> 0 Then
'        MuestraError Err.Description
'    End If
'
'
'End Function


'Private Function ProcesarFicheroCatadau(nomDir As String) As Boolean
'Dim NF As Long
'Dim cad As String
'Dim I As Integer
'Dim longitud As Long
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim NumReg As Long
'Dim Sql As String
'Dim Sql1 As String
'Dim total As Long
'Dim v_cant As Currency
'Dim v_impo As Currency
'Dim v_prec As Currency
'Dim b As Boolean
'Dim NomFic As String
'
'    ProcesarFicheroCatadau = False
'
'    ' Muestra los nombres en C:\ que representan directorios.
'    NomFic = Dir(nomDir, vbDirectory)   ' Recupera la primera entrada.
'    Do While NomFic <> "" And b   ' Inicia el bucle.
'       ' Ignora el directorio actual y el que lo abarca.
'       If NomFic <> "." And NomFic <> ".." Then
'          ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
'          If (GetAttr(nomDir & NomFic) And vbDirectory) = vbDirectory Then
'            NF = FreeFile
'
'            Open nomDir & NomFic For Input As #NF
'
'            Line Input #NF, cad
'            I = 0
'            Dir
'            lblProgres(0).Caption = "Procesando Fichero: " & NomFic
'            longitud = FileLen(NomFic)
'
'            pb1.visible = True
'            Me.pb1.Max = longitud
'            Me.Refresh
'            Me.pb1.Value = 0
'
'
'            b = True
'            While Not EOF(NF)
'                I = I + 1
'
'                Me.pb1.Value = Me.pb1.Value + Len(cad)
'                lblProgres(1).Caption = "Linea " & I
'                Me.Refresh
'
'                b = ProcesarLineaCatadau(NF, cad, 1) '1=calibrador pequeño
'
'                If b = False Then
'                    ProcesarFicheroCatadau = False
'                    Exit Function
'                End If
'
'                Line Input #NF, cad
'            Wend
'            Close #NF
'
'            If cad <> "" And b Then
'                b = ProcesarLineaCatadau(NF, cad, 1) '1=calibrador pequeño
'                If b = False Then
'                    ProcesarFicheroCatadau = False
'                    Exit Function
'                End If
'            End If
'
'          End If   ' solamente si representa un directorio.
'       End If
'       NomFic = Dir   ' Obtiene siguiente entrada.
'    Loop
'
'
'    ProcesarFicheroCatadau = b
'
'    pb1.visible = False
'    lblProgres(0).Caption = ""
'    lblProgres(1).Caption = ""
'
'End Function
'
''[Monica] 22/09/2009 nuevo calibrador grande para Catadau
'Private Function ProcesarDirectorioCatadau(nomDir As String) As Boolean
'Dim NF As Long
'Dim cad As String
'Dim I As Integer
'Dim longitud As Long
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim NumReg As Long
'Dim Sql As String
'Dim Sql1 As String
'Dim total As Long
'Dim v_cant As Currency
'Dim v_impo As Currency
'Dim v_prec As Currency
'Dim b As Boolean
'Dim NomFic As String
'
'    ProcesarDirectorioCatadau = False
'    b = True
'    ' Muestra los nombres en C:\ que representan directorios.
'    NomFic = Dir(nomDir & "*.txt")  ' Recupera la primera entrada.
'
'    If Combo1(6).ListIndex = 0 Then
'    'CALIBRADOR GRANDE
'        Do While NomFic <> "" And b   ' Inicia el bucle.
'           ' Ignora el directorio actual y el que lo abarca.
'           If NomFic <> "." And NomFic <> ".." Then
'              ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
'              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
'
'                NF = FreeFile
'
'                Open nomDir & NomFic For Input As #NF
'
'                Line Input #NF, cad
'
'                lblProgres(0).Caption = "Procesando Fichero: " & NomFic
'                longitud = FileLen(nomDir & NomFic)
'
'                pb1.visible = True
'                Me.pb1.Max = longitud
'                Me.Refresh
'                Me.pb1.Value = 0
'
'                If cad <> "" Then
'                    b = ProcesarFicheroCatadauCGrande(NF, cad)
'                End If
'
'                Close #NF
'
'
'              End If   ' solamente si representa un directorio.
'           End If
'           NomFic = Dir   ' Obtiene siguiente entrada.
'        Loop
'    Else
'    'CALIBRADOR PEQUEÑO
'        Do While NomFic <> "" And b   ' Inicia el bucle.
'           ' Ignora el directorio actual y el que lo abarca.
'           If NomFic <> "." And NomFic <> ".." Then
'              ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
'              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
'
'                Sql = "delete from tmpcalibrador"
'                conn.Execute Sql
'
'                Sql = "load data local infile '" & Replace(nomDir & NomFic, "\", "/") & "' into table `tmpcalibrador` fields escaped by '\\' enclosed by '""' lines terminated by '\r\n' ( `numnota`, `fecnota`, `albaran`,`porcen1`,`porcen2`,`kilos1`, `kilos2`, `kilos3`,`numnota2`,`export`,`nomcalid`,`kilos4`,`kilos5`)  "
'                conn.Execute Sql
'
'                Sql = "delete from tmpcalibrador where numnota = ''"
'                conn.Execute Sql
'
'                lblProgres(0).Caption = "Procesando Fichero: " & NomFic
'                longitud = TotalRegistros("select count(*) from tmpcalibrador")
'
'                pb1.visible = True
'                Me.pb1.Max = longitud
'                Me.Refresh
'                Me.pb1.Value = 0
'
'                If longitud <> 0 Then
'                    b = ProcesarFicheroCatadauCPequeño()
'                End If
'
'              End If   ' solamente si representa un directorio.
'           End If
'           NomFic = Dir   ' Obtiene siguiente entrada.
'        Loop
'
'
'    End If
'
'    ProcesarDirectorioCatadau = b
'
'    pb1.visible = False
'    lblProgres(0).Caption = ""
'    lblProgres(1).Caption = ""
'
'End Function


'
' Proceso de traspaso para CATADAU
'
'Private Function ProcesarLineaCatadau(NF As Long, cad As String, Calibr As Byte) As Boolean
'Dim b As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim Mens As String
'Dim numlinea As Long
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim RSaux As ADODB.Recordset
'
'Dim Codsoc As String
'Dim Codcam As String
'Dim codpro As String
'Dim codVar As String
'Dim Observ As String
'Dim Notaca As String
'Dim Kilone As String
'
'Dim Destri As String
'Dim Podrid As String
'Dim Pequen As String
'Dim Muestra As String
'
'Dim NGrupos As String
'Dim Nombre1 As String
'Dim Kilos As String
'
'
'Dim I As Integer
'Dim Situacion As Byte
'
'Dim NomCal As Dictionary
'Dim KilCal As Dictionary
'
'Dim SQLaux As String
'Dim Nregs As Integer
'
'Dim SeInserta As Boolean
'
'
'    On Error GoTo eProcesarLineaCatadau
'
'    ProcesarLineaCatadau = False
'
'    Codsoc = 0
'    Codcam = 0
'    codpro = 0
'    codVar = 0
'    Observ = ""
'    Notaca = 0
'    Kilone = 0
'
'    Destri = 0
'    Podrid = 0
'    Pequen = 0
'
'    ' inicializamos las variables
'    Set NomCal = New Dictionary
'    Set KilCal = New Dictionary
'
'    Select Case Calibr
'        Case 0  ' CALIBRADOR GRANDE
'            'primera linea: cabecera
'            If cad <> "" Then
'                Notaca = RecuperaValorNew(cad, ",", 5)
'
'                Kilone = RecuperaValorNew(cad, ",", 6)
'                Destri = RecuperaValorNew(cad, ",", 11)
'                Podrid = RecuperaValorNew(cad, ",", 9)
'                Pequen = RecuperaValorNew(cad, ",", 10)
'
'                Sql = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
'                Set Rs = New ADODB.Recordset
'                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Rs.EOF Then
'                    Observ = "NOTA NO EXISTE"
'                    Situacion = 2
'                Else
'                    If DBLet(Rs.Fields(0).Value, "N") <> Kilone Then
'                        Observ = "K.NETOS DIF."
'                        Situacion = 4
'                    End If
'                End If
'                ' salto tipo b
'                Line Input #NF, cad
'
'                Me.pb1.Value = Me.pb1.Value + Len(cad)
'                lblProgres(1).Caption = "Linea " & I
'                Me.Refresh
'
'                ' salto tipo c
'                Line Input #NF, cad
'
'                Me.pb1.Value = Me.pb1.Value + Len(cad)
'                lblProgres(1).Caption = "Linea " & I
'                Me.Refresh
'
'                NGrupos = RecuperaValorNew(cad, ",", 4)
'
'                'salto tipo d
'                Line Input #NF, cad
'
'                Me.pb1.Value = Me.pb1.Value + Len(cad)
'                lblProgres(1).Caption = "Linea " & I
'                Me.Refresh
'
'                cad = cad & ","
'                For I = 0 To NGrupos - 1
'                    Nombre1 = RecuperaValorNew(cad, ",", 4 + I)
'
'
'                    If Situacion <> 2 Then
'                        ' si hay nota asociada busco los datos
'
'                        Sql = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(Rs!CodVarie, "N")
'                        Sql = Sql & " and nomcalibrador1 = " & DBSet(Nombre1, "T")
'
'                        Set RS1 = New ADODB.Recordset
'                        RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                        If RS1.EOF Then
'                            Observ = "NO EXIS.CAL"
'                            Situacion = 1
'                        Else
'                            NomCal(I) = DBLet(RS1!codcalid, "N")
'                        End If
'                        Set RS1 = Nothing
'                    End If
'
'                Next I
'
'                ' salto tipo e
'                Line Input #NF, cad
'
'                Me.pb1.Value = Me.pb1.Value + Len(cad)
'                lblProgres(1).Caption = "Linea " & I
'                Me.Refresh
'
'                ' salto tipo f: pesos de la calidad
'                Line Input #NF, cad
'                Me.pb1.Value = Me.pb1.Value + Len(cad)
'                lblProgres(1).Caption = "Linea " & I
'                Me.Refresh
'
'                cad = cad & ","
'                For I = 0 To NGrupos - 1
'                    KilCal(I) = RecuperaValorNew(cad, ",", I + 4)
'                Next I
'
'                Sql = "select count(*) from rclasifauto where numnotac = " & Notaca
'
'                SeInserta = (TotalRegistros(Sql) = 0)
'
'                If SeInserta Then
'                    If Situacion = 2 Then
'                        ' si no hay nota asociada no puedo meter la clasificacion
'                        Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'                        Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'                        Sql = Sql & "`observac`,`situacion`) values ("
'                        Sql = Sql & DBSet(Notaca, "N") & ","
'                        Sql = Sql & DBSet(0, "N") & ","
'                        Sql = Sql & DBSet(0, "N") & ","
'                        Sql = Sql & DBSet(0, "N") & ","
'                        Sql = Sql & DBSet(Kilone, "N") & ","
'                        Sql = Sql & DBSet(Destri, "N") & ","
'                        Sql = Sql & DBSet(Podrid, "N") & ","
'                        Sql = Sql & DBSet(Pequen, "N") & ","
'                        Sql = Sql & DBSet(Observ, "T") & ","
'                        Sql = Sql & DBSet(Situacion, "N") & ")"
'
'                    Else
'                        ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
'                        ' tabla: rclasifauto
'                        Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'                        Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'                        Sql = Sql & "`observac`,`situacion`) values ("
'                        Sql = Sql & DBSet(Notaca, "N") & ","
'                        Sql = Sql & DBSet(Rs!Codsocio, "N") & ","
'                        Sql = Sql & DBSet(Rs!codCampo, "N") & ","
'                        Sql = Sql & DBSet(Rs!CodVarie, "N") & ","
'                        Sql = Sql & DBSet(Kilone, "N") & ","
'                        Sql = Sql & DBSet(Destri, "N") & ","
'                        Sql = Sql & DBSet(Podrid, "N") & ","
'                        Sql = Sql & DBSet(Pequen, "N") & ","
'                        Sql = Sql & DBSet(Observ, "T") & ","
'                        Sql = Sql & DBSet(Situacion, "N") & ")"
'                    End If
'                    conn.Execute Sql
'
'                    ' tabla: rclasifauto_clasif
'                    Sql = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
'                    Sql = Sql & " values "
'
'                End If
'
'                'solo si tenemos nota asociada metemos toda la clasificacion
'                If Situacion <> 2 Then
'
'
'                    'borramos la tabla temporal
'                    SQLaux = "delete from tmpcata"
'                    conn.Execute SQLaux
'
'                    ' cargamos la tabla temporal
'                    For I = 0 To NGrupos - 1
'                        If NomCal(I) <> "" Then
'                            Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(I), "N"))
'                            If Nregs = 0 Then
'                                'insertamos en la temporal
'                                SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(I), "N")
'                                SQLaux = SQLaux & "," & KilCal(I) & ")"
'
'                                conn.Execute SQLaux
'                            Else
'                                'actualizamos la temporal
'                                SQLaux = "update tmpcata set kilosnet = kilosnet + " & KilCal(I)
'                                SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(I), "N")
'
'                                conn.Execute SQLaux
'                            End If
'                        End If
'                    Next I
'
'                    SQLaux = "select * from tmpcata order by codcalid"
'
'                    Set RSaux = New ADODB.Recordset
'                    RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                    Sql2 = ""
'
'                    While Not RSaux.EOF
'                        If SeInserta Then
'                            Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(Rs!CodVarie, "N") & ","
'                            Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
'                        Else
'                            Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
'                            Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
'                            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'                            Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
'
'                            conn.Execute Sql2
'                        End If
'
'                        RSaux.MoveNext
'                    Wend
'
'                    Set RSaux = Nothing
'
'
'                    If SeInserta Then
'                        If Sql2 <> "" Then
'                            Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
'                        End If
'                        Sql = Sql & Sql2
'                        conn.Execute Sql
'                    End If
'                End If ' si la situacion es distinta de 2
'
'
'' 18-05-2009
''                Sql2 = ""
''                For I = 0 To NomCal.Count - 1
''                    Sql2 = "(" & DBSet(Notaca, "N") & "," & DBSet(rs!CodVarie, "N") & ","
''                    Sql2 = Sql2 & DBSet(NomCal(I), "N") & "," & DBSet(KilCal(I), "N") & "),"
''                Next I
''                If Sql2 <> "" Then
''                    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
''                End If
''                SQL = SQL & Sql2
''                conn.Execute SQL
'
'                ' salto tipo g
'                Line Input #NF, cad
'
'                Set Rs = Nothing
'                Set NomCal = Nothing
'                Set KilCal = Nothing
'
'            Else
'                Exit Function
'            End If
'
'        Case 1 ' CALIBRADOR PEQUEÑO
'            ' saltamos 5 lineas mas
'            For I = 1 To 5
'                Line Input #NF, cad
'            Next I
'            Muestra = cad
'            ' saltamos para kilosnetos
'            Line Input #NF, cad
'            Kilone = cad
'            ' saltamos para podrido
'            Line Input #NF, cad
'            Podrid = cad
'            ' saltamos para destrio
'            Line Input #NF, cad
'            Destri = cad
'
'            Kilos = CCur(ImporteSinFormato(Kilone)) - CCur(ImporteSinFormato(Podrid)) - CCur(ImporteSinFormato(Destri))
'
'            ' saltamos para nota de campo
'            Line Input #NF, cad
'
'
''****************falsta esto
''            Notaca = Mid(NomFic, 1, 7)
'
'            Sql = "select codsocio, codcampo, codvarie, kilosnet from rclasifica"
'            Sql = Sql & " where numnotac = " & DBSet(Notaca, "N")
'
'            Set RS1 = New ADODB.Recordset
'            RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            If RS1.EOF Then
'                Observ = "NOTA NO EXI."
'                Situacion = 2
'            Else
'                If DBLet(RS1!KilosNet, "N") < Kilos Then
'                    Observ = "K.NETOS DIF."
'                    Situacion = 4
'                End If
'            End If
'            ' ++++++++++++++++++++estoy aqui linea 360 de agre1104
'
'
'    End Select
'    ProcesarLineaCatadau = True
'    Exit Function
'
'
'eProcesarLineaCatadau:
'    If Err.Number <> 0 Then
'        ProcesarLineaCatadau = False
'        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
'    End If
'End Function
'
''
'' Proceso de traspaso para VALSUR
''
'Private Function ProcesarLineaValsur(cad As String, Calibrador As Byte) As Boolean
'Dim Rs As ADODB.Recordset
'Dim Rs2 As ADODB.Recordset
'Dim Sql As String
'Dim Sql2 As String
'Dim Sql3 As String
'
'Dim NumNota As String
'Dim KilosNet As String
'Dim KilosDes As String
'Dim KilosPod As String
'Dim KilosTot As String
'
'Dim NomCal As Dictionary
'Dim KilCal As Dictionary
'
'Dim Situacion As Byte
'
'Dim CodCal As Integer
'Dim Observac As String
'Dim KilosNota As Long
'
'Dim I As Integer
'
'Dim c_Cantidad As Currency
'Dim c_Importe As Currency
'Dim c_Precio As Currency
'Dim Mens As String
'Dim numlinea As Long
'
'    On Error GoTo eProcesarLineaValsur
'
'    ProcesarLineaValsur = True
'
'    ' inicializamos las variables
'    Set NomCal = New Dictionary
'    Set KilCal = New Dictionary
'
'    NumNota = 0
'    KilosNet = 0
'    KilosDes = 0
'    KilosPod = 0
'    KilosTot = 0
'    Observac = ""
'    Situacion = 0
'
'    NumNota = RecuperaValor(cad, 3)
'    KilosNet = RecuperaValor(cad, 4)
'    KilosDes = RecuperaValor(cad, 17)
'    KilosPod = RecuperaValor(cad, 18)
'    KilosTot = RecuperaValor(cad, 19)
'
'    For I = 1 To 12
'        NomCal(I) = RecuperaValor(cad, I + 4)
'        KilCal(I) = RecuperaValor(cad, I + 19)
'    Next I
'
'    Sql = "select codsocio, codcampo, codvarie from rclasifica where numnotac = " & DBSet(NumNota, "N")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Rs.EOF Then
'        Observac = "NOTA NO EXISTE"
'        Situacion = 2
'
'        'insertamos la cabecera de la clasificacion
'        Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`,"
'        Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,`observac`,`situacion` ) values ("
'        Sql = Sql & DBSet(NumNota, "N") & ","
'        Sql = Sql & ValorNulo & ","
'        Sql = Sql & 0 & ","
'        Sql = Sql & ValorNulo & ","
'        Sql = Sql & DBSet(KilosTot, "N") & ","
'        Sql = Sql & DBSet(KilosDes, "N") & ","
'        Sql = Sql & DBSet(KilosPod, "N") & ","
'        Sql = Sql & DBSet(KilosNet, "N") & ","
'        Sql = Sql & DBSet(Observac, "T") & ","
'        Sql = Sql & DBSet(Situacion, "N") & ")"
'
'        conn.Execute Sql
'
'        ' no metemos la clasificacion pq no se corresponde con ninguna nota
'    Else
'        ' insertamos las calidades si existen
'        For I = 1 To 12
'            If NomCal(I) <> "" And KilCal(I) <> 0 Then
'                Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")
'                Select Case Combo1(6).ListIndex
'                    Case 0 ' calibrador 1
'                        Sql2 = Sql2 & " and nomcalibrador1 = " & DBSet(NomCal(I), "T")
'                    Case 1 ' calibrador 2
'                        Sql2 = Sql2 & " and nomcalibrador2 = " & DBSet(NomCal(I), "T")
'                End Select
'                Set Rs2 = New ADODB.Recordset
'                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                If Not Rs2.EOF Then
'                    CodCal = DBLet(Rs2!codcalid, "N")
'                    Situacion = 0
'                Else
''                    CodCal = 999
''                    Observac = "NO EXIS.CAL."
''                    Situacion = 1
'                    MsgBox "No existe la calidad " & NomCal(I) & ".Revise.", vbExclamation
'
'                    ProcesarLineaValsur = False
'
'                    Set Rs = Nothing
'                    Set Rs2 = Nothing
'
'                    Set NomCal = Nothing
'                    Set KilCal = Nothing
'
'                    Exit Function
'                End If
'
'                Set Rs2 = Nothing
'
'                Sql3 = "insert into rclasifauto_clasif(numnotac,codvarie,codcalid,kiloscal) values ("
'                Sql3 = Sql3 & DBSet(NumNota, "N") & ","
'                Sql3 = Sql3 & DBSet(Rs!CodVarie, "N") & ","
'                Sql3 = Sql3 & DBSet(CodCal, "N") & ","
'                Sql3 = Sql3 & DBSet(KilCal(I), "N") & ")"
'
'                conn.Execute Sql3
'            End If
'        Next I
'
'        'insertamos la cabecera de la clasificacion
'        Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`,"
'        Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,`observac`,`situacion`) values ("
'        Sql = Sql & DBSet(NumNota, "N") & ","
'        Sql = Sql & DBSet(Rs!Codsocio, "N") & ","
'        Sql = Sql & DBSet(Rs!codCampo, "N") & ","
'        Sql = Sql & DBSet(Rs!CodVarie, "N") & ","
'        Sql = Sql & DBSet(KilosTot, "N") & ","
'        Sql = Sql & DBSet(KilosDes, "N") & ","
'        Sql = Sql & DBSet(KilosPod, "N") & ","
'        Sql = Sql & DBSet(KilosNet, "N") & ","
'        Sql = Sql & DBSet(Observac, "T") & ","
'        Sql = Sql & DBSet(Situacion, "N") & ")"
'
'        conn.Execute Sql
'
'    End If
'
'    Set Rs = Nothing
'
'    Set NomCal = Nothing
'    Set KilCal = Nothing
'
'eProcesarLineaValsur:
'    If Err.Number <> 0 Then
'        ProcesarLineaValsur = False
'        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
'    End If
'End Function



Private Function ComprobarErrores() As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim Sql1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim Mens As String
Dim Tipo As Integer


    On Error GoTo eComprobarErrores

    ComprobarErrores = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    

    I = 0
    lblProgres(2).Caption = "Comprobando errores Tabla temporal entradas "
    
    Sql = "select count(*) from tmpentrada"
    longitud = TotalRegistros(Sql)

    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0
    
    
    Sql = "select * from tmpentrada"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText


    b = True
    I = 0
    While Not Rs.EOF And b
        I = I + 1

        Me.pb2.Value = Me.pb2.Value + 1
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh

        ' comprobamos que no exista el albaran en rclasifica
        Sql = "select count(*) from rclasifica where numnotac = " & DBSet(Rs!numalbar, "N")
        If TotalRegistros(Sql) > 0 Then
            Mens = "Nro. de Nota ya existe"
            Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Albarán:" & Rs!numalbar, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que no exista el albaran en el historico
        Sql = "select numalbar from rhisfruta_entradas where numnotac = " & DBSet(Rs!numalbar, "N")
        If DevuelveValor(Sql) <> 0 Then
            Mens = "Nro.Nota existe en hco. albarán:" & DevuelveValor(Sql)
            Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Albarán:" & Rs!numalbar, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If



        ' comprobamos que exista el socio
        Sql = "select count(*) from rsocios where codsocio = " & DBSet(Rs!Codsocio, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Socio no existe"
            Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Socio:" & Rs!Codsocio, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que exista la variedad
        Sql = "select count(*) from variedades where codvarie = " & DBSet(Rs!CodVarie, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Variedad no existe"
            Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Variedad:" & Rs!CodVarie, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que exista el campo
        Sql = "select count(*) from rcampos where codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql = Sql & " and nrocampo = " & DBSet(Rs!codCampo, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!CodVarie, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Campo no existe"
            Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet(("Socio:" & Rs!Codsocio & "-Campo:" & Rs!codCampo) & "-Variedad:" & Rs!CodVarie, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que no exista mas de un campo con ese numero de orden campo (scampo.codcampo MB)
        Sql = "select count(*) from rcampos where codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql = Sql & " and nrocampo = " & DBSet(Rs!codCampo, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!CodVarie, "N")
        If TotalRegistros(Sql) > 1 Then
            Mens = "Campo con más de un registro"
            Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet(("Socio:" & Rs!Codsocio & "-Campo:" & Rs!codCampo) & "-Variedad:" & Rs!CodVarie, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
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
'                    Mens = "Campo sin clasificación "
'                    SQL = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
'                          vUsu.Codigo & "," & DBSet(("Nro.Campo:" & Rs!CodCampo) & "-Variedad:" & Rs!CodVarie, "T") & "," & DBSet(Mens, "T") & ")"
'                    conn.Execute SQL
'                End If
'            Else ' es en almacen
'                SQL = "select count(*) from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")
'                If TotalRegistros(SQL) = 0 Then
'                    Mens = "Variedad sin calidades para clasificación "
'                    SQL = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
'                          vUsu.Codigo & "," & DBSet(("Nro.Campo:" & Rs!CodCampo) & "-Variedad:" & Rs!CodVarie, "T") & "," & DBSet(Mens, "T") & ")"
'                    conn.Execute SQL
'                End If
'            End If
'        End If

        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    
    lblProgres(2).Caption = "Comprobando errores Tabla temporal clasifica "
    
    Sql = "select count(*) from tmpclasific"
    longitud = TotalRegistros(Sql)

    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0

    Sql = "select * from tmpclasific"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True
    I = 0
    While Not Rs.EOF And b
        I = I + 1

        Me.pb2.Value = Me.pb2.Value + 1
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh

        ' comprobamos que no exista el albaran en rclasifica
        Sql = "select count(*) from rclasifica where numnotac = " & DBSet(Rs!numalbar, "N")
        If TotalRegistros(Sql) > 0 Then
            Mens = "Nro. de Nota ya existe"
            Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Albarán:" & Rs!numalbar, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que exista la variedad
        Sql = "select count(*) from variedades where codvarie = " & DBSet(Rs!CodVarie, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Variedad no existe"
            Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet("Variedad:" & Rs!CodVarie, "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que exista la calidad
        Sql = "select count(*) from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")
        Sql = Sql & " and codcalid = " & DBSet(Rs!codcalir, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Calidad no existe"
            Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                  vUsu.Codigo & "," & DBSet(("Variedad:" & Rs!CodVarie & "-Calidad:" & Rs!codcalir), "T") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If


        Rs.MoveNext
    Wend
    Set Rs = Nothing
    

    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    ComprobarErrores = b
    Exit Function

eComprobarErrores:
    ComprobarErrores = False
End Function



Private Function CargarTablasTemporales(nomFich1 As String, nomFich2 As String) As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim Variedad As String
Dim HoraEntrada As String

Dim Sql3 As String
Dim campo As String

    On Error GoTo eCargarTablasTemporales
    
    CargarTablasTemporales = False
    
'    SQL = "DROP TABLE IF EXISTS tmpEntrada; "
'    conn.Execute SQL
'
'    SQL = "DROP TABLE IF EXISTS tmpClasific; "
'    conn.Execute SQL
'
'
'    SQL = "CREATE TEMPORARY TABLE tmpEntrada ("
'    SQL = SQL & " codsocio int, codcampo int, numalbar int, codvarie int, fecalbar date, "
'    SQL = SQL & " horalbar datetime, kilosbru int, kilosnet int, numcajon int) "
'    conn.Execute SQL
'
'    SQL = "CREATE TEMPORARY TABLE tmpClasific ("
'    SQL = SQL & " numalbar int, codvarie int, codcalir int, porcenta decimal(5,2)) "
'    conn.Execute SQL
'
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

    Sql = "insert into tmpentrada(codsocio, codcampo, numalbar, codvarie, fecalbar, "
    Sql = Sql & "horalbar, kilosbru, kilosnet, numcajon) values  "
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
    
    Sql = Sql & Mid(Sql2, 1, Len(Sql2) - 1)
    conn.Execute Sql
    
    
    
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

    Sql = "insert into tmpclasific(numalbar, codvarie, codcalir, porcenta) values  "
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
    
    
    Sql = Sql & Mid(Sql2, 1, Len(Sql2) - 1)
    conn.Execute Sql
    
    
    
    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    CargarTablasTemporales = True
    Exit Function

eCargarTablasTemporales:
    CargarTablasTemporales = False
End Function


Private Function CargarClasificacion() As Boolean
Dim Sql As String
Dim Sql1 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
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
    
    Sql = "select count(*) from tmpentrada order by numalbar"
    longitud = TotalRegistros(Sql)
    
    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0
    
    
    Sql = "select * from tmpentrada order by numalbar"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Me.pb2.Value = Me.pb2.Value + 1
        lblProgres(3).Caption = "Albarán " & DBLet(Rs!numalbar, "N")
        Me.Refresh
        
'        Sql1 = "select imptrans from rportespobla, rpartida, rcampos, variedades "
'        Sql1 = Sql1 & " where rpartida.codparti = rcampos.codparti and "
'        Sql1 = Sql1 & " variedades.codprodu = rportespobla.codprodu and "
'        Sql1 = Sql1 & " rpartida.codpobla = rportespobla.codpobla and "
'        Sql1 = Sql1 & " variedades.codvarie = " & DBSet(rs!CodVarie, "N") & " and "
'        Sql1 = Sql1 & " rcampos.nrocampo = " & DBSet(rs!CodCampo, "N") & " and "
'        Sql1 = Sql1 & " rcampos.codvarie = variedades.codvarie "
'
'        Set Rs2 = New ADODB.Recordset
'        Rs2.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        Precio = 0
'        If Not Rs2.EOF Then
'            Precio = DBLet(Rs2.Fields(0).Value, "N")
'        End If
'
'        Set Rs2 = Nothing
'
'        Transporte = Round2(DBLet(rs!KilosNet, "N") * Precio, 2)
        
        Transporte = 0
    
        Sql = "insert into rclasifica (numnotac,fechaent,horaentr,codvarie,codsocio,codcampo,tipoentr,recolect,codtrans,codcapat,"
        Sql = Sql & "codtarif,kilosbru,numcajon,kilosnet,observac,"
        Sql = Sql & "imptrans,impacarr,imprecol,imppenal,tiporecol,horastra,numtraba,numalbar,fecalbar,impreso) values "
    
        campo = 0
        campo = DevuelveValor("select codcampo from rcampos where nrocampo = " & DBSet(Rs!codCampo, "N") & " and codsocio=" & DBSet(Rs!Codsocio, "N") & " and codvarie=" & DBSet(Rs!CodVarie, "N"))
    
        Sql = Sql & "(" & DBSet(Rs!numalbar, "N") & ","
        Sql = Sql & DBSet(Rs!fecalbar, "F") & ","
        Sql = Sql & DBSet(Rs!horalbar, "FH") & ","
        Sql = Sql & DBSet(Rs!CodVarie, "N") & ","
        Sql = Sql & DBSet(Rs!Codsocio, "N") & ","
'        Sql = Sql & DBSet(Rs!codCampo, "N") & ","
        Sql = Sql & DBSet(campo, "N") & ","
        Sql = Sql & "0," ' tipoentr 0=normal
        Sql = Sql & "1," ' recolect 1=socio
        Sql = Sql & ValorNulo & "," 'transportista
        Sql = Sql & ValorNulo & "," 'capataz
        Sql = Sql & ValorNulo & "," 'tarifa
        Sql = Sql & DBSet(Rs!KilosBru, "N") & ","
        Sql = Sql & DBSet(Rs!NumCajon, "N") & ","
        Sql = Sql & DBSet(Rs!KilosNet, "N") & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & DBSet(Transporte, "N") & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & "0," 'tiporecol 0=horas 1=destajo no admite valor nulo
        Sql = Sql & ValorNulo & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & DBSet(Rs!numalbar, "N") & ","
        Sql = Sql & DBSet(Rs!fecalbar, "F") & ",0)"
        
        conn.Execute Sql
        
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing


    ' 21-05-2009: cargamos las clasificacion dependiendo de si es por campo o almacen de aquellas que
    ' no tengan clasificacion
    Sql = "select * from tmpentrada order by numalbar "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Sql = "select count(*) from tmpclasific where numalbar = " & DBSet(Rs!numalbar, "N")
        If TotalRegistros(Sql) = 0 Then ' si no hay clasificacion en el fichero metemos la correspondiente
            Tipo = DevuelveValor("select tipoclasifica from variedades where codvarie = " & DBSet(Rs!CodVarie, "N"))
            If Tipo = 0 Then ' clasificacion en campo
                campo = 0
                campo = DevuelveValor("select codcampo from rcampos where nrocampo = " & DBSet(Rs!codCampo, "N") & " and codsocio=" & DBSet(Rs!Codsocio, "N") & " and codvarie=" & DBSet(Rs!CodVarie, "N"))

                Sql = "insert into tmpclasific (numalbar, codvarie, codcalir, porcenta) "
                Sql = Sql & " select " & DBSet(Rs!numalbar, "N") & ", codvarie, codcalid, muestra "
                Sql = Sql & " from rcampos_clasif where codcampo = " & DBSet(campo, "N")

                conn.Execute Sql
            Else ' clasificacion en almacen
                Sql = "insert into tmpclasific (numalbar, codvarie, codcalir, porcenta) "
                Sql = Sql & " select " & DBSet(Rs!numalbar, "N") & ", codvarie, codcalid, 0 "
                Sql = Sql & " from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")

                conn.Execute Sql
            End If
        End If
        Rs.MoveNext
    Wend
    Set Rs = Nothing

    ' 21-05-2009
    
    lblProgres(2).Caption = "Cargando Clasificación"
    
    Sql = "select count(*) from tmpclasific, tmpentrada "
    Sql = Sql & " where tmpclasific.numalbar=tmpentrada.numalbar "
    longitud = TotalRegistros(Sql)
    
    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0
    
    
    Sql = "select *, tmpentrada.kilosnet as kilosent from tmpclasific, tmpentrada "
    Sql = Sql & " where tmpclasific.numalbar=tmpentrada.numalbar "
    Sql = Sql & " order by tmpclasific.numalbar, tmpclasific.codcalir"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        AlbarAnt = DBLet(Rs!numalbar, "N")
        KilosNetAnt = DBLet(Rs!Kilosent, "N")
        VarieAnt = DBLet(Rs!CodVarie, "N")
        CalidAnt = DBLet(Rs!codcalir, "N")
    End If
        
    KilosAlbar = 0
    While Not Rs.EOF
        
        Me.pb2.Value = Me.pb2.Value + 1
        lblProgres(3).Caption = "Albarán " & DBLet(Rs!numalbar, "N") & " Variedad " & DBLet(Rs!CodVarie, "N") & " Calidad " & DBLet(Rs!codcalir, "N")
        Me.Refresh
        
        Kilos = Round2(DBLet(Rs!Kilosent, "N") * DBLet(Rs!porcenta, "N") / 100, 0)
        
        If AlbarAnt <> DBLet(Rs!numalbar, "N") Then
            If KilosNetAnt <> KilosAlbar Then
                Sql3 = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNetAnt - KilosAlbar, "N")
                Sql3 = Sql3 & " where numnotac = " & DBSet(AlbarAnt, "N")
                Sql3 = Sql3 & " and codvarie = " & DBSet(VarieAnt, "N")
                Sql3 = Sql3 & " and codcalid = " & DBSet(CalidAnt, "N")
            
                conn.Execute Sql3
            End If
        
            KilosAlbar = Kilos
            KilosNetAnt = DBLet(Rs!Kilosent, "N")
            
            AlbarAnt = DBLet(Rs!numalbar, "N")
        Else
            KilosAlbar = KilosAlbar + Kilos
        End If
    
        VarieAnt = DBLet(Rs!CodVarie, "N")
        CalidAnt = DBLet(Rs!codcalir, "N")
        
        
        Sql = "insert into rclasifica_clasif (numnotac,codvarie,codcalid,muestra,kilosnet) values"
        Sql = Sql & "(" & DBSet(Rs!numalbar, "N") & ","
        Sql = Sql & DBSet(Rs!CodVarie, "N") & ","
        Sql = Sql & DBSet(Rs!codcalir, "N") & ","
        Sql = Sql & DBSet(Rs!porcenta, "N") & ","
        Sql = Sql & DBSet(Kilos, "N") & ")"
        
        
        conn.Execute Sql
        
        Rs.MoveNext
    Wend
    
    ' si la clasificacion es diferente actualizamos en la ultima calidad
    If KilosNetAnt <> KilosAlbar Then
        Sql3 = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNetAnt - KilosAlbar, "N")
        Sql3 = Sql3 & " where numnotac = " & DBSet(AlbarAnt, "N")
        Sql3 = Sql3 & " and codvarie = " & DBSet(VarieAnt, "N")
        Sql3 = Sql3 & " and codcalid = " & DBSet(CalidAnt, "N")
    
        conn.Execute Sql3
    End If
    
    Set Rs = Nothing
    
    Sql = "select rclasifica.* from rclasifica, tmpentrada where rclasifica.numnotac = tmpentrada.numalbar "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not ActualizarTransporte(Rs, cadMen) Then
            cadMen = "Actualizando gastos de transporte" & cadMen
            MsgBox cadMen, vbExclamation
            Set Rs = Nothing
            
            pb2.visible = False
            lblProgres(2).Caption = ""
            lblProgres(3).Caption = ""
        
            CargarClasificacion = False
            Exit Function
        End If
    End If
    
    Set Rs = Nothing
    
    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    CargarClasificacion = True
    Exit Function
    
eCargarClasificacion:
    MuestraError Err.Number, "Cargar clasificación", Err.Description
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
        cadMen = "La carpeta de los ficheros de traza " & vParamAplic.PathTraza & " de parámetros no existe. Revise."
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


Private Function ActualizarTransporte(Rs As ADODB.Recordset, cadErr As String) As Boolean
Dim Sql1 As String
Dim Rs2 As ADODB.Recordset
Dim KilosDestrio As Currency
Dim Precio As Currency
Dim Transporte As Currency
Dim Kilos As Currency


    On Error GoTo eActualizarTransporte

    If Not Rs.EOF Then Rs.MoveFirst
    While Not Rs.EOF
        Sql1 = "select imptrans from rportespobla, rpartida, rcampos, variedades "
        Sql1 = Sql1 & " where rpartida.codparti = rcampos.codparti and "
        Sql1 = Sql1 & " variedades.codprodu = rportespobla.codprodu and "
        Sql1 = Sql1 & " rpartida.codpobla = rportespobla.codpobla and "
        Sql1 = Sql1 & " variedades.codvarie = " & DBSet(Rs!CodVarie, "N") & " and "
        Sql1 = Sql1 & " rcampos.codcampo = " & DBSet(Rs!codCampo, "N") & " and "
        Sql1 = Sql1 & " rcampos.codvarie = variedades.codvarie "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Precio = 0
        If Not Rs2.EOF Then
            Precio = DBLet(Rs2.Fields(0).Value, "N")
        End If
        
        Set Rs2 = Nothing
        
        ' cogemos los kilos de la clasificacion que sean de destrio
        Sql1 = "select kilosnet from rclasifica_clasif, rcalidad where numnotac = " & DBSet(Rs!numnotac, "N")
        Sql1 = Sql1 & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
        Sql1 = Sql1 & " and rclasifica_clasif.codcalid = rcalidad.codcalid  "
        Sql1 = Sql1 & " and rcalidad.tipcalid = 1 "
        KilosDestrio = DevuelveValor(Sql1)
        
        
        ' los gastos de transporte se calculan sobre los kilosnetos - los de destrio
        Kilos = DBLet(Rs!KilosNet, "N") - KilosDestrio
        Transporte = Round2(Kilos * Precio, 2)
        
        Sql1 = "update rclasifica set imptrans = " & DBSet(Transporte, "N")
        Sql1 = Sql1 & " where numnotac = " & DBSet(Rs!numnotac, "N")
        conn.Execute Sql1
        
        Rs.MoveNext
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
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim vSocio As CSocio
Dim b As Boolean
Dim Nregs As Long
Dim total As Variant

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
    
    Set Rs = Nothing
    
    'Imprimimos las lineas
    Aux = "select  rfactsoc.* "
    Aux = Aux & " from " & cTabla
    Aux = Aux & " where " & vWhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        'No hayningun registro
    Else
        b = True
        Regs = 0
        While Not Rs.EOF And b
            Regs = Regs + 1

            b = LineaTraspasoCoop(NFic, txtcodigo(45).Text, Rs)
            
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
            
    Close (NFic)
    
    If Regs > 0 And b Then GeneraFicheroTraspasoCoop = True
    Exit Function
    
EGen:
    Set Rs = Nothing
    Close (NFic)
    MuestraError Err.Number, Err.Description
End Function


Private Function GeneraFicheroTraspasoROPAS(pTabla As String, pWhere As String, pTabla1 As String, pWhere1 As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim vSocio As CSocio
Dim b As Boolean
Dim Nregs As Long
Dim total As Variant

Dim cTabla As String
Dim vWhere As String

Dim Lin As Integer

Dim AntSocio As Long
Dim AntPoligono As Long
Dim AntParcela As Long

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
    
    Open App.Path & "\socios.csv" For Output As #NFic
    
    Set Rs = Nothing
    
    'Imprimimos las lineas
    Aux = "select  rsocios.*, rsocios_seccion.* "
    Aux = Aux & " from " & cTabla
    If vWhere <> "" Then
        Aux = Aux & " where " & vWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        'No hayningun registro
    Else
        b = True
        Regs = 0
        While Not Rs.EOF And b
            Regs = Regs + 1

            b = LineaTraspasoSocioROPAS(NFic, Rs)
            
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
            
    Close (NFic)
    
    ' traspaso de campos de seccion horto
    If b Then
    
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
        
        Open App.Path & "\parcelas.csv" For Output As #NFic
        
        Set Rs = Nothing
        
        Aux = "select rcampos.codsocio, rcampos.codvarie, rsocios.nifsocio, rcampos.poligono,  "
        Aux = Aux & " rcampos.parcela, rcampos.subparce, rcampos.codparti, rcampos.supsigpa, "
        Aux = Aux & " rcampos.recintos, rcampos.supcoope, rcampos.canaforo, rcampos.fecaltas, "
        Aux = Aux & " rcampos.fecbajas, rcampos.supcatas, rsocios_seccion.fecalta "
        Aux = Aux & " from " & cTabla
        If vWhere <> "" Then
            Aux = Aux & " where " & vWhere
        End If
        Aux = Aux & " order by rcampos.codsocio, rcampos.poligono, rcampos.parcela, "
        Aux = Aux & " rcampos.subparce, rcampos.recintos, rcampos.codvarie"
        
        Set Rs = New ADODB.Recordset
        Rs.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Rs.EOF Then
            'No hayningun registro
        Else
            b = True
            Regs = 0
            Lin = 0
            If Not Rs.EOF Then
                AntSocio = DBLet(Rs!Codsocio, "N")
                AntPoligono = DBLet(Rs!poligono, "N")
                AntParcela = DBLet(Rs!parcela, "N")
            End If
            While Not Rs.EOF And b
                Regs = Regs + 1
    
                If AntSocio <> Rs!Codsocio Or AntPoligono <> Rs!poligono Or AntParcela <> Rs!parcela Then
                    Lin = 0
                End If
                Lin = Lin + 1
    
                b = LineaTraspasoCampoROPAS(NFic, Rs, Lin)
                
                Rs.MoveNext
            Wend
        End If
        Rs.Close
        Set Rs = Nothing
                
        Close (NFic)
        
    End If
    
    If Regs > 0 And b Then GeneraFicheroTraspasoROPAS = True
    Exit Function
    
EGen:
    Set Rs = Nothing
    Close (NFic)
    MuestraError Err.Number, Err.Description
End Function






Private Function LineaTraspasoCoop(NFich As Integer, Coop As String, ByRef Rs As ADODB.Recordset) As Boolean
Dim cad As String
Dim Areas As Long
Dim Tipo As Integer
Dim Sql As String
Dim vSocio As CSocio
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

Dim nifsocio As String
Dim Kilos As Long
Dim vPorcGasto As String
Dim vImporte As Currency
Dim Gastos As Currency



    On Error GoTo eLineaTraspasoCoop

    LineaTraspasoCoop = False

    cad = ""
    
    Sql = "select count(*) from rfactsoc_variedad where codtipom = " & DBSet(Rs!CodTipom, "T")
    Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N") & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
    
    If TotalRegistros(Sql) > 1 Then
        Producto = "00"
        Variedad = "00"
        NomVar = "Varias Var."
    Else
        Sql = "select rfactsoc_variedad.codvarie  from rfactsoc_variedad where codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N") & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        
        codVar = DevuelveValor(Sql)
        
        Producto = Mid(Format(codVar, "0000"), 1, 2)
        Variedad = Mid(Format(codVar, "0000"), 3, 2)
        
        NomVar = DevuelveValor("select nomvarie from variedades where codvarie = " & DBSet(codVar, "N"))
    End If
    
    
    If CInt(Coop) = 1 Or CInt(Coop) = 3 Or CInt(Coop) = 4 Then
        Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T"))
        Select Case Tipo
            Case 1, 3 'anticipo
                cad = "0|"
            Case 2, 4 'liquidacion
                cad = "1|"
            
        End Select
'        Producto = DevuelveValor("select codprodu from variedades where codvarie = " & DBSet(Rs!CodVarie, "N"))
        nifsocio = DevuelveValor("select nifsocio from rsocios where codsocio =" & DBSet(Rs!Codsocio, "N"))
        
        Sql = "select sum(kilosnet) from rfactsoc_variedad where codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N") & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        Kilos = DevuelveValor(Sql)
        
        If CInt(Coop) = 3 Or CInt(Coop) = 4 Then
            cad = cad & Format(DBLet(Rs!numfactu, "N"), "000000") & "|"
            cad = cad & Format(DBLet(Rs!fecfactu, "F"), "yymmdd") & "|"
            cad = cad & Format(DBLet(Rs!Codsocio, "N"), "0000") & "|"
            cad = cad & Format(DBLet(Producto, "N"), "00") & "|"
            cad = cad & Format(DBLet(Variedad, "N"), "00") & "|"
            cad = cad & RellenaABlancos(NomVar, True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(Rs!baseimpo, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(Rs!imporiva, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(Rs!TotalFac, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(Rs!ImpReten, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(Rs!ImpApor, "N"), "#######0.00"), True, 11) & "|"
            
        Else
            cad = cad & Format(DBLet(Rs!numfactu, "N"), "000000") & "|"
            cad = cad & Format(DBLet(Rs!fecfactu, "F"), "yymmdd") & "|"
            cad = cad & Format(DBLet(Rs!Codsocio, "N"), "000000") & "|"
            cad = cad & Format(DBLet(Producto, "N"), "00") & "|"
            cad = cad & Format(DBLet(Variedad, "N"), "00") & "|"
            cad = cad & RellenaABlancos(NomVar, True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(Rs!baseimpo, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(Rs!imporiva, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(Rs!TotalFac, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(Rs!ImpReten, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(Format(DBLet(Rs!ImpApor, "N"), "#######0.00"), True, 11) & "|"
            cad = cad & RellenaABlancos(nifsocio, True, 9) & "|"
            cad = cad & Format(Kilos, "00000000") & "|"
            
        End If
    Else
        cad = cad & Format(DBLet(Rs!numfactu, "N"), "0000000")
        cad = cad & Format(DBLet(Rs!Codsocio, "N"), "0000000")
        cad = cad & Format(DBLet(Rs!fecfactu, "F"), "yymmdd")
        cad = cad & RellenaABlancos(NomVar, True, 11)
        cad = cad & RellenaABlancos(Format(Abs(DBLet(Rs!baseimpo, "N")), "00000.00"), True, 8)
        
        If DBLet(Rs!baseimpo, "N") < 0 Then
            cad = cad & "-"
        Else
            cad = cad & "+"
        End If
        
        vPorcIva = Round2(DBLet(Rs!Porc_Iva, "N") * 100, 0)
        
        cad = cad & Format(vPorcIva, "0000")
        cad = cad & "0000"
        cad = cad & Format(Abs(DBLet(Rs!imporiva, "N")), "000.00")
        
        If DBLet(Rs!imporiva, "N") < 0 Then
            cad = cad & "-"
        Else
            cad = cad & "+"
        End If
        
        ' total factura
        cad = cad & Format(Abs(DBLet(Rs!TotalFac, "N")), "00000.00")
        
        If DBLet(Rs!TotalFac, "N") < 0 Then
            cad = cad & "-"
        Else
            cad = cad & "+"
        End If
        
        cad = cad & "00000000"
        
        ' base de retencion
        If DBLet(Rs!BaseReten, "N") = 0 Then
            cad = cad & "00000000+"
        Else
            If DBLet(Rs!BaseReten, "N") < 0 Then
                cad = cad & Format(Abs(DBLet(Rs!BaseReten, "N")), "00000.00") & "-"
            Else
                cad = cad & Format(Abs(DBLet(Rs!BaseReten, "N")), "00000.00") & "+"
            End If
        End If
        
        ' porcentaje de retencion
        cad = cad & Format(Round2(DBLet(Rs!porc_ret, "N") * 100, 0), "0000")
        If DBLet(Rs!ImpReten, "N") >= 0 Then
            cad = cad & Format(DBLet(Rs!ImpReten, "N"), "000.00") & "+"
        Else
            cad = cad & Format(Abs(DBLet(Rs!ImpReten, "N")), "000.00") & "-"
        End If
        
        ' gastos de la cooperativa
        CoopSoc = DevuelveValor("select codcoope from rsocios where codsocio = " & DBLet(Rs!Codsocio, "N"))
        
        vPorcGasto = ""
        vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", CStr(CoopSoc), "N")
        If vPorcGasto = "" Then vPorcGasto = "0"
        
        Sql = "select sum(imporvar) from rfacsoc_variedad where codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
        Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        
        vImporte = DevuelveValor(Sql)
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
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim Kilos1 As Long
Dim Kilos2 As Long
Dim Kilos3 As Long
Dim Kilos4 As Long
Dim Kilos5 As Long
Dim Kilos6 As Long
Dim Kilos7 As Long
Dim Kilos8 As Long
Dim Kilos9 As Long
Dim vCond As String
Dim vCond2 As String
Dim vResult As String


    On Error GoTo eCargarTemporal
    
    Screen.MousePointer = vbHourglass
    
    CargarTemporal6 = False

    Sql2 = "delete from tmpclasifica where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql = "Select variedades.codvarie FROM variedades "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " where " & cWhere
    End If

    vCond = ""
    vCond2 = ""
    
    If txtcodigo(54).Text <> "" Then vCond = vCond & " and rhisfruta.codsocio >= " & DBSet(txtcodigo(54).Text, "N")
    If txtcodigo(55).Text <> "" Then vCond = vCond & " and rhisfruta.codsocio <= " & DBSet(txtcodigo(55).Text, "N")
    
    If txtcodigo(52).Text <> "" Then vCond = vCond & " and rhisfruta.fecalbar >= " & DBSet(txtcodigo(52).Text, "F")
    If txtcodigo(53).Text <> "" Then vCond = vCond & " and rhisfruta.fecalbar <= " & DBSet(txtcodigo(53).Text, "F")
    
    If Check7.Value = 1 Then
        If txtcodigo(54).Text <> "" Then vCond2 = vCond2 & " and rclasifica.codsocio >= " & DBSet(txtcodigo(54).Text, "N")
        If txtcodigo(55).Text <> "" Then vCond2 = vCond2 & " and rclasifica.codsocio <= " & DBSet(txtcodigo(55).Text, "N")
        
        If txtcodigo(52).Text <> "" Then vCond2 = vCond2 & " and rclasifica.fechaent >= " & DBSet(txtcodigo(52).Text, "F")
        If txtcodigo(53).Text <> "" Then vCond2 = vCond2 & " and rclasifica.fechaent <= " & DBSet(txtcodigo(53).Text, "F")
    End If
    
    vResult = ""
    
    
    ' obtenemos los kilos de cada variedad con las condiciones
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        
        'KILOS PRODUCCION NORMAL COOPERATIVA --> KILOS1
        Sql2 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif, rhisfruta "
        Sql2 = Sql2 & " where rhisfruta.codvarie = " & DBSet(Rs!CodVarie, "N")
        Sql2 = Sql2 & " and rhisfruta.tipoentr <> 2 " ' produccion normal
        Sql2 = Sql2 & " and rhisfruta.recolect = 0 " ' recolectado cooperativa
        Sql2 = Sql2 & " and rhisfruta.numalbar = rhisfruta_clasif.numalbar "
        If vCond <> "" Then
            Sql2 = Sql2 & vCond
        End If
        
        Kilos1 = DevuelveValor(Sql2)
        
        If Check7.Value Then
            Sql2 = "select sum(rclasifica.kilosnet) from rclasifica "
            Sql2 = Sql2 & " where rclasifica.codvarie = " & DBSet(Rs!CodVarie, "N")
            Sql2 = Sql2 & " and rclasifica.tipoentr <> 2 " ' produccion normal
            Sql2 = Sql2 & " and rclasifica.recolect = 0 "  ' recolectado cooperativa
            If vCond2 <> "" Then
                Sql2 = Sql2 & vCond2
            End If
        
            Kilos1 = Kilos1 + DevuelveValor(Sql2)
        End If
        
        
        'KILOS PRODUCCION NORMAL SOCIO --> KILOS2
        Sql2 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif, rhisfruta "
        Sql2 = Sql2 & " where rhisfruta.codvarie = " & DBSet(Rs!CodVarie, "N")
        Sql2 = Sql2 & " and rhisfruta.tipoentr <> 2 " ' produccion normal
        Sql2 = Sql2 & " and rhisfruta.recolect = 1 " ' recolectado socio
        Sql2 = Sql2 & " and rhisfruta.numalbar = rhisfruta_clasif.numalbar "
        If vCond <> "" Then
            Sql2 = Sql2 & vCond
        End If
        
        Kilos2 = DevuelveValor(Sql2)
    
        If Check7.Value Then
            Sql2 = "select sum(rclasifica.kilosnet) from rclasifica "
            Sql2 = Sql2 & " where rclasifica.codvarie = " & DBSet(Rs!CodVarie, "N")
            Sql2 = Sql2 & " and rclasifica.tipoentr <> 2 " ' produccion normal
            Sql2 = Sql2 & " and rclasifica.recolect = 1 "  ' recolectado socio
            If vCond2 <> "" Then
                Sql2 = Sql2 & vCond2
            End If
        
            Kilos2 = Kilos2 + DevuelveValor(Sql2)
        End If
    
    
        ' KILOS PRODUCCION INTEGRADA COOPERATIVA --> KILOS3
        Sql2 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif, rhisfruta "
        Sql2 = Sql2 & " where rhisfruta.codvarie = " & DBSet(Rs!CodVarie, "N")
        Sql2 = Sql2 & " and rhisfruta.tipoentr = 2 " ' produccion integrada
        Sql2 = Sql2 & " and rhisfruta.recolect = 0 " ' recolectado cooperativa
        Sql2 = Sql2 & " and rhisfruta.numalbar = rhisfruta_clasif.numalbar "
        If vCond <> "" Then
            Sql2 = Sql2 & vCond
        End If
        
        Kilos3 = DevuelveValor(Sql2)
        
        If Check7.Value Then
            Sql2 = "select sum(rclasifica.kilosnet) from rclasifica "
            Sql2 = Sql2 & " where rclasifica.codvarie = " & DBSet(Rs!CodVarie, "N")
            Sql2 = Sql2 & " and rclasifica.tipoentr = 2 " ' produccion integrada
            Sql2 = Sql2 & " and rclasifica.recolect = 0 "  ' recolectado cooperativa
            If vCond2 <> "" Then
                Sql2 = Sql2 & vCond2
            End If
        
            Kilos3 = Kilos3 + DevuelveValor(Sql2)
        End If
        
        ' KILOS PRODUCCION INTEGRADA SOCIO --> KILOS4
        Sql2 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif, rhisfruta "
        Sql2 = Sql2 & " where rhisfruta.codvarie = " & DBSet(Rs!CodVarie, "N")
        Sql2 = Sql2 & " and rhisfruta.tipoentr = 2 " ' produccion integrada
        Sql2 = Sql2 & " and rhisfruta.recolect = 1 " ' recolectado socio
        Sql2 = Sql2 & " and rhisfruta.numalbar = rhisfruta_clasif.numalbar "
        If vCond <> "" Then
            Sql2 = Sql2 & vCond
        End If
        
        Kilos4 = DevuelveValor(Sql2)
        
        If Check7.Value Then
            Sql2 = "select sum(rclasifica.kilosnet) from rclasifica "
            Sql2 = Sql2 & " where rclasifica.codvarie = " & DBSet(Rs!CodVarie, "N")
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
    
    
        vResult = vResult & "(" & vUsu.Codigo & "," & DBSet(Rs!CodVarie, "N") & ","
        vResult = vResult & DBSet(Kilos2, "N", "S") & "," & DBSet(Kilos1, "N", "S") & ","
        vResult = vResult & DBSet(Kilos5, "N", "S") & "," & DBSet(Kilos4, "N", "S") & ","
        vResult = vResult & DBSet(Kilos3, "N", "S") & "," & DBSet(Kilos6, "N", "S") & ","
        vResult = vResult & DBSet(Kilos7, "N", "S") & "," & DBSet(Kilos8, "N", "S") & ","
        vResult = vResult & DBSet(Kilos9, "N", "S") & "),"
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If vResult <> "" Then
        Sql2 = "insert into tmpclasifica (codusu,codvarie,cal1,cal2,cal3,cal4,cal5,"
        Sql2 = Sql2 & "cal6,cal7,cal8,cal9) values "
        
        Sql2 = Sql2 & Mid(vResult, 1, Len(vResult) - 1)  ' quitamos la ultima coma
    End If

    conn.Execute Sql2
    
    ' borramos aquellos registros que no tienen kilos de ningun tipo
    Sql2 = "delete from tmpclasifica where cal1 is null and cal2 is null and cal3 is null and "
    Sql2 = Sql2 & " cal4 is null and cal5 is null and cal6 is null and cal7 is null and "
    Sql2 = Sql2 & " cal8 is null and cal9 is null and codusu = " & DBSet(vUsu.Codigo, "N")
    
    conn.Execute Sql2
    
    CargarTemporal6 = True
    Screen.MousePointer = vbDefault

    Exit Function
    
eCargarTemporal:
    Screen.MousePointer = vbDefault
    MuestraError "Cargando temporal", Err.Description
End Function


'Lineas traspaso ropas


Private Function LineaTraspasoSocioROPAS(NFich As Integer, ByRef Rs As ADODB.Recordset) As Boolean
Dim cad As String
Dim Areas As Long
Dim Tipo As Integer
Dim Sql As String
Dim vSocio As CSocio
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

Dim nifsocio As String
Dim Kilos As Long
Dim vPorcGasto As String
Dim vImporte As Currency
Dim Gastos As Currency



    On Error GoTo eLineaTraspasoSocioROPAS

    LineaTraspasoSocioROPAS = False

    cad = ""
    cad = cad & Format(txtcodigo(62).Text, "0000") & ";"
    cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 12) & ";"
    cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 12) & ";"
    cad = cad & Format(Rs!Codsocio, "######") & ";"
    cad = cad & RellenaABlancos(Rs!nomsocio, True, 60) & ";ES;"
    
    If DBLet(Rs!TipoIRPF, "N") <> 2 Then
        cad = cad & "P;"
    Else
        cad = cad & "J;"
    End If
    
    cad = cad & Format(DBLet(Rs!FecAlta, "F"), "dd/mm/yyyy")
    cad = cad & Format(DBLet(Rs!fecbaja, "F"), "dd/mm/yyyy")

    Print #NFich, cad
    
    LineaTraspasoSocioROPAS = True
    Exit Function
    
eLineaTraspasoSocioROPAS:
    MuestraError Err.Number, "Carga Linea de Traspaso Socios ROPAS", Err.Description
End Function



Private Function LineaTraspasoCampoROPAS(NFich As Integer, ByRef Rs As ADODB.Recordset, Lin As Integer) As Boolean
Dim cad As String
Dim Areas As Long
Dim Tipo As Integer
Dim Sql As String
Dim vSocio As CSocio
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

Dim nifsocio As String
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

    On Error GoTo eLineaTraspasoCampoROPAS

    LineaTraspasoCampoROPAS = False

    cad = ""
    cad = cad & Format(txtcodigo(62).Text, "0000") & ";"
    cad = cad & RellenaABlancos(vParam.CifEmpresa, True, 12) & ";"
    cad = cad & RellenaABlancos(Rs!nifsocio, True, 12) & ";ES;R;"
    cad = cad & Space(27) & ";"
    
    Pobla = ""
    Pobla = DevuelveValor("select codpobla from rpartida where codparti = " & DBSet(Rs!Codparti, "N"))
    
    cad = cad & Mid(Pobla, 1, 2) & ";"
    
    CodSigPa = ""
    CodSigPa = DevuelveValor("select codsigpa from rpueblos where codpobla = " & DBSet(Pobla, "T"))
    
    cad = cad & Format(CodSigPa, "###") & ";"
    cad = cad & "000;"
    cad = cad & "00;"
    cad = cad & Format(Rs!poligono, "###") & ";"
    cad = cad & Format(Rs!parcela, "#####") & ";"
    cad = cad & Format(Rs!recintos, "#####") & ";"
    cad = cad & RellenaABlancos(Rs!subparce, True, 2) & ";"
    
    HectaSig = 0 '  SUPERFICIE TOTAL PARCELA
    
    Sql = "select sum(supcatas) from rcampos, rpartida where poligono = " & DBSet(Rs!poligono, "N")
    Sql = Sql & " and parcelas = " & DBSet(Rs!parcela, "N")
    Sql = Sql & " and rcampos.fecbajas is null "
    Sql = Sql & " and rpartida.codpobla = " & DBSet(Pobla, "T")
    Sql = Sql & " and rcampos.codparti = rpartida.codparti "
    
    HectaSig = DevuelveValor(Sql)
    
    cad = cad & Format(HectaSig, "##0.0000") & ";"
    cad = cad & Format(Rs!supsigpa, "##0.0000") & ";"
    cad = cad & Format(Rs!supcatas, "##0.0000") & ";"
    
    FecAlta = Rs!fecaltas
    If Rs!FecAlta > Rs!fecaltas Then ' fecha alta socio > fecha alta campo
        FecAlta = Rs!FecAlta
    End If
    
    cad = cad & Format(FecAlta, "dd/mm/yyyy") & ";"
    cad = cad & Format(Rs!fecbajas, "dd/mm/yyyy") & ";"
    cad = cad & Format(Lin, "#") & ";"  ' contador de subparcelas
    
    
    CodConse = 0
    CodConse = DevuelveValor("select codconse from variedades where codvarie = " & DBSet(Rs!CodVarie, "N"))
    
    cad = cad & RellenaABlancos(CStr(CodConse), True, 6) & ";"
    
    Super = DBLet(Rs!supcoope, "N")
    If DBLet(Rs!supcoope, "N") > DBLet(Rs!supsigpa, "N") Then
        Super = DBLet(Rs!supsigpa, "N")
    End If
    
    cad = cad & Format(Super, "000.0000") & ";"
    
    If CanAfo = 0 Then Let CanAfo = 10
    CanAfo = Round2(Rs!canaforo / 1000, 2) 'En toneladas
    
    cad = cad & Format(CanAfo, "0000.00") & ";"
    
    Print #NFich, cad
    
    LineaTraspasoCampoROPAS = True
    Exit Function
    
eLineaTraspasoCampoROPAS:
    MuestraError Err.Number, "Carga Linea de Traspaso Campos ROPAS", Err.Description
End Function


Private Function CopiarFicheroROPAS() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFicheroROPAS = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.cd1.DefaultExt = "csv"
    
    cd1.Filter = "Archivos csv|csv|"
    cd1.FilterIndex = 1
    cd1.FileName = "socios.csv"
    
    Me.cd1.ShowSave
    
    If cd1.FileName <> "" Then
        FileCopy App.Path & "\socios.csv", cd1.FileName
        cd1.FileName = "parcelas.csv"
        FileCopy App.Path & "\parcelas.csv", cd1.FileName
    End If
    
    CopiarFicheroROPAS = True
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
        Case 19 ' fichero de agriweb
            If b Then
                If txtcodigo(27).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente el año del ejercicios.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(26)
                End If
            End If
            If b Then
                If txtcodigo(28).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente el CIF de la industria transformadora.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(28)
                End If
            End If
            If b Then
                If txtcodigo(29).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente los kilos contratados.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(29)
                End If
            End If
            If b Then
                If txtcodigo(30).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de formalización.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(30)
                End If
            End If
            If b Then
                If txtcodigo(31).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Superficie Total de Contrato.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(31)
                End If
            End If
            If b Then
                If txtcodigo(32).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente el Precio Estipulado de Compra.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(32)
                End If
            End If
            
            
    End Select
    DatosOk = b

End Function

''[Monica]25/09/2009: han cambiado el CALIBRADOR GRANDE de catadau. Cada fichero se corresponde con
''                    una nota de entrada.
''        19/10/2009: el calibrador pequeño no se corresponde con el agre1104
'' Proceso de traspaso para CATADAU
''
'Private Function ProcesarFicheroCatadauCGrande(NF As Long, cad As String) As Boolean
'Dim b As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim Mens As String
'Dim numlinea As Long
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim RSaux As ADODB.Recordset
'
'Dim Codsoc As String
'Dim Codcam As String
'Dim codpro As String
'Dim codVar As String
'Dim Observ As String
'Dim Notaca As String
'Dim Kilone As String
'
'Dim Destri As String
'Dim Podrid As String
'Dim Pequen As String
'Dim Muestra As String
'
'Dim NGrupos As String
'Dim Nombre1 As String
'
'
'
'Dim I As Integer
'Dim Situacion As Byte
'
'Dim NomCal As Dictionary
'Dim KilCal As Dictionary
'
'Dim SQLaux As String
'Dim Nregs As Integer
'
'Dim Nsep As Integer
'
'Dim SeInserta As Boolean
'Dim KilosTot As Currency
'Dim cantidad As Long
'Dim Kilos As Currency
'
'Dim HoraIni As String
'Dim HoraFin As String
'
'Dim FechaEnt As String
'Dim UltimaLinea As Boolean
'Dim NroCalidad As Integer
'
'Dim Porcen As String
'Dim KilosMuestreo As String
'
'
'    On Error GoTo eProcesarFicheroCatadau
'
'    ProcesarFicheroCatadauCGrande = False
'
'    Codsoc = 0
'    Codcam = 0
'    codpro = 0
'    codVar = 0
'    Observ = ""
'    Notaca = 0
'    Kilone = 0
'    KilosTot = 0
'
'    Destri = 0
'    Podrid = 0
'    Pequen = 0
'
'    I = 0
'
'    ' inicializamos las variables
'    Set NomCal = New Dictionary
'    Set KilCal = New Dictionary
'
'    Notaca = RecuperaValorNew(cad, ";", 1)
'
'    Sql = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Rs.EOF Then
'        Observ = "NOTA NO EXISTE"
'        Situacion = 2
'    End If
'
'    b = True
'    UltimaLinea = False
'    NroCalidad = 0
'    While Not EOF(NF) And Not UltimaLinea
'        I = I + 1
'
'        Me.pb1.Value = Me.pb1.Value + Len(cad)
'        lblProgres(1).Caption = "Linea " & I
'        Me.Refresh
'
'        Nsep = NumeroSubcadenasInStr(cad, ";")
'
'        If Nsep = 14 Then ' estamos en una calidad
'            NroCalidad = NroCalidad + 1
'
'            Nombre1 = RecuperaValorNew(cad, ";", 4)
'            Kilone = RecuperaValorNew(cad, ";", 7)
'
'            Kilos = Round2(CCur(Kilone) / 1000, 2)
'
'            cantidad = RecuperaValorNew(cad, ";", 8)
'            KilosTot = KilosTot + Kilos
'
'            If Situacion <> 2 Then
'                ' si hay nota asociada busco los datos
'                Sql = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(Rs!CodVarie, "N")
'                Sql = Sql & " and nomcalibrador1 = " & DBSet(Nombre1, "T")
'
'                Set RS1 = New ADODB.Recordset
'                RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If RS1.EOF Then
'                    Observ = "NO EXIS.CAL"
'                    Situacion = 1
'                Else
'                    NomCal(I) = DBLet(RS1!codcalid, "N")
'                    KilCal(I) = Kilos
'                End If
'                Set RS1 = Nothing
'
'            End If
'        End If
'
'        If Nsep = 15 Then ' estamos en la ultima linea
'            HoraIni = RecuperaValorNew(cad, ";", 9)
'            HoraFin = RecuperaValorNew(cad, ";", 10)
'            FechaEnt = RecuperaValorNew(cad, ";", 11)
'
'            UltimaLinea = True
'        End If
'
'        Line Input #NF, cad
'    Wend
'
'    Close #NF
'
''    If DBLet(Rs.Fields(0).Value, "N") <> KilosTot Then
''        Observ = "K.NETOS DIF."
''        Situacion = 4
''    End If
'
'
'    Sql = "select count(*) from rclasifauto where numnotac = " & Notaca
'
'    SeInserta = (TotalRegistros(Sql) = 0)
'
'    If SeInserta Then
'        If Situacion = 2 Then
'            ' si no hay nota asociada no puedo meter la clasificacion
'            Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'            Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'            Sql = Sql & "`observac`,`situacion`) values ("
'            Sql = Sql & DBSet(Notaca, "N") & ","
'            Sql = Sql & DBSet(0, "N") & ","
'            Sql = Sql & DBSet(0, "N") & ","
'            Sql = Sql & DBSet(0, "N") & ","
'            Sql = Sql & DBSet(KilosTot, "N") & ","
'            Sql = Sql & DBSet(Destri, "N") & ","
'            Sql = Sql & DBSet(Podrid, "N") & ","
'            Sql = Sql & DBSet(Pequen, "N") & ","
'            Sql = Sql & DBSet(Observ, "T") & ","
'            Sql = Sql & DBSet(Situacion, "N") & ")"
'
'        Else
'            ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
'            ' tabla: rclasifauto
'            Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'            Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'            Sql = Sql & "`observac`,`situacion`) values ("
'            Sql = Sql & DBSet(Notaca, "N") & ","
'            Sql = Sql & DBSet(Rs!Codsocio, "N") & ","
'            Sql = Sql & DBSet(Rs!codCampo, "N") & ","
'            Sql = Sql & DBSet(Rs!CodVarie, "N") & ","
'            Sql = Sql & DBSet(Round2(KilosTot, 0), "N") & ","
'            Sql = Sql & DBSet(Destri, "N") & ","
'            Sql = Sql & DBSet(Podrid, "N") & ","
'            Sql = Sql & DBSet(Pequen, "N") & ","
'            Sql = Sql & DBSet(Observ, "T") & ","
'            Sql = Sql & DBSet(Situacion, "N") & ")"
'        End If
'        conn.Execute Sql
'
'        ' tabla: rclasifauto_clasif
'        Sql = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
'        Sql = Sql & " values "
'
'    End If
'
'    'solo si tenemos nota asociada metemos toda la clasificacion
'    If Situacion <> 2 Then
'
'        'borramos la tabla temporal
'        SQLaux = "delete from tmpcata"
'        conn.Execute SQLaux
'
'        ' cargamos la tabla temporal
'        For I = 1 To NroCalidad
'            If NomCal(I) <> "" Then
'                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(I), "N"))
'                If Nregs = 0 Then
'                    'insertamos en la temporal
'                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(I), "N")
'                    SQLaux = SQLaux & "," & DBSet(KilCal(I), "N") & ")"
'
'                    conn.Execute SQLaux
'                Else
'                    'actualizamos la temporal
'                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(I), "N")
'                    SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(I), "N")
'
'                    conn.Execute SQLaux
'                End If
'            End If
'        Next I
'
'        SQLaux = "select * from tmpcata order by codcalid"
'
'        Set RSaux = New ADODB.Recordset
'        RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        Sql2 = ""
'
'        While Not RSaux.EOF
'            If SeInserta Then
'                Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(Rs!CodVarie, "N") & ","
'                Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
'            Else
'                Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
'                Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
'                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'                Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
'
'                conn.Execute Sql2
'            End If
'
'            RSaux.MoveNext
'        Wend
'
'        Set RSaux = Nothing
'
'
'        If SeInserta Then
'            If Sql2 <> "" Then
'                Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
'            End If
'            Sql = Sql & Sql2
'            conn.Execute Sql
'        End If
'    End If ' si la situacion es distinta de 2
'
'
'    Set Rs = Nothing
'    Set NomCal = Nothing
'    Set KilCal = Nothing
'
'    ProcesarFicheroCatadauCGrande = True
'    Exit Function
'
'eProcesarFicheroCatadau:
'    If Err.Number <> 0 Then
'        ProcesarFicheroCatadauCGrande = False
'        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
'    End If
'End Function
'
''[Monica]19/10/2009: CALIBRADOR PEQUEÑO
'' ESTE NO SE CORRESPONDE CON AGRE1104 DE EUROAGRO
'Private Function ProcesarFicheroCatadauCPequeño() As Boolean
'Dim b As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim Mens As String
'Dim numlinea As Long
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim Rs2 As ADODB.Recordset
'Dim RSaux As ADODB.Recordset
'
'Dim Codsoc As String
'Dim Codcam As String
'Dim codpro As String
'Dim codVar As String
'Dim Observ As String
'Dim Notaca As String
'Dim Kilone As String
'
'Dim Destri As String
'Dim Podrid As String
'Dim Pequen As String
'Dim Muestra As String
'
'Dim NGrupos As String
'Dim Nombre1 As String
'
'Dim I As Integer
'Dim J As Integer
'Dim Situacion As Byte
'
'Dim NomCal As Dictionary
'Dim KilCal As Dictionary
'
'Dim SQLaux As String
'Dim Nregs As Integer
'
'Dim Nsep As Integer
'
'Dim SeInserta As Boolean
'Dim KilosTot As Currency
'Dim cantidad As Long
'Dim Kilos As Currency
'
'Dim HoraIni As String
'Dim HoraFin As String
'
'Dim FechaEnt As String
'Dim UltimaLinea As Boolean
'Dim NroCalidad As Integer
'
'Dim Porcen As String
'Dim KilosMuestreo As String
'
'
'    On Error GoTo eProcesarFicheroCatadauCPequeño
'
'    ProcesarFicheroCatadauCPequeño = False
'
'    Codsoc = 0
'    Codcam = 0
'    codpro = 0
'    codVar = 0
'    Observ = ""
'    Notaca = 0
'    Kilone = 0
'    KilosTot = 0
'
'    Destri = 0
'    Podrid = 0
'    Pequen = 0
'
'    I = 0
'
'    ' inicializamos las variables
'    Set NomCal = New Dictionary
'    Set KilCal = New Dictionary
'
'
'    Sql = "select * from tmpcalibrador "
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    Notaca = 0
'    If Not Rs.EOF Then
'        Notaca = DBLet(Rs.Fields(0).Value, "N")
'
'        If Notaca <> 0 Then
'            Sql2 = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
'            Set RS1 = New ADODB.Recordset
'            RS1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            If RS1.EOF Then
'                Observ = "NOTA NO EXISTE"
'                Situacion = 2
'            End If
'
'            b = True
'
'            While Not Rs.EOF
'                I = I + 1
'
'                Me.pb1.Value = Me.pb1.Value + 1
'                lblProgres(1).Caption = "Linea " & I
'                Me.Refresh
'
'                Nombre1 = DBLet(Rs!nomcalid, "T")
'                Destri = DBLet(Rs!Kilos3, "T")
'                Podrid = DBLet(Rs!Kilos2, "T")
'                'Pequen = DBLet(RS!Kilos4, "T")
''antes calculo de kilos segun porcentaje
''                Kilone = DBLet(RS!Kilos1, "T")
''                Porcen = DBLet(RS!porcen1, "T")
''                Kilos = Round2(CCur(Kilone) * CCur(Porcen) / 100, 2)
''                KilosTot = KilosTot + Kilos
''ahora me guardo el porcentaje
'                KilosTot = DBLet(Rs!Kilos1, "T")
'                Kilos = DBLet(Rs!porcen1, "T")
'
'                If Situacion <> 2 Then
'                    ' si hay nota asociada busco los datos
'                    Sql = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(RS1!CodVarie, "N")
'                    Sql = Sql & " and nomcalibrador2 = " & DBSet(Nombre1, "T")
'
'                    Set Rs2 = New ADODB.Recordset
'                    Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                    If Rs2.EOF Then
'                        Observ = "NO EXIS.CAL"
'                        Situacion = 1
'                    Else
'                        NomCal(I) = DBLet(Rs2!codcalid, "N")
'                        KilCal(I) = Kilos
'                    End If
'                    Set Rs2 = Nothing
'
'                End If
'
'                Rs.MoveNext
'            Wend
'
'            Sql = "select count(*) from rclasifauto where numnotac = " & Notaca
'
'            SeInserta = (TotalRegistros(Sql) = 0)
'
'            If SeInserta Then
'                If Situacion = 2 Then
'                    ' si no hay nota asociada no puedo meter la clasificacion
'                    Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'                    Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'                    Sql = Sql & "`observac`,`situacion`) values ("
'                    Sql = Sql & DBSet(Notaca, "N") & ","
'                    Sql = Sql & DBSet(0, "N") & ","
'                    Sql = Sql & DBSet(0, "N") & ","
'                    Sql = Sql & DBSet(0, "N") & ","
'                    Sql = Sql & DBSet(KilosTot, "N") & ","
'                    Sql = Sql & DBSet(Destri, "N") & ","
'                    Sql = Sql & DBSet(Podrid, "N") & ","
'                    Sql = Sql & DBSet(Pequen, "N") & ","
'                    Sql = Sql & DBSet(Observ, "T") & ","
'                    Sql = Sql & DBSet(Situacion, "N") & ")"
'
'                Else
'                    ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
'                    ' tabla: rclasifauto
'                    Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'                    Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'                    Sql = Sql & "`observac`,`situacion`) values ("
'                    Sql = Sql & DBSet(Notaca, "N") & ","
'                    Sql = Sql & DBSet(RS1!Codsocio, "N") & ","
'                    Sql = Sql & DBSet(RS1!codCampo, "N") & ","
'                    Sql = Sql & DBSet(RS1!CodVarie, "N") & ","
'                    Sql = Sql & DBSet(KilosTot, "N") & ","
'                    Sql = Sql & DBSet(Destri, "N") & ","
'                    Sql = Sql & DBSet(Podrid, "N") & ","
'                    Sql = Sql & DBSet(Pequen, "N") & ","
'                    Sql = Sql & DBSet(Observ, "T") & ","
'                    Sql = Sql & DBSet(Situacion, "N") & ")"
'                End If
'                conn.Execute Sql
'
'                ' tabla: rclasifauto_clasif
'                Sql = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
'                Sql = Sql & " values "
'
'            End If
'        End If
'        'solo si tenemos nota asociada metemos toda la clasificacion
'        If Situacion <> 2 Then
'            'borramos la tabla temporal
'            SQLaux = "delete from tmpcata"
'            conn.Execute SQLaux
'
'            ' cargamos la tabla temporal
'            For J = 1 To I
'                If NomCal(J) <> "" Then
'                    Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(J), "N"))
'                    If Nregs = 0 Then
'                        'insertamos en la temporal
'                        SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(J), "N")
'                        SQLaux = SQLaux & "," & DBSet(KilCal(J), "N") & ")"
'
'                        conn.Execute SQLaux
'                    Else
'                        'actualizamos la temporal
'                        SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(J), "N")
'                        SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(J), "N")
'
'                        conn.Execute SQLaux
'                    End If
'                End If
'            Next J
'
'            SQLaux = "select * from tmpcata order by codcalid"
'
'            Set RSaux = New ADODB.Recordset
'            RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            Sql2 = ""
'
'            While Not RSaux.EOF
'                If SeInserta Then
'                    Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(RS1!CodVarie, "N") & ","
'                    Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
'                Else
'                    Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
'                    Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
'                    Sql2 = Sql2 & " and codvarie = " & DBSet(RS1!CodVarie, "N")
'                    Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
'
'                    conn.Execute Sql2
'                End If
'
'                RSaux.MoveNext
'            Wend
'
'            Set RSaux = Nothing
'
'            If SeInserta Then
'                If Sql2 <> "" Then
'                    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
'                End If
'                Sql = Sql & Sql2
'                conn.Execute Sql
'            End If
'        End If ' si la situacion es distinta de 2
'
'        Set Rs = Nothing
'        Set RS1 = Nothing
'        Set NomCal = Nothing
'        Set KilCal = Nothing
'
'        ProcesarFicheroCatadauCPequeño = True
'        Exit Function
'
'    End If
'
''    Notaca = Mid(cad, 2, InStr(2, cad, "") + 1)
''
''    SQL = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
''    Set RS = New ADODB.Recordset
''    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''
''    If RS.EOF Then
''        Observ = "NOTA NO EXISTE"
''        Situacion = 2
''    End If
''
''    b = True
''    UltimaLinea = False
''    NroCalidad = 0
''    While Not EOF(NF) And Not UltimaLinea
''        I = I + 1
''
''        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
''        lblProgres(1).Caption = "Linea " & I
''        Me.Refresh
''
''        NroCalidad = NroCalidad + 1
''        Nombre1 = DevuelveNomCalidad(cad, 71)
'''        Nombre1 = Mid(cad, 71, InStr(55, cad, "export") + 10)
''        KilosMuestreo = Mid(cad, 44, 6)
''
''        Porcen = Mid(cad, 34, 5)
''
'''        Kilone = Round2(porcen * kilosmuestreo / 100, 2)
''
''        KilosTot = KilosTot + Kilone
''
''        If Situacion <> 2 Then
''            ' si hay nota asociada busco los datos
''            SQL = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(RS!CodVarie, "N")
''            SQL = SQL & " and nomcalibrador2 = " & DBSet(Nombre1, "T")
''
''            Set RS1 = New ADODB.Recordset
''            RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''
''            If RS1.EOF Then
''                Observ = "NO EXIS.CAL"
''                Situacion = 1
''            Else
''                NomCal(I) = DBLet(RS1!codcalid, "N")
''                KilCal(I) = Kilos
''            End If
''            Set RS1 = Nothing
''        End If
''
''        Line Input #NF, cad
''    Wend
''
''    Close #NF
''
''    SQL = "select count(*) from rclasifauto where numnotac = " & Notaca
''
''    SeInserta = (TotalRegistros(SQL) = 0)
''
''    If SeInserta Then
''        If Situacion = 2 Then
''            ' si no hay nota asociada no puedo meter la clasificacion
''            SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
''            SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
''            SQL = SQL & "`observac`,`situacion`) values ("
''            SQL = SQL & DBSet(Notaca, "N") & ","
''            SQL = SQL & DBSet(0, "N") & ","
''            SQL = SQL & DBSet(0, "N") & ","
''            SQL = SQL & DBSet(0, "N") & ","
''            SQL = SQL & DBSet(KilosTot, "N") & ","
''            SQL = SQL & DBSet(Destri, "N") & ","
''            SQL = SQL & DBSet(Podrid, "N") & ","
''            SQL = SQL & DBSet(Pequen, "N") & ","
''            SQL = SQL & DBSet(Observ, "T") & ","
''            SQL = SQL & DBSet(Situacion, "N") & ")"
''
''        Else
''            ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
''            ' tabla: rclasifauto
''            SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
''            SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
''            SQL = SQL & "`observac`,`situacion`) values ("
''            SQL = SQL & DBSet(Notaca, "N") & ","
''            SQL = SQL & DBSet(RS!Codsocio, "N") & ","
''            SQL = SQL & DBSet(RS!CodCampo, "N") & ","
''            SQL = SQL & DBSet(RS!CodVarie, "N") & ","
''            SQL = SQL & DBSet(Round2(KilosTot, 0), "N") & ","
''            SQL = SQL & DBSet(Destri, "N") & ","
''            SQL = SQL & DBSet(Podrid, "N") & ","
''            SQL = SQL & DBSet(Pequen, "N") & ","
''            SQL = SQL & DBSet(Observ, "T") & ","
''            SQL = SQL & DBSet(Situacion, "N") & ")"
''        End If
''        conn.Execute SQL
''
''        ' tabla: rclasifauto_clasif
''        SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
''        SQL = SQL & " values "
''
''    End If
''
''    'solo si tenemos nota asociada metemos toda la clasificacion
''    If Situacion <> 2 Then
''
''        'borramos la tabla temporal
''        SQLaux = "delete from tmpcata"
''        conn.Execute SQLaux
''
''        ' cargamos la tabla temporal
''        For I = 1 To NroCalidad
''            If NomCal(I) <> "" Then
''                nRegs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(I), "N"))
''                If nRegs = 0 Then
''                    'insertamos en la temporal
''                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(I), "N")
''                    SQLaux = SQLaux & "," & DBSet(KilCal(I), "N") & ")"
''
''                    conn.Execute SQLaux
''                Else
''                    'actualizamos la temporal
''                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(I), "N")
''                    SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(I), "N")
''
''                    conn.Execute SQLaux
''                End If
''            End If
''        Next I
''
''        SQLaux = "select * from tmpcata order by codcalid"
''
''        Set RSaux = New ADODB.Recordset
''        RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''
''        Sql2 = ""
''
''        While Not RSaux.EOF
''            If SeInserta Then
''                Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(RS!CodVarie, "N") & ","
''                Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
''            Else
''                Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
''                Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
''                Sql2 = Sql2 & " and codvarie = " & DBSet(RS!CodVarie, "N")
''                Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
''
''                conn.Execute Sql2
''            End If
''
''            RSaux.MoveNext
''        Wend
''
''        Set RSaux = Nothing
''
''
''        If SeInserta Then
''            If Sql2 <> "" Then
''                Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
''            End If
''            SQL = SQL & Sql2
''            conn.Execute SQL
''        End If
''    End If ' si la situacion es distinta de 2
'
'    Set Rs = Nothing
'    Set NomCal = Nothing
'    Set KilCal = Nothing
'
'    ProcesarFicheroCatadauCPequeño = True
'    Exit Function
'
'eProcesarFicheroCatadauCPequeño:
'    If Err.Number <> 0 Then
'        ProcesarFicheroCatadauCPequeño = False
'        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
'    End If
'End Function
'
'




''************************************************************************************
''*****************PROCESO DE TRASPASO DE CALIBRADOR DE ALZIRA************************
''************************************************************************************
'
'Private Function ProcesarDirectorioAlzira(nomDir As String) As Boolean
'Dim NF As Long
'Dim cad As String
'Dim I As Integer
'Dim longitud As Long
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim NumReg As Long
'Dim Sql As String
'Dim Sql1 As String
'Dim total As Long
'Dim v_cant As Currency
'Dim v_impo As Currency
'Dim v_prec As Currency
'Dim b As Boolean
'Dim NomFic As String
'
'    ProcesarDirectorioAlzira = False
'    b = True
'    ' Muestra los nombres en C:\ que representan directorios.
'    Select Case Combo1(6).ListIndex
'        Case 0, 1 ' calibrador 1 y 2 son txt
'            NomFic = Dir(nomDir & "*.txt")  ' Recupera la primera entrada.
'        Case 2 ' calibrador 3 (kaki) es .PTD
'            NomFic = Dir(nomDir & "*.ptd")  ' Recupera la primera entrada.
'    End Select
'
'    If Combo1(6).ListIndex = 0 Then
'    ' caso del precalibrado: cargamos todo el fichero en una tabla temporal
'
'        Do While NomFic <> "" And b   ' Inicia el bucle.
'           ' Ignora el directorio actual y el que lo abarca.
'           If NomFic <> "." And NomFic <> ".." Then
'              ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
'              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
'
'                Sql = "delete from tmpcalibrador"
'                conn.Execute Sql
'
'                Sql = "load data local infile '" & Replace(nomDir & NomFic, "\", "/") & "' into table `tmpcalibrador` fields escaped by '\\' enclosed by '""' lines terminated by '\r\n' ( `numnota`, `fecnota`, `nomcalid`, `kilos1`, `kilos2`, `kilos3`, `kilos4`)  "
'                conn.Execute Sql
'
'                lblProgres(0).Caption = "Procesando Fichero: " & NomFic
'                longitud = TotalRegistros("select count(*) from tmpcalibrador")
'
'                pb1.visible = True
'                Me.pb1.Max = longitud
'                Me.Refresh
'                Me.pb1.Value = 0
'
'                If longitud <> 0 Then
'                    b = ProcesarFicheroAlziraPrecalib()
'                End If
'
'              End If   ' solamente si representa un directorio.
'           End If
'           NomFic = Dir   ' Obtiene siguiente entrada.
'        Loop
'
'    Else
'    ' caso de escandalladora y el calibrador kaki se lee línea a linea del fichero de entrada
'        Do While NomFic <> "" And b   ' Inicia el bucle.
'           ' Ignora el directorio actual y el que lo abarca.
'           If NomFic <> "." And NomFic <> ".." Then
'              ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
'              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
'                NF = FreeFile
'
'                Open nomDir & NomFic For Input As #NF
'
'                Line Input #NF, cad
'
'                lblProgres(0).Caption = "Procesando Fichero: " & NomFic
'                longitud = FileLen(nomDir & NomFic)
'
'                pb1.visible = True
'                Me.pb1.Max = longitud
'                Me.Refresh
'                Me.pb1.Value = 0
'
'                If cad <> "" Then
'                    Select Case Combo1(6).ListIndex
'                        Case 1  'escandalladora
'                            b = ProcesarFicheroAlziraEscandalladora(NF, cad)
'                        Case 2  'Kaki
'                            b = ProcesarFicheroAlziraKaki(NF, cad)
'                    End Select
'                End If
'
'                Close #NF
'
'              End If   ' solamente si representa un directorio.
'           End If
'        NomFic = Dir   ' Obtiene siguiente entrada.
'        Loop
'    End If
'
'    ProcesarDirectorioAlzira = b
'
'    pb1.visible = False
'    lblProgres(0).Caption = ""
'    lblProgres(1).Caption = ""
'
'End Function
'
'
'
'Private Function ProcesarFicheroAlziraEscandalladora(NF As Long, cad As String) As Boolean
'Dim b As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim Mens As String
'Dim numlinea As Long
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim RSaux As ADODB.Recordset
'
'Dim Codsoc As String
'Dim Codcam As String
'Dim codpro As String
'Dim codVar As String
'Dim Observ As String
'Dim Notaca As String
'Dim Kilone As String
'
'Dim Destri As String
'Dim Podrid As String
'Dim Pequen As String
'Dim Muestra As String
'
'Dim NGrupos As String
'Dim Nombre1 As String
'Dim Kilos As Currency
'
'
'Dim I As Integer
'Dim J As Integer
'Dim Situacion As Byte
'
'Dim NomCal As Dictionary
'Dim KilCal As Dictionary
'
'Dim SQLaux As String
'Dim Nregs As Integer
'
'Dim Nsep As Integer
'
'Dim SeInserta As Boolean
'Dim KilosTot As Currency
'Dim cantidad As Long
'
'Dim HoraIni As String
'Dim HoraFin As String
'
'Dim FechaEnt As String
'Dim UltimaLinea As Boolean
'Dim NroCalidad As Integer
'Dim Linea As String
'
'    On Error GoTo eProcesarFicheroAlziraEscandalladora
'
'    ProcesarFicheroAlziraEscandalladora = False
'
'    Codsoc = 0
'    Codcam = 0
'    codpro = 0
'    codVar = 0
'    Observ = ""
'    Notaca = 0
'    Kilone = 0
'    Kilos = 0
'    KilosTot = 0
'
'    Destri = 0
'    Podrid = 0
'    Pequen = 0
'
'    I = 0
'    J = 0
'
'    ' inicializamos las variables
'    Set NomCal = New Dictionary
'    Set KilCal = New Dictionary
'
'
'    Notaca = RecuperaValorNew(cad, ";", 1)
'
'    Sql = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Rs.EOF Then
'        Observ = "NOTA NO EXISTE"
'        Situacion = 2
'    Else
'        codVar = DBLet(Rs!CodVarie, "N")
'    End If
'
'    b = True
'    UltimaLinea = False
'    NroCalidad = 0
'    While Not EOF(NF) And Not UltimaLinea
'        I = I + 1
'
'        Me.pb1.Value = Me.pb1.Value + Len(cad)
'        lblProgres(1).Caption = "Linea " & I
'        Me.Refresh
'
'        Nsep = NumeroSubcadenasInStr(cad, ";")
'
'        If Nsep = 14 Then ' estamos en una calidad
'            J = J + 1
'            NroCalidad = NroCalidad + 1
'
'            Linea = RecuperaValorNew(cad, ";", 2)
'
'            If CCur(Linea) = 1 Then
'                Nombre1 = RecuperaValorNew(cad, ";", 4)
'
'                ' quitamos "x.- " del nombre
'                If InStr(1, Nombre1, ".- ") <> 0 Then
'                    Nombre1 = Mid(Nombre1, InStr(1, Nombre1, ".- ") + 3, Len(Nombre1))
'                End If
'
'                Kilone = RecuperaValorNew(cad, ";", 7)
'                cantidad = RecuperaValorNew(cad, ";", 8)
'
'                Kilos = Round2(CCur(Kilone) / 1000, 2)
'                KilosTot = KilosTot + Kilos
'
'                If Situacion <> 2 Then
'                    ' si hay nota asociada busco los datos
'                    Sql = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(Rs!CodVarie, "N")
'                    Sql = Sql & " and nomcalibrador2 = " & DBSet(Trim(Nombre1), "T")
'
'                    Set RS1 = New ADODB.Recordset
'                    RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                    If RS1.EOF Then
'                        Observ = "NO EXIS.CAL"
'                        Situacion = 1
'                    Else
'                        NomCal(J) = DBLet(RS1!codcalid, "N")
'                        KilCal(J) = Kilos
'                    End If
'                    Set RS1 = Nothing
'
'                End If
'            Else ' se trata de destrio
'                Kilone = RecuperaValorNew(cad, ";", 7)
'
'                Kilos = Round2(CCur(Kilone) / 1000, 2)
'
'                Destri = Destri + Kilos
'            End If
'        End If
'
'        If Nsep = 15 Then ' estamos en la ultima linea
'            HoraIni = RecuperaValorNew(cad, ";", 9)
'            HoraFin = RecuperaValorNew(cad, ";", 10)
'            FechaEnt = RecuperaValorNew(cad, ";", 11)
'
'            UltimaLinea = True
'        End If
'
'        Line Input #NF, cad
'    Wend
'
''    Close #NF
'
'' solo tenemos la suma de kilos de destrio
'    If Situacion <> 2 Then
'        If Destri <> 0 Then
'            ' si hay kilos de destrio buscamos cual es la calidad de destrio
'            Sql = "select codcalid from rcalidad where codvarie = " & DBSet(codVar, "N")
'            Sql = Sql & " and tipcalid = 1 " ' calidad de destrio
'
'            Set RS1 = New ADODB.Recordset
'            RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            If RS1.EOF Then
'                Observ = "NO HAY DESTRIO"
'                Situacion = 5
'            Else
'                NomCal(J) = RS1.Fields(0).Value
'                KilCal(J) = Destri
'
'                NroCalidad = NroCalidad + 1
'            End If
'
'            Set RS1 = Nothing
'        End If
'    End If
'
''    If DBLet(Rs.Fields(0).Value, "N") <> KilosTot Then
''        Observ = "K.NETOS DIF."
''        Situacion = 4
''    End If
'
'    Sql = "select count(*) from rclasifauto where numnotac = " & Notaca
'
'    SeInserta = (TotalRegistros(Sql) = 0)
'
'    If SeInserta Then
'        If Situacion = 2 Then
'            ' si no hay nota asociada no puedo meter la clasificacion
'            Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'            Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'            Sql = Sql & "`observac`,`situacion`) values ("
'            Sql = Sql & DBSet(Notaca, "N") & ","
'            Sql = Sql & DBSet(0, "N") & ","
'            Sql = Sql & DBSet(0, "N") & ","
'            Sql = Sql & DBSet(0, "N") & ","
'            Sql = Sql & DBSet(KilosTot, "N") & ","
'            Sql = Sql & DBSet(Destri, "N") & ","
'            Sql = Sql & DBSet(Podrid, "N") & ","
'            Sql = Sql & DBSet(Pequen, "N") & ","
'            Sql = Sql & DBSet(Observ, "T") & ","
'            Sql = Sql & DBSet(Situacion, "N") & ")"
'
'        Else
'            ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
'            ' tabla: rclasifauto
'            Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'            Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'            Sql = Sql & "`observac`,`situacion`) values ("
'            Sql = Sql & DBSet(Notaca, "N") & ","
'            Sql = Sql & DBSet(Rs!Codsocio, "N") & ","
'            Sql = Sql & DBSet(Rs!codCampo, "N") & ","
'            Sql = Sql & DBSet(Rs!CodVarie, "N") & ","
'            Sql = Sql & DBSet(KilosTot, "N") & ","
'            Sql = Sql & DBSet(Destri, "N") & ","
'            Sql = Sql & DBSet(Podrid, "N") & ","
'            Sql = Sql & DBSet(Pequen, "N") & ","
'            Sql = Sql & DBSet(Observ, "T") & ","
'            Sql = Sql & DBSet(Situacion, "N") & ")"
'        End If
'        conn.Execute Sql
'
'        ' tabla: rclasifauto_clasif
'        Sql = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
'        Sql = Sql & " values "
'
'    End If
'
'    'solo si tenemos nota asociada metemos toda la clasificacion
'    If Situacion <> 2 Then
'
'        'borramos la tabla temporal
'        SQLaux = "delete from tmpcata"
'        conn.Execute SQLaux
'
'        ' cargamos la tabla temporal
'        For I = 1 To NroCalidad
'            If NomCal(I) <> "" Then
'                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(I), "N"))
'                If Nregs = 0 Then
'                    'insertamos en la temporal
'                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(I), "N")
'                    SQLaux = SQLaux & "," & DBSet(KilCal(I), "N") & ")"
'
'                    conn.Execute SQLaux
'                Else
'                    'actualizamos la temporal
'                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(I), "N")
'                    SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(I), "N")
'
'                    conn.Execute SQLaux
'                End If
'            End If
'        Next I
'
'        SQLaux = "select * from tmpcata order by codcalid"
'
'        Set RSaux = New ADODB.Recordset
'        RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        Sql2 = ""
'
'        If Not RSaux.EOF Then RSaux.MoveFirst
'
'        While Not RSaux.EOF
'            If SeInserta Then
'                Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(Rs!CodVarie, "N") & ","
'                Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
'            Else
'                Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
'                Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
'                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'                Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
'
'                conn.Execute Sql2
'            End If
'
'            RSaux.MoveNext
'        Wend
'
'        Set RSaux = Nothing
'
'        If SeInserta Then
'            If Sql2 <> "" Then
'                Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
'            End If
'            Sql = Sql & Sql2
'            conn.Execute Sql
'        End If
'    End If ' si la situacion es distinta de 2
'
'    Set Rs = Nothing
'    Set NomCal = Nothing
'    Set KilCal = Nothing
'
'    ProcesarFicheroAlziraEscandalladora = True
'    Exit Function
'
'eProcesarFicheroAlziraEscandalladora:
'    If Err.Number <> 0 Then
'        ProcesarFicheroAlziraEscandalladora = False
'        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
'    End If
'End Function
'
'
'Private Function ProcesarFicheroAlziraPrecalib() As Boolean
'Dim b As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim Mens As String
'Dim numlinea As Long
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim Rs2 As ADODB.Recordset
'Dim RSaux As ADODB.Recordset
'
'Dim Codsoc As String
'Dim Codcam As String
'Dim codpro As String
'Dim codVar As String
'Dim Observ As String
'Dim Notaca As String
'Dim Kilone As String
'
'Dim Destri As String
'Dim Podrid As String
'Dim Pequen As String
'Dim Muestra As String
'
'Dim NGrupos As String
'Dim Nombre1 As String
'Dim Kilos As Currency
'
'
'Dim I As Integer
'Dim J As Integer
'Dim Situacion As Byte
'
'Dim NomCal As Dictionary
'Dim KilCal As Dictionary
'
'Dim SQLaux As String
'Dim Nregs As Integer
'
'Dim Nsep As Integer
'
'Dim SeInserta As Boolean
'Dim KilosTot As Currency
'Dim cantidad As Long
'
'Dim HoraIni As String
'Dim HoraFin As String
'
'Dim FechaEnt As String
'Dim UltimaLinea As Boolean
'Dim NroCalidad As Integer
'Dim Linea As String
'Dim CalDestri As String
'Dim CalPeque As String
'
'
'    On Error GoTo eProcesarFicheroAlziraPrecalib
'
'    ProcesarFicheroAlziraPrecalib = False
'
'    Codsoc = 0
'    Codcam = 0
'    codpro = 0
'    codVar = 0
'    Observ = ""
'    Notaca = 0
'    Kilone = 0
'    Kilos = 0
'    KilosTot = 0
'
'    Destri = 0
'    Pequen = 0
'
'    I = 0
'
'    ' inicializamos las variables
'    Set NomCal = New Dictionary
'    Set KilCal = New Dictionary
'
'    Sql = "select * from tmpcalibrador "
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    Notaca = 0
'    If Not Rs.EOF Then
'        Notaca = DBLet(Rs.Fields(0).Value, "N")
'
'        Sql2 = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
'        Set RS1 = New ADODB.Recordset
'        RS1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        If RS1.EOF Then
'            Observ = "NOTA NO EXISTE"
'            Situacion = 2
'        End If
'
'        b = True
'
'        While Not Rs.EOF
'            I = I + 1
'
'            Me.pb1.Value = Me.pb1.Value + 1
'            lblProgres(1).Caption = "Linea " & I
'            Me.Refresh
'
'            Nombre1 = DBLet(Rs!nomcalid, "T")
'            Destri = DBLet(Rs!Kilos3, "T")
'            Pequen = DBLet(Rs!Kilos4, "T")
'
'            Kilone = DBLet(Rs!Kilos1, "T")
'
'            Kilos = Round2(CCur(Kilone), 2)
'            KilosTot = KilosTot + Kilos
'
'            If Situacion <> 2 Then
'                ' si hay nota asociada busco los datos
'                Sql = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(RS1!CodVarie, "N")
'                Sql = Sql & " and nomcalibrador1 = " & DBSet(Nombre1, "T")
'
'                Set Rs2 = New ADODB.Recordset
'                Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Rs2.EOF Then
'                    Observ = "NO EXIS.CAL"
'                    Situacion = 1
'                Else
'                    NomCal(I) = DBLet(Rs2!codcalid, "N")
'                    KilCal(I) = Kilos
'                End If
'                Set Rs2 = Nothing
'
'            End If
'
'            Rs.MoveNext
'        Wend
'
'        Sql = "select count(*) from rclasifauto where numnotac = " & Notaca
'
'        SeInserta = (TotalRegistros(Sql) = 0)
'
'        If SeInserta Then
'            If Situacion = 2 Then
'                ' si no hay nota asociada no puedo meter la clasificacion
'                Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'                Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'                Sql = Sql & "`observac`,`situacion`) values ("
'                Sql = Sql & DBSet(Notaca, "N") & ","
'                Sql = Sql & DBSet(0, "N") & ","
'                Sql = Sql & DBSet(0, "N") & ","
'                Sql = Sql & DBSet(0, "N") & ","
'                Sql = Sql & DBSet(KilosTot, "N") & ","
'                Sql = Sql & DBSet(Destri, "N") & ","
'                Sql = Sql & DBSet(Podrid, "N") & ","
'                Sql = Sql & DBSet(Pequen, "N") & ","
'                Sql = Sql & DBSet(Observ, "T") & ","
'                Sql = Sql & DBSet(Situacion, "N") & ")"
'
'            Else
'                ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
'                ' tabla: rclasifauto
'                Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'                Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'                Sql = Sql & "`observac`,`situacion`) values ("
'                Sql = Sql & DBSet(Notaca, "N") & ","
'                Sql = Sql & DBSet(RS1!Codsocio, "N") & ","
'                Sql = Sql & DBSet(RS1!codCampo, "N") & ","
'                Sql = Sql & DBSet(RS1!CodVarie, "N") & ","
'                Sql = Sql & DBSet(KilosTot, "N") & ","
'                Sql = Sql & DBSet(Destri, "N") & ","
'                Sql = Sql & DBSet(Podrid, "N") & ","
'                Sql = Sql & DBSet(Pequen, "N") & ","
'                Sql = Sql & DBSet(Observ, "T") & ","
'                Sql = Sql & DBSet(Situacion, "N") & ")"
'            End If
'            conn.Execute Sql
'
'            ' tabla: rclasifauto_clasif
'            Sql = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
'            Sql = Sql & " values "
'
'        End If
'
'        'solo si tenemos nota asociada metemos toda la clasificacion
'        If Situacion <> 2 Then
'
'            'borramos la tabla temporal
'            SQLaux = "delete from tmpcata"
'            conn.Execute SQLaux
'
'            ' cargamos la tabla temporal
'            For J = 1 To I
'                If NomCal(J) <> "" Then
'                    Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(J), "N"))
'                    If Nregs = 0 Then
'                        'insertamos en la temporal
'                        SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(J), "N")
'                        SQLaux = SQLaux & "," & DBSet(KilCal(J), "N") & ")"
'
'                        conn.Execute SQLaux
'                    Else
'                        'actualizamos la temporal
'                        SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(J), "N")
'                        SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(J), "N")
'
'                        conn.Execute SQLaux
'                    End If
'                End If
'            Next J
'
'            'le sumamos los kilos de destrio
'            CalDestri = CalidadDestrio(RS1!CodVarie)
'            If CalDestri <> "" Then
'                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(CalDestri, "N"))
'                If Nregs = 0 Then
'                    'insertamos en la temporal
'                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(CalDestri, "N")
'                    SQLaux = SQLaux & "," & DBSet(Destri, "N") & ")"
'
'                    conn.Execute SQLaux
'                Else
'                    'actualizamos la temporal
'                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(Destri, "N")
'                    SQLaux = SQLaux & " where codcalid = " & DBSet(CalDestri, "N")
'
'                    conn.Execute SQLaux
'                End If
'            End If
'
'            'le sumamos los kilos de menut
'            CalPeque = CalidadMenut(RS1!CodVarie)
'            If CalPeque <> "" Then
'                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(CalPeque, "N"))
'                If Nregs = 0 Then
'                    'insertamos en la temporal
'                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(CalPeque, "N")
'                    SQLaux = SQLaux & "," & DBSet(Pequen, "N") & ")"
'
'                    conn.Execute SQLaux
'                Else
'                    'actualizamos la temporal
'                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(Pequen, "N")
'                    SQLaux = SQLaux & " where codcalid = " & DBSet(CalPeque, "N")
'
'                    conn.Execute SQLaux
'                End If
'            End If
'
'
'
'
'            SQLaux = "select * from tmpcata order by codcalid"
'
'            Set RSaux = New ADODB.Recordset
'            RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            Sql2 = ""
'
'            While Not RSaux.EOF
'                If SeInserta Then
'                    Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(RS1!CodVarie, "N") & ","
'                    Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
'                Else
'                    Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
'                    Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
'                    Sql2 = Sql2 & " and codvarie = " & DBSet(RS1!CodVarie, "N")
'                    Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
'
'                    conn.Execute Sql2
'                End If
'
'                RSaux.MoveNext
'            Wend
'
'            Set RSaux = Nothing
'
'            If SeInserta Then
'                If Sql2 <> "" Then
'                    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
'                End If
'                Sql = Sql & Sql2
'                conn.Execute Sql
'            End If
'        End If ' si la situacion es distinta de 2
'
'        Set Rs = Nothing
'        Set RS1 = Nothing
'        Set NomCal = Nothing
'        Set KilCal = Nothing
'
'        ProcesarFicheroAlziraPrecalib = True
'        Exit Function
'
'    End If
'
'eProcesarFicheroAlziraPrecalib:
'    If Err.Number <> 0 Then
'        ProcesarFicheroAlziraPrecalib = False
'        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
'    End If
'End Function
'
'
'Private Function ProcesarFicheroAlziraKaki(NF As Long, cad As String) As Boolean
'Dim b As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim Mens As String
'Dim numlinea As Long
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim RSaux As ADODB.Recordset
'
'Dim Codsoc As String
'Dim Codcam As String
'Dim codpro As String
'Dim codVar As String
'Dim Observ As String
'Dim Notaca As String
'Dim Kilone As String
'
'Dim Destri As String
'Dim Podrid As String
'Dim Pequen As String
'Dim Muestra As String
'
'Dim NGrupos As String
'Dim Nombre1 As String
'Dim Kilos As Currency
'
'
'Dim I As Integer
'Dim J As Integer
'Dim Situacion As Byte
'
'Dim NomCal As Dictionary
'Dim KilCal As Dictionary
'
'Dim SQLaux As String
'Dim Nregs As Integer
'
'Dim Nsep As Integer
'
'Dim SeInserta As Boolean
'Dim KilosTot As Currency
'Dim cantidad As Long
'
'Dim HoraIni As String
'Dim HoraFin As String
'
'Dim FechaEnt As String
'Dim UltimaLinea As Boolean
'Dim NroCalidad As Integer
'Dim Linea As String
'Dim PorcenDestrio As String
'
'    On Error GoTo eProcesarFicheroAlziraKaki
'
'    ProcesarFicheroAlziraKaki = False
'
'    Codsoc = 0
'    Codcam = 0
'    codpro = 0
'    codVar = 0
'    Observ = ""
'    Notaca = 0
'    Kilone = 0
'    Kilos = 0
'    KilosTot = 0
'
'    Destri = 0
'    Podrid = 0
'    Pequen = 0
'
'    I = 0
'    J = 0
'
'    ' inicializamos las variables
'    Set NomCal = New Dictionary
'    Set KilCal = New Dictionary
'
'
'    ' saltamos 3 lineas
'    For J = 1 To 3
'        Line Input #NF, cad
'
'        I = I + 1
'
'        Me.pb1.Value = Me.pb1.Value + Len(cad)
'        lblProgres(1).Caption = "Linea " & I
'        Me.Refresh
'    Next J
'
'    Notaca = Mid(cad, 10, 10) ' posicion de la [10,19]
'
'    Sql = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Rs.EOF Then
'        Observ = "NOTA NO EXISTE"
'        Situacion = 2
'    Else
'        codVar = DBLet(Rs!CodVarie, "N")
'    End If
'
'    ' saltamos 9 lineas
'    For J = 1 To 10
'        Line Input #NF, cad
'
'        I = I + 1
'
'        Me.pb1.Value = Me.pb1.Value + Len(cad)
'        lblProgres(1).Caption = "Linea " & I
'        Me.Refresh
'    Next J
'
'    b = True
'    UltimaLinea = False
'    NroCalidad = 0
'
'    J = 0
'    While Not EOF(NF) And Not UltimaLinea
'        I = I + 1
'
'        Me.pb1.Value = Me.pb1.Value + Len(cad)
'        lblProgres(1).Caption = "Linea " & I
'        Me.Refresh
'
'        J = J + 1
'        NroCalidad = NroCalidad + 1
'
'        Nombre1 = Mid(cad, 6, 11)
'        Kilone = Mid(cad, 17, 11)
'        Kilos = CCur(Kilone)
'
'        KilosTot = KilosTot + Kilos
'
'        If Situacion <> 2 Then
'            ' si hay nota asociada busco los datos
'            Sql = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(Rs!CodVarie, "N")
'            Sql = Sql & " and nomcalibrador3 = " & DBSet(Trim(Nombre1), "T")
'
'            Set RS1 = New ADODB.Recordset
'            RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            If RS1.EOF Then
'                Observ = "NO EXIS.CAL"
'                Situacion = 1
'            Else
'                NomCal(J) = DBLet(RS1!codcalid, "N")
'                KilCal(J) = Kilos
''YA VEREMOS
''                ' si la calidad es de destrio sumamos los kilos a los kilos de destrio
''                If CalidadDestrio(Rs!CodVarie) = DBLet(RS1!codcalid) Then
''                    Destri = Destri + Kilos
''                End If
'            End If
'            Set RS1 = Nothing
'
'        End If
'        Line Input #NF, cad
'        UltimaLinea = (Mid(cad, 17, 11) = "-----------")
'    Wend
'
'' solo tenemos la suma de kilos de destrio
'    If Situacion <> 2 Then
'        If Destri <> 0 Then
'            ' si hay kilos de destrio buscamos cual es la calidad de destrio
'            Sql = "select codcalid from rcalidad where codvarie = " & DBSet(codVar, "N")
'            Sql = Sql & " and tipcalid = 1 " ' calidad de destrio
'
'            Set RS1 = New ADODB.Recordset
'            RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            If RS1.EOF Then
'                Observ = "NO HAY DESTRIO"
'                Situacion = 5
'' ya veremos
''            Else
''                CalDestri = DBLet(RS1!codcalid, "N")
''                ' comprobamos qu no supera el destrio no supera el 50%
''                PorcenDestrio = Round2(Destri * 100 / KilosTot, 2)
''                If PorcenDestrio >= 50 Then
''                    Observ = "DESTRIO SUPERIOR AL 50%"
''                    Situacion = 3
''                End If
'            End If
'
'            Set RS1 = Nothing
'        End If
'    End If
'
''    If DBLet(Rs.Fields(0).Value, "N") <> KilosTot Then
''        Observ = "K.NETOS DIF."
''        Situacion = 4
''    End If
'
'    Sql = "select count(*) from rclasifauto where numnotac = " & Notaca
'
'    SeInserta = (TotalRegistros(Sql) = 0)
'
'    If SeInserta Then
'        If Situacion = 2 Then
'            ' si no hay nota asociada no puedo meter la clasificacion
'            Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'            Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'            Sql = Sql & "`observac`,`situacion`) values ("
'            Sql = Sql & DBSet(Notaca, "N") & ","
'            Sql = Sql & DBSet(0, "N") & ","
'            Sql = Sql & DBSet(0, "N") & ","
'            Sql = Sql & DBSet(0, "N") & ","
'            Sql = Sql & DBSet(KilosTot, "N") & ","
'            Sql = Sql & DBSet(Destri, "N") & ","
'            Sql = Sql & DBSet(Podrid, "N") & ","
'            Sql = Sql & DBSet(Pequen, "N") & ","
'            Sql = Sql & DBSet(Observ, "T") & ","
'            Sql = Sql & DBSet(Situacion, "N") & ")"
'
'        Else
'            ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
'            ' tabla: rclasifauto
'            Sql = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'            Sql = Sql & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'            Sql = Sql & "`observac`,`situacion`) values ("
'            Sql = Sql & DBSet(Notaca, "N") & ","
'            Sql = Sql & DBSet(Rs!Codsocio, "N") & ","
'            Sql = Sql & DBSet(Rs!codCampo, "N") & ","
'            Sql = Sql & DBSet(Rs!CodVarie, "N") & ","
'            Sql = Sql & DBSet(KilosTot, "N") & ","
'            Sql = Sql & DBSet(Destri, "N") & ","
'            Sql = Sql & DBSet(Podrid, "N") & ","
'            Sql = Sql & DBSet(Pequen, "N") & ","
'            Sql = Sql & DBSet(Observ, "T") & ","
'            Sql = Sql & DBSet(Situacion, "N") & ")"
'        End If
'        conn.Execute Sql
'
'        ' tabla: rclasifauto_clasif
'        Sql = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
'        Sql = Sql & " values "
'
'    End If
'
'    'solo si tenemos nota asociada metemos toda la clasificacion
'    If Situacion <> 2 Then
'
'        'borramos la tabla temporal
'        SQLaux = "delete from tmpcata"
'        conn.Execute SQLaux
'
'        ' cargamos la tabla temporal
'        For I = 1 To NroCalidad
'            If NomCal(I) <> "" Then
'                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(I), "N"))
'                If Nregs = 0 Then
'                    'insertamos en la temporal
'                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(I), "N")
'                    SQLaux = SQLaux & "," & DBSet(KilCal(I), "N") & ")"
'
'                    conn.Execute SQLaux
'                Else
'                    'actualizamos la temporal
'                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(I), "N")
'                    SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(I), "N")
'
'                    conn.Execute SQLaux
'                End If
'            End If
'        Next I
'
'        SQLaux = "select * from tmpcata order by codcalid"
'
'        Set RSaux = New ADODB.Recordset
'        RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        Sql2 = ""
'
'        If Not RSaux.EOF Then RSaux.MoveFirst
'
'        While Not RSaux.EOF
'            If SeInserta Then
'                Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(Rs!CodVarie, "N") & ","
'                Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
'            Else
'                Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
'                Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
'                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'                Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
'
'                conn.Execute Sql2
'            End If
'
'            RSaux.MoveNext
'        Wend
'
'        Set RSaux = Nothing
'
'        If SeInserta Then
'            If Sql2 <> "" Then
'                Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
'            End If
'            Sql = Sql & Sql2
'            conn.Execute Sql
'        End If
''ya veremos
''        If Destri <> 0 Then
''            Sql = "delete from rclasifauto_clasif where numnotac = " & DBSet(Notaca, "N")
''            Sql = Sql & " and codvarie = " & DBSet(Rs!CodVarie, "N")
''            Sql = Sql & " and codcalid = " & CalDestri
''
''            conn.Execute Sql
''        End If
'    End If ' si la situacion es distinta de 2
'
'    Set Rs = Nothing
'    Set NomCal = Nothing
'    Set KilCal = Nothing
'
'    ProcesarFicheroAlziraKaki = True
'    Exit Function
'
'eProcesarFicheroAlziraKaki:
'    If Err.Number <> 0 Then
'        ProcesarFicheroAlziraKaki = False
'        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
'    End If
'End Function
'
