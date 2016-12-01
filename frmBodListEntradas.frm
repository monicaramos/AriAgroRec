VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBodListEntradas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6705
   Icon            =   "frmBodListEntradas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDiarioFrasRetirada 
      Height          =   5175
      Left            =   0
      TabIndex        =   157
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   570
         TabIndex        =   181
         Top             =   3840
         Width           =   5715
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   1350
            MaxLength       =   6
            TabIndex        =   165
            Top             =   255
            Width           =   750
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   43
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   182
            Text            =   "Text5"
            Top             =   255
            Width           =   3375
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
            Index           =   56
            Left            =   90
            TabIndex        =   183
            Top             =   120
            Width           =   885
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   19
            Left            =   1050
            MouseIcon       =   "frmBodListEntradas.frx":000C
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cooperativa"
            Top             =   270
            Width           =   240
         End
      End
      Begin VB.CommandButton Command13 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":015E
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command12 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":0468
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   41
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   168
         Text            =   "Text5"
         Top             =   3090
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   42
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   167
         Text            =   "Text5"
         Top             =   3510
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   166
         Top             =   3090
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   164
         Top             =   3510
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   163
         Tag             =   "Nº Factura|N|N|||rbodfacturas|numfactu|0000000|S|"
         Top             =   1290
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   162
         Tag             =   "Nº Factura|N|N|||rbodfacturas|numfactu|0000000|S|"
         Top             =   1665
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepDiarioFra 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   161
         Top             =   4545
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5130
         TabIndex        =   160
         Top             =   4545
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   159
         Top             =   2130
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   158
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   11
         Left            =   1620
         Picture         =   "frmBodListEntradas.frx":0772
         ToolTipText     =   "Buscar fecha"
         Top             =   2550
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   10
         Left            =   1620
         Picture         =   "frmBodListEntradas.frx":07FD
         ToolTipText     =   "Buscar fecha"
         Top             =   2130
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":0888
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":09DA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3510
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
         Index           =   66
         Left            =   675
         TabIndex        =   180
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   65
         Left            =   1005
         TabIndex        =   179
         Top             =   3525
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   64
         Left            =   1005
         TabIndex        =   178
         Top             =   3135
         Width           =   465
      End
      Begin VB.Label Label7 
         Caption         =   "Diario de Facturación"
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
         TabIndex        =   177
         Top             =   420
         Width           =   5805
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
         Index           =   63
         Left            =   675
         TabIndex        =   176
         Top             =   2895
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   62
         Left            =   960
         TabIndex        =   175
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   61
         Left            =   960
         TabIndex        =   174
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   60
         Left            =   1005
         TabIndex        =   173
         Top             =   2565
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   59
         Left            =   1005
         TabIndex        =   172
         Top             =   2220
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
         Index           =   58
         Left            =   675
         TabIndex        =   171
         Top             =   1980
         Width           =   435
      End
   End
   Begin VB.Frame FrameAutoconsumo 
      Height          =   5190
      Left            =   0
      TabIndex        =   127
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check3 
         Caption         =   "Salta página por Socio"
         Height          =   285
         Left            =   390
         TabIndex        =   153
         Top             =   4320
         Width           =   2445
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   36
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   151
         Text            =   "Text5"
         Top             =   3735
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   138
         Top             =   3735
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1650
         MaxLength       =   7
         TabIndex        =   133
         Top             =   2400
         Width           =   945
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1650
         MaxLength       =   7
         TabIndex        =   132
         Top             =   2040
         Width           =   945
      End
      Begin VB.CommandButton Command11 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":0B2C
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command10 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":0E36
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   33
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   136
         Text            =   "Text5"
         Top             =   1455
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   32
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   134
         Text            =   "Text5"
         Top             =   1095
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   131
         Top             =   1455
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   32
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   130
         Top             =   1095
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepAutocons 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   129
         Top             =   4620
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   5190
         TabIndex        =   128
         Top             =   4605
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   137
         Top             =   3285
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   135
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1350
         MouseIcon       =   "frmBodListEntradas.frx":1140
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cooperativa"
         Top             =   3735
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cooperativa"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   57
         Left            =   390
         TabIndex        =   152
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   54
         Left            =   765
         TabIndex        =   150
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   53
         Left            =   765
         TabIndex        =   149
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   52
         Left            =   390
         TabIndex        =   148
         Top             =   1830
         Width           =   540
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   9
         Left            =   1350
         Picture         =   "frmBodListEntradas.frx":1292
         ToolTipText     =   "Buscar fecha"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1350
         Picture         =   "frmBodListEntradas.frx":131D
         ToolTipText     =   "Buscar fecha"
         Top             =   3285
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   33
         Left            =   1380
         MouseIcon       =   "frmBodListEntradas.frx":13A8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   32
         Left            =   1380
         MouseIcon       =   "frmBodListEntradas.frx":14FA
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
         Index           =   51
         Left            =   390
         TabIndex        =   147
         Top             =   900
         Width           =   405
      End
      Begin VB.Label Label6 
         Caption         =   "Informe Liquidación Oliva"
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
         TabIndex        =   146
         Top             =   345
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   50
         Left            =   720
         TabIndex        =   145
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   720
         TabIndex        =   144
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   48
         Left            =   735
         TabIndex        =   143
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   735
         TabIndex        =   142
         Top             =   2940
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   35
         Left            =   405
         TabIndex        =   141
         Top             =   2700
         Width           =   450
      End
   End
   Begin VB.Frame FrameListadoConsumo 
      Height          =   5175
      Left            =   30
      TabIndex        =   40
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check4 
         Caption         =   "Resumen"
         Height          =   285
         Left            =   660
         TabIndex        =   154
         Top             =   3930
         Width           =   2445
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clasificado por"
         ForeColor       =   &H00972E0B&
         Height          =   885
         Left            =   600
         TabIndex        =   67
         Top             =   3900
         Width           =   2925
         Begin VB.OptionButton Option1 
            Caption         =   "Variedad"
            Height          =   345
            Index           =   1
            Left            =   1680
            TabIndex        =   54
            Top             =   330
            Width           =   1125
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Socio"
            Height          =   315
            Index           =   0
            Left            =   270
            TabIndex        =   53
            Top             =   330
            Width           =   1395
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   50
         Top             =   2550
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   49
         Top             =   2130
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5130
         TabIndex        =   58
         Top             =   4545
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepListCons 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   56
         Top             =   4545
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   48
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   47
         Top             =   1290
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "Text5"
         Top             =   1665
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   52
         Top             =   3510
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   51
         Top             =   3090
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text5"
         Top             =   3510
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Text5"
         Top             =   3090
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":164C
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command1 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":1956
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   255
         Left            =   570
         TabIndex        =   156
         Top             =   4230
         Visible         =   0   'False
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando datos temporal"
         Height          =   195
         Index           =   55
         Left            =   630
         TabIndex        =   155
         Top             =   4650
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   29
         Left            =   675
         TabIndex        =   66
         Top             =   1980
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   1005
         TabIndex        =   65
         Top             =   2220
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   16
         Left            =   1005
         TabIndex        =   64
         Top             =   2565
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   63
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   62
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   13
         Left            =   675
         TabIndex        =   61
         Top             =   2895
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Listado de Consumo de Entradas"
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
         TabIndex        =   60
         Top             =   420
         Width           =   5805
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   12
         Left            =   1005
         TabIndex        =   59
         Top             =   3135
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   1005
         TabIndex        =   57
         Top             =   3525
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   9
         Left            =   675
         TabIndex        =   55
         Top             =   1080
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":1C60
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":1DB2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":1F04
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3510
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":2056
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1620
         Picture         =   "frmBodListEntradas.frx":21A8
         ToolTipText     =   "Buscar fecha"
         Top             =   2565
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1620
         Picture         =   "frmBodListEntradas.frx":2233
         ToolTipText     =   "Buscar fecha"
         Top             =   2160
         Width           =   240
      End
   End
   Begin VB.Frame FrameAsignacionPreciosABN 
      Height          =   5175
      Left            =   0
      TabIndex        =   184
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   51
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   194
         Top             =   2550
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   50
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   193
         Top             =   2130
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5130
         TabIndex        =   200
         Top             =   4425
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepAsigPrecABN 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   198
         Top             =   4425
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   192
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   48
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   191
         Top             =   1290
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   49
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   190
         Text            =   "Text5"
         Top             =   1665
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   48
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   189
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   47
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   196
         Top             =   3510
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   46
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   195
         Top             =   3090
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   47
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   188
         Text            =   "Text5"
         Top             =   3510
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   46
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   187
         Text            =   "Text5"
         Top             =   3090
         Width           =   3375
      End
      Begin VB.CommandButton Command15 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":22BE
         Style           =   1  'Graphical
         TabIndex        =   186
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command14 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":25C8
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   77
         Left            =   675
         TabIndex        =   208
         Top             =   1980
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   76
         Left            =   1005
         TabIndex        =   207
         Top             =   2220
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   75
         Left            =   1005
         TabIndex        =   206
         Top             =   2565
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   74
         Left            =   960
         TabIndex        =   205
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   73
         Left            =   960
         TabIndex        =   204
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   72
         Left            =   675
         TabIndex        =   203
         Top             =   2895
         Width           =   390
      End
      Begin VB.Label Label8 
         Caption         =   "Asignación de Precios Masiva"
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
         TabIndex        =   202
         Top             =   450
         Width           =   5805
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   71
         Left            =   1005
         TabIndex        =   201
         Top             =   3135
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   70
         Left            =   1005
         TabIndex        =   199
         Top             =   3525
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   69
         Left            =   675
         TabIndex        =   197
         Top             =   1080
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1650
         MouseIcon       =   "frmBodListEntradas.frx":28D2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   1650
         MouseIcon       =   "frmBodListEntradas.frx":2A24
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":2B76
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3510
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   22
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":2CC8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   13
         Left            =   1620
         Picture         =   "frmBodListEntradas.frx":2E1A
         ToolTipText     =   "Buscar fecha"
         Top             =   2565
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   12
         Left            =   1620
         Picture         =   "frmBodListEntradas.frx":2EA5
         ToolTipText     =   "Buscar fecha"
         Top             =   2160
         Width           =   240
      End
   End
   Begin VB.Frame FrameAsignacionPrecios 
      Height          =   5175
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1920
         MaxLength       =   12
         TabIndex        =   79
         Top             =   4395
         Width           =   1365
      End
      Begin VB.CommandButton CmdAcepGastosLiq 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   96
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   1935
         MaxLength       =   12
         TabIndex        =   78
         Top             =   3990
         Width           =   1365
      End
      Begin VB.CommandButton Command7 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":2F30
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command6 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":323A
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Text5"
         Top             =   3090
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text5"
         Top             =   3510
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   75
         Top             =   3090
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   77
         Top             =   3510
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   1665
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   69
         Top             =   1290
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   71
         Top             =   1665
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepAsigPrec 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   81
         Top             =   4425
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   5130
         TabIndex        =   83
         Top             =   4425
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   73
         Top             =   2130
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   74
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio Excedido"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   34
         Left            =   690
         TabIndex        =   97
         Top             =   4380
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio Venta"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   33
         Left            =   675
         TabIndex        =   95
         Top             =   3975
         Width           =   915
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1620
         Picture         =   "frmBodListEntradas.frx":3544
         ToolTipText     =   "Buscar fecha"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   1620
         Picture         =   "frmBodListEntradas.frx":35CF
         ToolTipText     =   "Buscar fecha"
         Top             =   2565
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":365A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":37AC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3510
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":38FE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":3A50
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   32
         Left            =   675
         TabIndex        =   94
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   31
         Left            =   1005
         TabIndex        =   93
         Top             =   3525
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   30
         Left            =   1005
         TabIndex        =   92
         Top             =   3135
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Asignación de Precios Masiva"
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
         TabIndex        =   91
         Top             =   450
         Width           =   5805
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   8
         Left            =   675
         TabIndex        =   90
         Top             =   2895
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   89
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   88
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   1005
         TabIndex        =   87
         Top             =   2565
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   1005
         TabIndex        =   86
         Top             =   2220
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   675
         TabIndex        =   85
         Top             =   1980
         Width           =   450
      End
   End
   Begin VB.Frame FrameBonificacion 
      Height          =   5190
      Left            =   0
      TabIndex        =   99
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   110
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   112
         Top             =   3285
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5190
         TabIndex        =   116
         Top             =   4605
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarBonif 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   114
         Top             =   4620
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   106
         Top             =   1095
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   107
         Top             =   1455
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "Text5"
         Top             =   1095
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   28
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   104
         Text            =   "Text5"
         Top             =   1455
         Width           =   3375
      End
      Begin VB.CommandButton Command8 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":3BA2
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command5 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":3EAC
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   108
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   109
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   101
         Text            =   "Text5"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   26
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   100
         Text            =   "Text5"
         Top             =   2400
         Width           =   3375
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   420
         TabIndex        =   111
         Top             =   3870
         Visible         =   0   'False
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   46
         Left            =   405
         TabIndex        =   126
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   45
         Left            =   735
         TabIndex        =   125
         Top             =   2940
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   44
         Left            =   735
         TabIndex        =   124
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   43
         Left            =   720
         TabIndex        =   123
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   42
         Left            =   720
         TabIndex        =   122
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label5 
         Caption         =   "Cálculo de Bonificación"
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
         TabIndex        =   121
         Top             =   345
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   41
         Left            =   390
         TabIndex        =   120
         Top             =   900
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1350
         MouseIcon       =   "frmBodListEntradas.frx":41B6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1350
         MouseIcon       =   "frmBodListEntradas.frx":4308
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1350
         Picture         =   "frmBodListEntradas.frx":445A
         ToolTipText     =   "Buscar fecha"
         Top             =   3285
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1350
         Picture         =   "frmBodListEntradas.frx":44E5
         ToolTipText     =   "Buscar fecha"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   40
         Left            =   390
         TabIndex        =   119
         Top             =   1830
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   39
         Left            =   765
         TabIndex        =   118
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   38
         Left            =   765
         TabIndex        =   117
         Top             =   2475
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1350
         MouseIcon       =   "frmBodListEntradas.frx":4570
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1350
         MouseIcon       =   "frmBodListEntradas.frx":46C2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   37
         Left            =   450
         TabIndex        =   115
         Top             =   4200
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   36
         Left            =   450
         TabIndex        =   113
         Top             =   4410
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6960
      Top             =   5220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameEntradasCampo 
      Height          =   6855
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check2 
         Caption         =   "Agrupado por Variedad"
         Height          =   285
         Left            =   630
         TabIndex        =   98
         Top             =   6060
         Width           =   2445
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Salta página por Socio"
         Height          =   285
         Left            =   630
         TabIndex        =   39
         Top             =   5910
         Width           =   2445
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   8
         Top             =   4485
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   7
         Top             =   4095
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   4485
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text5"
         Top             =   4095
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   4
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2220
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":4814
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListEntradas.frx":4B1E
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   6
         Top             =   3570
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   5
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   2
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   11
         Top             =   6135
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5130
         TabIndex        =   12
         Top             =   6135
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   10
         Top             =   5445
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   9
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Depósito"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   660
         TabIndex        =   38
         Top             =   3930
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   990
         TabIndex        =   37
         Top             =   4170
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   36
         Top             =   4560
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":4E28
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar depósito"
         Top             =   4500
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":4F7A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar depósito"
         Top             =   4110
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":50CC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":521E
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
         TabIndex        =   33
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   1005
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   2025
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1590
         Picture         =   "frmBodListEntradas.frx":5370
         ToolTipText     =   "Buscar fecha"
         Top             =   5040
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1590
         Picture         =   "frmBodListEntradas.frx":53FB
         ToolTipText     =   "Buscar fecha"
         Top             =   5445
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":5486
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":55D8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":572A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1620
         MouseIcon       =   "frmBodListEntradas.frx":587C
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
         TabIndex        =   28
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   1005
         TabIndex        =   27
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   1005
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   420
         Width           =   5805
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   675
         TabIndex        =   24
         Top             =   3015
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   960
         TabIndex        =   23
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   960
         TabIndex        =   22
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   975
         TabIndex        =   21
         Top             =   5445
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   975
         TabIndex        =   20
         Top             =   5100
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   645
         TabIndex        =   19
         Top             =   4860
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmBodListEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-+

'   LISTADO DE ENTRADAS DE BODEGA

Option Explicit

Public OpcionListado  As String
    ' 0 = Informe de Entradas de Bodega
    ' 1 = Extracto de entradas por Socio / Variedad

    ' 2 = Listado de consumo de entradas de vino
    
    ' 3 = Asignacion de precios masiva en albaranes de retirada aceite / vino (almazara/bodega)
    ' 4 = Reparto de Gastos de liquidacion de bodega
    ' 5 = Reparto de Gastos de liquidacion de almazara
    
    ' 6 = Diferencia de consumo/producido por socio
    
    ' 7 = Calculo de Porcentaje bonificado de bodega
    
    
    ' 8 = Informe de autoconsumo (facturas de liquidacion de almazara VALSUR)
    ' 9 = Diario de facturas de retirada
    
Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmDep As frmManDepositos ' DEPOSITOS
Attribute frmDep.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmCla As frmComercial 'Ayuda de Clases de comercial
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmCoop As frmManCoope  ' cooperativas
Attribute frmCoop.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
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
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub CmdAcepAsigPrec_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vSocio As cSocio

Dim Rs As ADODB.Recordset
Dim Importe As String
Dim Sql As String

    If txtcodigo(22).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un precio de venta. Revise.", vbExclamation
        PonerFoco txtcodigo(22)
        Exit Sub
    End If

    If txtcodigo(23).Text = "" Then
        txtcodigo(23).Text = txtcodigo(22).Text
    End If

    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H SOCIO
    cDesde = Trim(txtcodigo(4).Text)
    cHasta = Trim(txtcodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    'D/H VARIEDAD
    cDesde = Trim(txtcodigo(18).Text)
    cHasta = Trim(txtcodigo(19).Text)
    nDesde = txtNombre(18).Text
    nHasta = txtNombre(19).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodalbaran_variedad.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
    End If

    'D/H fecha
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
    nTabla = "(rbodalbaran INNER JOIN rbodalbaran_variedad ON rbodalbaran.numalbar = rbodalbaran_variedad.numalbar) "
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        
        If Not BloqueaRegistro(nTabla, cadSelect) Then
            MsgBox "No se pueden Actualizar precios. Hay registros bloqueados.", vbExclamation
            Screen.MousePointer = vbDefault
        Else
            If ProcesarCambios(nTabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (2)
            End If
        End If
    End If


End Sub

Private Sub CmdAcepAsigPrecABN_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vSocio As cSocio

Dim Rs As ADODB.Recordset
Dim Importe As String
Dim Sql As String


    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H SOCIO
    cDesde = Trim(txtcodigo(48).Text)
    cHasta = Trim(txtcodigo(49).Text)
    nDesde = txtNombre(48).Text
    nHasta = txtNombre(49).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    'D/H CLASE
    cDesde = Trim(txtcodigo(46).Text)
    cHasta = Trim(txtcodigo(47).Text)
    nDesde = txtNombre(46).Text
    nHasta = txtNombre(47).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
    End If

    'D/H fecha
    cDesde = Trim(txtcodigo(50).Text)
    cHasta = Trim(txtcodigo(51).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
    nTabla = "(rbodalbaran INNER JOIN rbodalbaran_variedad ON rbodalbaran.numalbar = rbodalbaran_variedad.numalbar) "
    nTabla = nTabla & " INNER JOIN variedades ON rbodalbaran_variedad.codvarie = variedades.codvarie "
        
        
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        
        '[Monica]10/03/2016:
        If ProcesoYaRealizado(nTabla, cadSelect) Then
            If MsgBox("El proceso de cálculo de precios ya ha sido realizado" & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        End If
        
        
        If Not BloqueaRegistro(nTabla, cadSelect) Then
            MsgBox "No se pueden Actualizar precios. Hay registros bloqueados.", vbExclamation
            Screen.MousePointer = vbDefault
        Else
            If ProcesarCambiosABN(nTabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (2)
            End If
        End If
    End If


End Sub


Private Function ProcesoYaRealizado(nTabla As String, cadSelect As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from " & nTabla
    Sql = Sql & " where ampliaci = 'Regularización de Precios' "
    If cadSelect <> "" Then Sql = Sql & " and " & cadSelect

    ProcesoYaRealizado = (TotalRegistros(Sql) <> 0)


End Function


Private Sub CmdAcepAutocons_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vSocio As cSocio

Dim Rs As ADODB.Recordset
Dim Importe As String
Dim Sql As String
Dim Albaranes As String
Dim NomGasto As String
Dim Cad As String

Dim tipoMov As String

    If Not DatosOk Then Exit Sub



    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H SOCIO
    cDesde = Trim(txtcodigo(32).Text)
    cHasta = Trim(txtcodigo(33).Text)
    nDesde = txtNombre(32).Text
    nHasta = txtNombre(33).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    'D/H FACTURA
    cDesde = Trim(txtcodigo(34).Text)
    cHasta = Trim(txtcodigo(35).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc_albaran.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura=""") Then Exit Sub
    End If

    'D/H fecha
    cDesde = Trim(txtcodigo(24).Text)
    cHasta = Trim(txtcodigo(31).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc_albaran.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
    tipoMov = DevuelveDesdeBDNew(cAgro, "rcoope", "codtipomliqalmz", "codcoope", txtcodigo(36).Text, "N")
    
    If Not AnyadirAFormula(cadSelect, "{rfactsoc_albaran.codtipom} = '" & tipoMov & "'") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "{rfactsoc_albaran.codtipom} = '" & tipoMov & "'") Then Exit Sub
    
    ' seleccionamos unicamente las lineas de autoconsumo
    If Not AnyadirAFormula(cadSelect, "{rfactsoc_albaran.kilosnet} <> 0") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "{rfactsoc_albaran.kilosnet} <> 0") Then Exit Sub
    
    ' salto de pagina por socio
    CadParam = CadParam & "pSalto=" & Check3.Value & "|"
    numParam = numParam + 1
    
    
    nTabla = "rfactsoc_albaran INNER JOIN rfactsoc ON rfactsoc_albaran.codtipom = rfactsoc.codtipom "
    nTabla = nTabla & " and rfactsoc_albaran.numfactu = rfactsoc.numfactu "
    nTabla = nTabla & " and rfactsoc_albaran.fecfactu = rfactsoc.fecfactu "
    nTabla = "(" & nTabla & ") INNER JOIN rsocios ON rfactsoc.codsocio = rsocios.codsocio "
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        cadNombreRPT = "rBodLiqAutoconsumo.rpt"
        cadTitulo = "Listado Liquidación Autoconsumo"
        
        LlamarImprimir
    End If
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
    
    b = True
    Select Case OpcionListado
        Case 8 ' informe de autoconsumo
            If txtcodigo(36).Text = "" Then
                MsgBox "Debe introducir obligatoriamente la cooperativa. Revise.", vbExclamation
                b = False
            End If
    End Select
    
    DatosOk = b
End Function


Private Sub CmdAcepDiarioFra_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vSocio As cSocio

Dim Rs As ADODB.Recordset
Dim Importe As String
Dim Sql As String
Dim Albaranes As String
Dim NomGasto As String
Dim Cad As String

Dim tipoMov As String

    If Not DatosOk Then Exit Sub

    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H SOCIO
    cDesde = Trim(txtcodigo(41).Text)
    cHasta = Trim(txtcodigo(42).Text)
    nDesde = txtNombre(41).Text
    nHasta = txtNombre(42).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodfacturas.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    'D/H FACTURA
    cDesde = Trim(txtcodigo(39).Text)
    cHasta = Trim(txtcodigo(40).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodfacturas.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura=""") Then Exit Sub
    End If

    'D/H fecha
    cDesde = Trim(txtcodigo(37).Text)
    cHasta = Trim(txtcodigo(38).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodfacturas.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
    '[Monica]01/04/2016: solo si es abn miramos cooperativa
    If vParamAplic.Cooperativa = 1 Then
        tipoMov = DevuelveDesdeBDNew(cAgro, "rcoope", "codtipomfacalmz", "codcoope", txtcodigo(43).Text, "N")
        
        If Not AnyadirAFormula(cadSelect, "{rbodfacturas.codtipom} = '" & tipoMov & "'") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rbodfacturas.codtipom} = '" & tipoMov & "'") Then Exit Sub
    End If
    
    
    nTabla = "rbodfacturas INNER JOIN rsocios ON rbodfacturas.codsocio = rsocios.codsocio "
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        cadNombreRPT = "rBodDiarioFacturas.rpt"
        cadTitulo = "Diario de Facturación"
        
        LlamarImprimir
    End If

End Sub

Private Sub CmdAcepGastosLiq_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vSocio As cSocio

Dim Rs As ADODB.Recordset
Dim Importe As String
Dim Sql As String
Dim Albaranes As String
Dim NomGasto As String
Dim Cad As String


    NomGasto = ""
    If OpcionListado = 4 Then
        NomGasto = DevuelveDesdeBDNew(cAgro, "rconcepgasto", "nomgasto", "codgasto", vParamAplic.CodGastoBOD, "N")
    Else
        NomGasto = DevuelveDesdeBDNew(cAgro, "rconcepgasto", "nomgasto", "codgasto", vParamAplic.CodGastoAlmz, "N")
    End If
    
    If NomGasto = "" Then
        MsgBox "No existe el concepto de gasto para el prorrateo o no se ha especificado en parámetros. Revise.", vbExclamation
        Exit Sub
    End If

    If txtcodigo(22).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un importe de gastos a repartir. Revise.", vbExclamation
        PonerFoco txtcodigo(22)
        Exit Sub
    End If

    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H SOCIO
    cDesde = Trim(txtcodigo(4).Text)
    cHasta = Trim(txtcodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    'D/H VARIEDAD
    cDesde = Trim(txtcodigo(18).Text)
    cHasta = Trim(txtcodigo(19).Text)
    nDesde = txtNombre(18).Text
    nHasta = txtNombre(19).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rhisfruta.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
    End If

    'D/H fecha
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
    ' sólo los registros del hco de entradas de bodega
    nTabla = "(rhisfruta INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
    Select Case OpcionListado
        Case 4
            nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu and codgrupo = 6 "
        Case 5
            nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu and codgrupo = 5 "
    End Select
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        Albaranes = CadenaAlbaranes(nTabla, cadSelect)
        
        Cad = "Este proceso modifica los gastos de albaranes eliminando previamente los correspondiente al concepto:  "
        If OpcionListado = 4 Then
            Cad = Cad & vParamAplic.CodGastoBOD & " - " & NomGasto
        Else
            Cad = Cad & vParamAplic.CodGastoAlmz & " - " & NomGasto
        End If
        Cad = Cad & vbCrLf & vbCrLf & "              ¿ Desea continuar ? "
        
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            If Not BloqueaRegistro("rhisfruta_gastos", "numalbar in (" & Albaranes & ")") Then
                MsgBox "No se pueden prorratear gastos liquidación. Hay registros bloqueados.", vbExclamation
                Screen.MousePointer = vbDefault
            Else
                If ProcesarRepartoGastos("rhisfruta", "numalbar in (" & Albaranes & ")", txtcodigo(22).Text) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (2)
                End If
            End If
        End If
    End If

End Sub

Private Sub CmdAcepListCons_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String


    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H SOCIO
    cDesde = Trim(txtcodigo(10).Text)
    cHasta = Trim(txtcodigo(11).Text)
    nDesde = txtNombre(10).Text
    nHasta = txtNombre(11).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    'D/H VARIEDAD
    cDesde = Trim(txtcodigo(8).Text)
    cHasta = Trim(txtcodigo(9).Text)
    nDesde = txtNombre(8).Text
    nHasta = txtNombre(9).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodalbaran_variedad.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
    End If

    'D/H fecha
    cDesde = Trim(txtcodigo(16).Text)
    cHasta = Trim(txtcodigo(17).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
    nTabla = "(rbodalbaran INNER JOIN rbodalbaran_variedad ON rbodalbaran.numalbar = rbodalbaran_variedad.numalbar) "
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
    
        Select Case OpcionListado
            Case 2 ' listado de consumo
                cadNombreRPT = "rBodListConsumo.rpt"
                cadTitulo = "Listado de Consumo"
                
                If Me.Option1(0).Value = True Then
                    numOp = PonerGrupo(1, "Socio")
                    CadParam = CadParam & "pTipo=0|"
                End If
                
                If Me.Option1(1).Value = True Then
                    numOp = PonerGrupo(1, "Variedad")
                    CadParam = CadParam & "pTipo=1|"
                End If
                numParam = numParam + 1
                LlamarImprimir
                
           Case 6 ' diferencia de consumo/producido por socio
                
                ' si no es resumen el listado saca las diferencias por variedad (valsur tiene la misma variedad de entrada que de salida)
                ' si es resumen las diferencias van por socio (moixent no tiene la misma variedad de entrada que de salida) solo indicamos
                '               que las variedades sean del grupo 5 almazara
                
                If Check4.Value = 0 Then ' resumen = false
                    If vParamAplic.Cooperativa = 1 Then
                        '[Monica]09/03/2016: ahora va por clases pq hay distintas variedades
                        If CargarDatosTemporalABN(nTabla, cadSelect) Then
                            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                            
                            cadNombreRPT = "rBodListDiferencia.rpt"
                            cadTitulo = "Listado de Diferencia Consumo/Producido"
                            
                            LlamarImprimir
                        End If
                    
                    Else
                        If CargarDatosTemporal(nTabla, cadSelect) Then
                            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                            
                            cadNombreRPT = "rBodListDiferencia.rpt"
                            cadTitulo = "Listado de Diferencia Consumo/Producido"
                            
                            LlamarImprimir
                        End If
                    End If
                Else ' resumen = true
                    If CargarDatosTemporal2(nTabla, cadSelect) Then
                        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                        
                        cadNombreRPT = "rBodListDiferenciaRes.rpt"
                        cadTitulo = "Diferencia Consumo/Producido Resumido"
                        
                        LlamarImprimir
                    End If
                End If
        End Select
        
    End If

End Sub

Private Sub cmdAceptar_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String


    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
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
    
    
    vSQL = ""
    If txtcodigo(20).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(20).Text, "N")
    If txtcodigo(21).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(21).Text, "N")
    
    
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

    If txtcodigo(14).Text <> "" Then vSQL = vSQL & " and variedades.codvarie >= " & DBSet(txtcodigo(14).Text, "N")
    If txtcodigo(15).Text <> "" Then vSQL = vSQL & " and variedades.codvarie <= " & DBSet(txtcodigo(15).Text, "N")

    'D/H DEPOSITO
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".coddeposito}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHDeposito=""") Then Exit Sub
    End If
    
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
        
    If Not AnyadirAFormula(cadFormula, "{grupopro.codgrupo} = 6") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{grupopro.codgrupo} = 6") Then Exit Sub
    
    nTabla = "(rhisfruta INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
    nTabla = "(" & nTabla & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
    nTabla = "(" & nTabla & ") INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
    If CargarTablaTemporal(nTabla, cadSelect) Then
            
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme("tmpinformes", "codusu = " & vUsu.Codigo) Then 'nTabla, cadSelect) Then
            
            ConSubInforme = False
            
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            Select Case OpcionListado
                Case 0
                    '[Monica]24/10/2011: Personalizacion de lso informes de bodega por Quatretonda
                    indRPT = 81 ' Informe de Entradas Bodega
                    
                    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                    
                    cadNombreRPT = nomDocu ' rBodInfEntradas.rpt
                    cadTitulo = "Informe de Entradas Bodega"
                    If Check2.Value Then cadNombreRPT = Replace(cadNombreRPT, "Entradas.rpt", "EntradasVariedad.rpt") 'rBodInfEntradasVariedad.rpt
                Case 1
                    If Check1.Value = 0 Then
                        ' no saltamos pagina por socio
                        cadNombreRPT = "rBodExtEntradas.rpt"
                        ConSubInforme = True
                    Else
                        If vParamAplic.Cooperativa = 3 Then
                            indRPT = 36 ' extracto de entradas por socio Bodega
                            
                            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                            
                            cadNombreRPT = nomDocu ' rBodExtSocEntradas.rpt
                        Else
                            ' saltamos pagina por socio
                            cadNombreRPT = "rBodExtSocEntradas.rpt"
                        End If
                    End If
                    cadTitulo = "Extracto Entradas por Socio/Variedad"
            End Select
            
            LlamarImprimir
        End If
    End If
End Sub

Private Sub cmdAceptarBonif_Click()
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
Dim TipoPrec As Byte ' 0 anticipos
                     ' 1 liquidaciones
Dim b As Boolean
Dim Sql2 As String

Dim Seccion As Integer
Dim vTipo As Byte

        InicializarVbles
        
        'Añadir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        
        If vParamAplic.SeccionBodega = "" Then
            MsgBox "No tiene asignada la seccion de bodega en parámetros. Revise", vbExclamation
            Exit Sub
        Else
            Seccion = CInt(vParamAplic.SeccionBodega)
        End If
            
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(27).Text)
        cHasta = Trim(txtcodigo(28).Text)
        nDesde = txtNombre(27).Text
        nHasta = txtNombre(28).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(25).Text)
        cHasta = Trim(txtcodigo(26).Text)
        nDesde = txtNombre(25).Text
        nHasta = txtNombre(26).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(25).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(25).Text, "N")
        If txtcodigo(26).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(26).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(29).Text)
        cHasta = Trim(txtcodigo(30).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
            
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & Seccion) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & Seccion) Then Exit Sub
        
        'Socio que no sea tercero
        If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        
        
        'sólo entradas distintas de VENTA CAMPO
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1") Then Exit Sub
        
        'sólo las entradas que no tengan bonificacion especial
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.esbonifespecial} = 0") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.esbonifespecial} = 0") Then Exit Sub
        
        nTabla = "((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo = 6 " ' grupo SOLO puede ser 6=bodega
        
        
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        If HayRegParaInforme(nTabla, cadSelect) Then
            If CalcularGradoBonificado(nTabla, cadSelect, Me.Pb1) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (3)
            End If
        End If

End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        Select Case OpcionListado
            Case 7 ' grabacion del porcentaje bonificado
                PonerFoco txtcodigo(27)
        
            Case 8
                PonerFoco txtcodigo(32)
        
            Case 9
                PonerFoco txtcodigo(39)
        
            Case Else
                PonerFoco txtcodigo(12)
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
    
    ConSubInforme = False

    For H = 0 To 27
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 32 To 33
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    'Ocultar todos los Frames de Formulario
    FrameEntradasCampo.visible = False
    FrameListadoConsumo.visible = False
    FrameAsignacionPrecios.visible = False
    FrameBonificacion.visible = False
    FrameAutoconsumo.visible = False
    FrameDiarioFrasRetirada.visible = False
    FrameAsignacionPreciosABN.visible = False
    '###Descomentar
'    CommitConexion
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    
    
    Select Case OpcionListado
        Case 0 'Informe de entradas
            FrameEntradaBasculaVisible True, H, W
            indFrame = 1
            Tabla = "rhisfruta"
            Label3.Caption = "Informe de Entradas Bodega"
        Case 1 'Extracto de entradas
            FrameEntradaBasculaVisible True, H, W
            indFrame = 1
            Tabla = "rhisfruta"
            Label3.Caption = "Extracto Entradas por Socio/Variedad"
        Case 2 'Listado de consumo
            FrameListadoConsumoVisible True, H, W
            indFrame = 1
            Tabla = "rbodalbaran"
            Me.Option1(0).Value = True
            Me.Check4.visible = False
            Me.Check4.Enabled = False
            
        Case 3 'Asignacion de precios masiva en albaranes de retirada de vino / aceite
            If vParamAplic.Cooperativa = 1 Then
                FrameAsignacionPreciosABNVisible True, H, W
            Else
                FrameAsignacionPreciosVisible True, H, W
                
                Me.CmdAcepAsigPrec.Enabled = True
                Me.CmdAcepAsigPrec.visible = True
                Me.CmdAcepGastosLiq.Enabled = False
                Me.CmdAcepGastosLiq.visible = False
            End If
            indFrame = 1
            Tabla = "rbodalbaran"
        Case 4 'Reparto de Gastos de liquidacion de bodega
            FrameAsignacionPreciosVisible True, H, W
            indFrame = 1
            Tabla = "rhisfruta"
            Label4.Caption = "Reparto Gastos Liquidación Bodega"
            Label2(33).Caption = "Importe Gastos"
            Label2(34).visible = False
            txtcodigo(23).Enabled = False
            txtcodigo(23).visible = False
            Me.CmdAcepAsigPrec.Enabled = False
            Me.CmdAcepAsigPrec.visible = False
            Me.CmdAcepGastosLiq.Enabled = True
            Me.CmdAcepGastosLiq.visible = True
            Me.CmdAcepGastosLiq.Top = 4425
        Case 5 'Reparto de Gastos de liquidacion de almazara
            FrameAsignacionPreciosVisible True, H, W
            indFrame = 1
            Tabla = "rhisfruta"
            Label4.Caption = "Reparto Gastos Liquidación Almazara"
            Label2(33).Caption = "Importe Gastos"
            Label2(34).visible = False
            txtcodigo(23).Enabled = False
            txtcodigo(23).visible = False
            Me.CmdAcepAsigPrec.Enabled = False
            Me.CmdAcepAsigPrec.visible = False
            Me.CmdAcepGastosLiq.Enabled = True
            Me.CmdAcepGastosLiq.visible = True
            Me.CmdAcepGastosLiq.Top = 4425
            
        Case 6 'Listado diferencia de consumo/entradas
            FrameListadoConsumoVisible True, H, W
            indFrame = 1
            Tabla = "rbodalbaran"
            Me.Option1(0).Value = True
            Frame1.visible = False
            Frame1.Enabled = False
            Label1.Caption = "Listado Diferencia Consumo/Producido"
            Me.Check4.visible = True
            Me.Check4.Enabled = True
            
        Case 7 'Cálculo de Bonificacion
            FrameBonificacionVisible True, H, W
            indFrame = 1
            Tabla = "rhisfruta"
        
        Case 8 ' informe de liquidacion de autoconsumo
            FrameAutoconsumoVisible True, H, W
            indFrame = 1
            Tabla = "rfactsoc_albaran"
        
        Case 9 ' diario de facturacion de retirada
            FrameDiarioFrasRetiradaVisible True, H, W
            indFrame = 1
            Tabla = "rbodfacturas"
            
            '[Monica]31/03/2014: la cooperativa solo se pide si es ABN
            Frame2.visible = (vParamAplic.Cooperativa = 1)
            Frame2.Enabled = (vParamAplic.Cooperativa = 1)
            
        
    End Select
    
    Check1.visible = (OpcionListado = 1)
    Check1.Enabled = (OpcionListado = 1)
    
    Check2.visible = (OpcionListado = 0)
    Check2.Enabled = (OpcionListado = 0)
    
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

Private Sub frmDep_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") ' codigo de clase
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
        Case 0, 1  ' Depósito
            AbrirFrmDeposito (Index)
        
        Case 6, 7 ' Clase
            AbrirFrmClase (Index)
        
        Case 22 ' Clase
            AbrirFrmClase (Index + 24)
        Case 25 ' Clase
            AbrirFrmClase (Index + 22)
            
        
        Case 10, 11, 12, 13, 32, 33 'SOCIOS
            AbrirFrmSocios (Index)
        
        Case 2, 3  'SOCIOS
            AbrirFrmSocios (Index + 2)
        
        Case 23, 24 'socios
            AbrirFrmSocios (Index + 18)
        
        Case 26, 27  'SOCIOS
            AbrirFrmSocios (Index + 22)
            
        Case 8, 9, 14, 15 'VARIEDADES
            AbrirFrmVariedad (Index)
    
        Case 4, 5 'VARIEDADES
            AbrirFrmVariedad (Index + 14)
    
        Case 18 ' cooperativa
            AbrirFrmCooperativa (Index)
    
        Case 19 ' cooperativa
            AbrirFrmCooperativa (Index + 6)
            
    
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
        Case 2, 3
            indice = Index + 14
        Case 9
            indice = 24
        Case 6
            indice = 31
        Case 4, 5
            indice = Index - 2
        Case 10, 11
            indice = Index + 27
        Case 12, 13
            indice = Index + 38
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
            Case 0: KEYBusqueda KeyAscii, 0 'deposito desde
            Case 1: KEYBusqueda KeyAscii, 1 'deposito hasta
            
            Case 6: KEYFecha KeyAscii, 0 'fecha entrada
            Case 7: KEYFecha KeyAscii, 1 'fecha entrada
            
            Case 12: KEYBusqueda KeyAscii, 12 'socio desde
            Case 13: KEYBusqueda KeyAscii, 13 'socio hasta
            
            Case 14: KEYBusqueda KeyAscii, 14 'variedad desde
            Case 15: KEYBusqueda KeyAscii, 15 'variedad hasta
            
            Case 20: KEYBusqueda KeyAscii, 6 'clase desde
            Case 21: KEYBusqueda KeyAscii, 7 'clase hasta
        
            Case 16: KEYFecha KeyAscii, 2 'fecha albaran
            Case 17: KEYFecha KeyAscii, 3 'fecha albaran
            
            Case 10: KEYBusqueda KeyAscii, 10 'socio desde
            Case 11: KEYBusqueda KeyAscii, 11 'socio hasta
        
            Case 8: KEYBusqueda KeyAscii, 8 'variedad desde
            Case 9: KEYBusqueda KeyAscii, 9 'variedad hasta
        
            Case 27: KEYBusqueda KeyAscii, 16 'socio desde
            Case 28: KEYBusqueda KeyAscii, 17 'socio hasta
        
            Case 25: KEYBusqueda KeyAscii, 20 'clase desde
            Case 26: KEYBusqueda KeyAscii, 21 'clase hasta
            
            Case 29: KEYFecha KeyAscii, 7 'fecha entrada
            Case 30: KEYFecha KeyAscii, 8 'fecha entrada
        
            Case 32: KEYBusqueda KeyAscii, 32 'socio desde
            Case 33: KEYBusqueda KeyAscii, 33 'socio hasta
        
            Case 24: KEYFecha KeyAscii, 9 'fecha desde
            Case 31: KEYFecha KeyAscii, 6 'fecha hasta
        
            Case 36: KEYBusqueda KeyAscii, 36 'cooperativa
        
            Case 4: KEYBusqueda KeyAscii, 2 'socio desde
            Case 5: KEYBusqueda KeyAscii, 3 'socio hasta
        
            Case 18: KEYBusqueda KeyAscii, 4 'variedad desde
            Case 19: KEYBusqueda KeyAscii, 5 'variedad hasta
        
            Case 2: KEYFecha KeyAscii, 4 'fecha entrada
            Case 3: KEYFecha KeyAscii, 5 'fecha entrada
        
            Case 41: KEYBusqueda KeyAscii, 23 'socio desde
            Case 42: KEYBusqueda KeyAscii, 24 'socio hasta
            Case 37: KEYFecha KeyAscii, 10 'fecha factura
            Case 38: KEYFecha KeyAscii, 11 'fecha factura
            Case 43: KEYBusqueda KeyAscii, 19 'cooperativa
        
            '[Monica]08/03/2016: nueva asignacion de precios ABN
            ' asignacion de precios de albaranes de retirada abn
            Case 48: KEYBusqueda KeyAscii, 26 'socio desde
            Case 49: KEYBusqueda KeyAscii, 27 'socio hasta
        
            Case 46: KEYBusqueda KeyAscii, 22 'clase desde
            Case 47: KEYBusqueda KeyAscii, 25 'clase hasta
        
            Case 50: KEYFecha KeyAscii, 12 'fecha entrada
            Case 51: KEYFecha KeyAscii, 13 'fecha entrada
            
            
        
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

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'DEPOSITO
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rdeposito", "nomdeposito", "coddeposito", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000")
        
        Case 20, 21, 25, 26, 46, 47 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
    
        Case 4, 5, 10, 11, 12, 13, 4, 5, 27, 28, 32, 33, 41, 42, 48, 49 'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            
        Case 2, 3, 6, 7, 16, 17, 2, 3, 29, 30, 24, 31, 37, 38, 50, 51 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
            
        Case 8, 9, 14, 15, 18, 19 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
        
        Case 22, 23 ' 22 precio de venta para asignacion de precios de albaranes de retirada
                    ' 23 precio de venta de excedido
            If OpcionListado = 3 Then
                PonerFormatoDecimal txtcodigo(Index), 8
            Else
                If PonerFormatoDecimal(txtcodigo(Index), 3) Then
                    If OpcionListado = 4 Then Me.CmdAcepGastosLiq.SetFocus
                End If
            End If
    
        Case 36, 43 ' COOPERATIVA
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rcoope", "nomcoope", "codcoope", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000")
    
        Case 39, 40 ' facturas
            PonerFormatoEntero txtcodigo(Index)
    
    End Select
End Sub

Private Sub FrameEntradaBasculaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameEntradasCampo.visible = visible
    If visible = True Then
        Me.FrameEntradasCampo.Top = -90
        Me.FrameEntradasCampo.Left = 0
        Me.FrameEntradasCampo.Height = 6855
        Me.FrameEntradasCampo.Width = 6615
        W = Me.FrameEntradasCampo.Width
        H = Me.FrameEntradasCampo.Height
    End If
End Sub

Private Sub FrameListadoConsumoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameListadoConsumo.visible = visible
    If visible = True Then
        Me.FrameListadoConsumo.Top = -90
        Me.FrameListadoConsumo.Left = 0
        Me.FrameListadoConsumo.Height = 5175
        Me.FrameListadoConsumo.Width = 6615
        W = Me.FrameListadoConsumo.Width
        H = Me.FrameListadoConsumo.Height
    End If
End Sub

Private Sub FrameBonificacionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameBonificacion.visible = visible
    If visible = True Then
        Me.FrameBonificacion.Top = -90
        Me.FrameBonificacion.Left = 0
        Me.FrameBonificacion.Height = 5190
        Me.FrameBonificacion.Width = 6615
        W = Me.FrameBonificacion.Width
        H = Me.FrameBonificacion.Height
    End If
End Sub


Private Sub FrameAutoconsumoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameAutoconsumo.visible = visible
    If visible = True Then
        Me.FrameAutoconsumo.Top = -90
        Me.FrameAutoconsumo.Left = 0
        Me.FrameAutoconsumo.Height = 5190
        Me.FrameAutoconsumo.Width = 6615
        W = Me.FrameAutoconsumo.Width
        H = Me.FrameAutoconsumo.Height
    End If
End Sub


Private Sub FrameDiarioFrasRetiradaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameDiarioFrasRetirada.visible = visible
    If visible = True Then
        Me.FrameDiarioFrasRetirada.Top = -90
        Me.FrameDiarioFrasRetirada.Left = 0
        Me.FrameDiarioFrasRetirada.Height = 5175
        Me.FrameDiarioFrasRetirada.Width = 6615
        W = Me.FrameDiarioFrasRetirada.Width
        H = Me.FrameDiarioFrasRetirada.Height
    End If
End Sub



Private Sub FrameAsignacionPreciosABNVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameAsignacionPreciosABN.visible = visible
    If visible = True Then
        Me.FrameAsignacionPreciosABN.Top = -90
        Me.FrameAsignacionPreciosABN.Left = 0
        Me.FrameAsignacionPreciosABN.Height = 5175
        Me.FrameAsignacionPreciosABN.Width = 6615
        W = Me.FrameAsignacionPreciosABN.Width
        H = Me.FrameAsignacionPreciosABN.Height
    End If
End Sub





Private Sub FrameAsignacionPreciosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameAsignacionPrecios.visible = visible
    If visible = True Then
        Me.FrameAsignacionPrecios.Top = -90
        Me.FrameAsignacionPrecios.Left = 0
        Me.FrameAsignacionPrecios.Height = 5175
        Me.FrameAsignacionPrecios.Width = 6615
        W = Me.FrameAsignacionPrecios.Width
        H = Me.FrameAsignacionPrecios.Height
    End If
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
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = True
        .Opcion = 0
        .Show vbModal
    End With
End Sub


Private Sub AbrirFrmDeposito(indice As Integer)
    indCodigo = indice
    Set frmDep = New frmManDepositos
    frmDep.Caption = "Depósitos"
    frmDep.DatosADevolverBusqueda = "0|1|"
    frmDep.Show vbModal
    Set frmDep = Nothing
End Sub


Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmClase(indice As Integer)
    indCodigo = indice + 14
    
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtcodigo(indice).Text
    
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmCooperativa(indice As Integer)
    indCodigo = indice + 18
    Set frmCoop = New frmManCoope
    frmCoop.DatosADevolverBusqueda = "0|1|"
    frmCoop.Show vbModal
    Set frmCoop = Nothing
End Sub

Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub


Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = CadParam
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
        .Opcion = 0
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
        Sql = "select distinct rhisfruta.numalbar  from " & cTabla & " where " & cWhere
    Else
        Sql = "select distinct rhisfruta.numalbar  from " & cTabla
    End If
    
    Sql1 = "select " & vUsu.Codigo & ", rhisfruta.numalbar, 0 from rhisfruta where numalbar in (" & Sql & ")"
        
    
    Sql2 = "insert into tmpinformes (codusu, codigo1, importe1) " & Sql1
    conn.Execute Sql2
    
    CargarTemporal2 = True
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


Private Function ProductoCampo(campo As String) As String
Dim Sql As String

    ProductoCampo = ""
    
    Sql = "select variedades.codprodu from rcampos inner join variedades on rcampos.codvarie = variedades.codvarie "
    Sql = Sql & " where rcampos.codcampo = " & DBSet(campo, "N")
    
    ProductoCampo = DevuelveValor(Sql)

End Function

Private Function CargarDatosTemporalABN(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Porcen As Currency
Dim Grado As Currency
    
    On Error GoTo eCargarTemporal
    
    CargarDatosTemporalABN = False

    conn.Execute "delete from tmpalmazara where codusu = " & vUsu.Codigo
    
    
    Sql = "insert into tmpalmazara (codusu, codsocio, codvarie, entradas, cantidad) "
    Sql = Sql & "select " & vUsu.Codigo & ", rhisfruta.codsocio, variedades.codclase, 'Entradas', sum(round(rhisfruta.kilosnet * rhisfruta.prestimado / 100,0)) cantidad "
    Sql = Sql & " from rhisfruta, variedades "
    Sql = Sql & " where rhisfruta.codvarie = variedades.codvarie and "
    Sql = Sql & " variedades.codprodu in (select codprodu from productos where codgrupo = 5) "
    
    If cWhere <> "" Then Sql = Sql & " and " & Replace(Replace(Replace(cWhere, "rbodalbaran_variedad", "rhisfruta"), "rbodalbaran", "rhisfruta"), "fechaalb", "fecalbar")
    
    Sql = Sql & " group by 1, 2, 3, 4 "
    Sql = Sql & " union  "
    Sql = Sql & " select " & vUsu.Codigo & ", codsocio, variedades.codclase, 'Salidas', sum(rbodalbaran_variedad.cantidad) cantidad "
    Sql = Sql & " from rbodalbaran, rbodalbaran_variedad, variedades  "
    Sql = Sql & " where  rbodalbaran.numalbar = rbodalbaran_variedad.numalbar   "
    Sql = Sql & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
    
    If cWhere <> "" Then Sql = Sql & " and " & cWhere
    
    Sql = Sql & " group by 1, 2, 3, 4 "
    Sql = Sql & " order by 1, 2, 3, 4"


    conn.Execute Sql

    ' una vez insertado en la tabla temporal grabamos la tabla de tmpinformes
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
' select salidas.codsocio, salidas.codvarie, salidas.cantidad salen, if(entradas.cantidad is null, 0, entradas.cantidad) entran, round(salidas.cantidad - entradas.cantidad,0) diferencia
'   from tmp_almazara salidas left join tmp_almazara entradas on salidas.codsocio = entradas.codsocio and salidas.codvarie = entradas.codvarie
'      where entradas.Entradas = 'Entradas' and
'            salidas.Entradas = 'Salidas' and round(salidas.cantidad - entradas.cantidad,0) > 0
' order by salidas.codsocio, salidas.codvarie
    
    Sql = "insert into tmpinformes (codusu, importe1, importe2, importe3, importe4, importe5) "
    Sql = Sql & "select salidas.codusu, salidas.codsocio, salidas.codvarie, salidas.cantidad salen, if(entradas.cantidad is null, 0, entradas.cantidad) entran, round(salidas.cantidad - if(entradas.cantidad is null, 0, entradas.cantidad),0) diferencia "
    Sql = Sql & "  from tmpalmazara salidas left join tmpalmazara entradas on salidas.codsocio = entradas.codsocio and salidas.codvarie = entradas.codvarie "
    Sql = Sql & "   and salidas.codusu = entradas.codusu  and "
    Sql = Sql & "   entradas.entradas = 'Entradas' "
    Sql = Sql & " where salidas.codusu = " & vUsu.Codigo & " and round(salidas.cantidad - if(entradas.cantidad is null, 0, entradas.cantidad),0) > 0"
    Sql = Sql & " and salidas.entradas = 'Salidas' "
    Sql = Sql & "  order by salidas.codsocio, salidas.codvarie"
    
    conn.Execute Sql

    CargarDatosTemporalABN = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description

End Function






Private Function CargarDatosTemporal(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Porcen As Currency
Dim Grado As Currency
    
    On Error GoTo eCargarTemporal
    
    CargarDatosTemporal = False


' No funciona hacer un alias sobre una tabla temporal
'    conn.Execute " DROP TABLE IF EXISTS tmpalmazara;"
'
'    Sql = "CREATE TEMPORARY TABLE `tmp_almazara` ( "
'    Sql = Sql & "`codsocio` int(7) ,"
'    Sql = Sql & "`codvarie` int(7) ,"
'    Sql = Sql & "`entradas` varchar(8) ,"
'    Sql = Sql & "`cantidad` int(7)) "
'
'    conn.Execute Sql


    conn.Execute "delete from tmpalmazara where codusu = " & vUsu.Codigo
    
    
'create table tmpalmazara
'   select  rhisfruta.codsocio, rhisfruta.codvarie, "Entradas", sum(round(rhisfruta.kilosnet * rhisfruta.prestimado / 100,0)) cantidad
'           From rhisfruta
'          where rhisfruta.codvarie in (60,61)
'       group by 1, 2, 3
'   Union
'   select codsocio, rbodalbaran_variedad.codvarie, "Salidas", sum(rbodalbaran_variedad.cantidad) cantidad
'           From rbodalbaran, rbodalbaran_variedad
'           where rbodalbaran_variedad.codvarie in (60,61) and
'           rbodalbaran.numalbar = rbodalbaran_variedad.numalbar
'       group by 1, 2, 3
'       order by 1, 2, 3
'
    
    Sql = "insert into tmpalmazara (codusu, codsocio, codvarie, entradas, cantidad) "
    Sql = Sql & "select " & vUsu.Codigo & ", rhisfruta.codsocio, rhisfruta.codvarie, 'Entradas', sum(round(rhisfruta.kilosnet * rhisfruta.prestimado / 100,0)) cantidad "
    Sql = Sql & " from rhisfruta, variedades "
    Sql = Sql & " where rhisfruta.codvarie = variedades.codvarie and "
    Sql = Sql & " variedades.codprodu in (select codprodu from productos where codgrupo = 5) "
    
    If cWhere <> "" Then Sql = Sql & " and " & Replace(Replace(Replace(cWhere, "rbodalbaran_variedad", "rhisfruta"), "rbodalbaran", "rhisfruta"), "fechaalb", "fecalbar")
    
    Sql = Sql & " group by 1, 2, 3, 4 "
    Sql = Sql & " union  "
    Sql = Sql & " select " & vUsu.Codigo & ", codsocio, rbodalbaran_variedad.codvarie, 'Salidas', sum(rbodalbaran_variedad.cantidad) cantidad "
    Sql = Sql & " from rbodalbaran, rbodalbaran_variedad "
    Sql = Sql & " where  rbodalbaran.numalbar = rbodalbaran_variedad.numalbar   "
    
    If cWhere <> "" Then Sql = Sql & " and " & cWhere
    
    Sql = Sql & " group by 1, 2, 3, 4 "
    Sql = Sql & " order by 1, 2, 3, 4"


    conn.Execute Sql

    ' una vez insertado en la tabla temporal grabamos la tabla de tmpinformes
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
' select salidas.codsocio, salidas.codvarie, salidas.cantidad salen, if(entradas.cantidad is null, 0, entradas.cantidad) entran, round(salidas.cantidad - entradas.cantidad,0) diferencia
'   from tmp_almazara salidas left join tmp_almazara entradas on salidas.codsocio = entradas.codsocio and salidas.codvarie = entradas.codvarie
'      where entradas.Entradas = 'Entradas' and
'            salidas.Entradas = 'Salidas' and round(salidas.cantidad - entradas.cantidad,0) > 0
' order by salidas.codsocio, salidas.codvarie
    
    Sql = "insert into tmpinformes (codusu, importe1, importe2, importe3, importe4, importe5) "
    Sql = Sql & "select salidas.codusu, salidas.codsocio, salidas.codvarie, salidas.cantidad salen, if(entradas.cantidad is null, 0, entradas.cantidad) entran, round(salidas.cantidad - if(entradas.cantidad is null, 0, entradas.cantidad),0) diferencia "
    Sql = Sql & "  from tmpalmazara salidas left join tmpalmazara entradas on salidas.codsocio = entradas.codsocio and salidas.codvarie = entradas.codvarie "
    Sql = Sql & "   and salidas.codusu = entradas.codusu  and "
    Sql = Sql & "   entradas.entradas = 'Entradas' "
    Sql = Sql & " where salidas.codusu = " & vUsu.Codigo & " and round(salidas.cantidad - if(entradas.cantidad is null, 0, entradas.cantidad),0) > 0"
    Sql = Sql & " and salidas.entradas = 'Salidas' "
    Sql = Sql & "  order by salidas.codsocio, salidas.codvarie"
    
    conn.Execute Sql

    CargarDatosTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description

End Function


Private Function CargarTablaTemporal(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Porcen As Currency
Dim Grado As Currency
    
    On Error GoTo eCargarTemporal
    
    CargarTablaTemporal = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rhisfruta.* FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    While Not Rs.EOF
        Sql1 = "select porcentaje from rbonifica_lineas where codvarie = " & DBSet(Rs!codvarie, "N")
        Sql1 = Sql1 & " and desdegrado <= " & DBSet(Rs!PrEstimado, "N")
        Sql1 = Sql1 & " and " & DBSet(Rs!PrEstimado, "N") & " <= hastagrado "
        
' he cambiado esto por los recordset siguientes
'        Porcen = DevuelveValor(Sql1)
        
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RS1.EOF Then
            Porcen = DBLet(RS1.Fields(0).Value, "N")
            Grado = DBLet(Rs!PrEstimado, "N")
        Else
            'cogemos el registro con el hasta mayor para coger el porcentaje
            Porcen = 0
            Grado = DBLet(Rs!PrEstimado, "N")
            
            Sql2 = "select * from rbonifica_lineas "
            Sql2 = Sql2 & " where codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " and hastagrado = (select max(hastagrado) from rbonifica_lineas"
            Sql2 = Sql2 & " where codvarie = " & DBSet(Rs!codvarie, "N") & ")"
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs2.EOF Then
                Porcen = DBLet(Rs2!Porcentaje, "N")
                Grado = DBLet(Rs2!hastagrado, "N")
            End If
            Set Rs2 = Nothing
            
        End If
        
                                                'Variedad,Albaran,Grado,Porcentaje
        Sql1 = "insert into tmpinformes (codusu, codigo1, importe1,importe2, porcen1) values ("
        Sql1 = Sql1 & vUsu.Codigo & ","
        Sql1 = Sql1 & DBSet(Rs!codvarie, "N") & ","
        Sql1 = Sql1 & DBSet(Rs!numalbar, "N") & ","
        Sql1 = Sql1 & DBSet(Grado, "N") & ","
        Sql1 = Sql1 & DBSet(Porcen, "N") & ")"
        
        conn.Execute Sql1
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    
    CargarTablaTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function




Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    campo = "pGroup" & numGrupo & "="
    nomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Socio"
            CadParam = CadParam & campo & "{rbodalbaran.codsocio}" & "|"
'            If numGrupo = 1 Then
'                cadParam = cadParam & nomCampo & "|"
'            End If
            numParam = numParam + 1
            
        Case "Variedad"
            CadParam = CadParam & campo & "{rbodalbaran_variedad.codvarie}" & "|"
            numParam = numParam + 1
    End Select

End Function



Private Function ProcesarCambios(nTabla As String, cadSelect As String) As Boolean
Dim vSQL As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim Importe As String
Dim Sql As String
Dim Sql1 As String

Dim Albaran As Long
Dim Linea As Long

Dim Codigiva As String

    On Error GoTo eProcesarCambios


    ProcesarCambios = False
    
    conn.BeginTrans

    If cadSelect = "" Then cadSelect = "(1=1)"
    
    nTabla = QuitarCaracterACadena(nTabla, "{")
    nTabla = QuitarCaracterACadena(nTabla, "}")

    If cadSelect <> "" Then
        cadSelect = QuitarCaracterACadena(cadSelect, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
        cadSelect = QuitarCaracterACadena(cadSelect, "_1")
    End If

    vSQL = "select rbodalbaran_variedad.* from " & nTabla
    If cadSelect <> "" Then vSQL = vSQL & " where " & cadSelect

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
        Importe = CalcularImporte(CStr(Rs!cantidad), txtcodigo(22).Text, CStr(Rs!dtolinea), 0, 0, 0)
    
        Sql = "update rbodalbaran_variedad set precioar = " & DBSet(txtcodigo(22).Text, "N")
        Sql = Sql & ",importel = " & DBSet(Importe, "N")
        Sql = Sql & " where numalbar = " & DBSet(Rs!numalbar, "N")
        Sql = Sql & " and numlinea = " & DBSet(Rs!numlinea, "N")
        
        conn.Execute Sql
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing
        
    ' una vez está todo calculado con respecto al precio1 de  los litros consumidos,
    ' calculamos y ajustamos los precios de los litros excedidos en negativo al precio 1
    ' y en positivo al precio2
    
    If CargarDatosTemporal(nTabla, cadSelect) Then
        ' tenemos cargada la tabla temporal con los datos de entrada y de salida,
        ' procesamos la tabla temporal para grabar las lineas en negativo (precio1) y en positivo
        ' (precio2) sobre el ultimo albaran de cada socio a regular
        
        Sql = "select importe1 codsocio, importe2 codvarie, importe3 entran, importe4 salen, importe5 diferencia "
        Sql = Sql & " from tmpinformes "
        Sql = Sql & " where codusu = " & vUsu.Codigo
        Sql = Sql & " order by importe1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            If DBLet(Rs!Diferencia, "N") > 0 And (txtcodigo(22).Text <> txtcodigo(23).Text) Then ' el socio produce mas que consume
                Sql1 = "select max(rbodalbaran.numalbar) "
                Sql1 = Sql1 & " from rbodalbaran INNER JOIN rbodalbaran_variedad ON rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
                Sql1 = Sql1 & " where codsocio = " & DBSet(Rs!Codsocio, "N") & " and "
                Sql1 = Sql1 & " rbodalbaran_variedad.codvarie = " & DBSet(Rs!codvarie, "N")
                If cadSelect <> "" Then Sql1 = Sql1 & " and " & cadSelect
                
                Albaran = DevuelveValor(Sql1)
                
                If Albaran <> 0 Then
                    Sql1 = "select max(numlinea) from rbodalbaran_variedad where numalbar = " & DBSet(Albaran, "N")
                    Linea = DevuelveValor(Sql1)
                    
                    If Linea <> 0 Then
                        Codigiva = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", Rs!codvarie, "N")
                    
                        ' en negativo
                        Linea = Linea + 1
                        
                        Importe = CalcularImporte(CStr(Rs!Diferencia * (-1)), txtcodigo(22).Text, CStr(0), 0, 0, 0)
                        
                        Sql1 = "insert into rbodalbaran_variedad (numalbar, numlinea, codvarie, unidades, cantidad, "
                        Sql1 = Sql1 & "precioar, dtolinea, importel, ampliaci, codigiva) values ("
                        Sql1 = Sql1 & DBSet(Albaran, "N") & ","
                        Sql1 = Sql1 & DBSet(Linea, "N") & ","
                        Sql1 = Sql1 & DBSet(Rs!codvarie, "N") & ","
                        Sql1 = Sql1 & DBSet(Rs!Diferencia * (-1), "N") & ","
                        Sql1 = Sql1 & DBSet(Rs!Diferencia * (-1), "N") & ","
                        Sql1 = Sql1 & DBSet(txtcodigo(22).Text, "N") & ","
                        Sql1 = Sql1 & "0," ' la filas añadidas no tienen dtolinea
                        Sql1 = Sql1 & DBSet(Importe, "N") & ","
                        Sql1 = Sql1 & ValorNulo & ","
                        Sql1 = Sql1 & DBSet(Codigiva, "N") & ")"
                        
                        conn.Execute Sql1
                        
                        ' en positivo
                        Linea = Linea + 1
                        
                        Importe = CalcularImporte(CStr(Rs!Diferencia), txtcodigo(23).Text, CStr(0), 0, 0, 0)
                        
                        Sql1 = "insert into rbodalbaran_variedad (numalbar, numlinea, codvarie, unidades, cantidad, "
                        Sql1 = Sql1 & "precioar, dtolinea, importel, ampliaci, codigiva) values ("
                        Sql1 = Sql1 & DBSet(Albaran, "N") & ","
                        Sql1 = Sql1 & DBSet(Linea + 1, "N") & ","
                        Sql1 = Sql1 & DBSet(Rs!codvarie, "N") & ","
                        Sql1 = Sql1 & DBSet(Rs!Diferencia, "N") & ","
                        Sql1 = Sql1 & DBSet(Rs!Diferencia, "N") & ","
                        Sql1 = Sql1 & DBSet(txtcodigo(23).Text, "N") & ","
                        Sql1 = Sql1 & "0," ' la filas añadidas no tienen dtolinea
                        Sql1 = Sql1 & DBSet(Importe, "N") & ","
                        Sql1 = Sql1 & ValorNulo & ","
                        Sql1 = Sql1 & DBSet(Codigiva, "N") & ")"
                        
                        conn.Execute Sql1
                         
                    End If
                End If
        
            End If
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
        
    End If
    
       
    conn.CommitTrans
    ProcesarCambios = True
    Exit Function
    
eProcesarCambios:
    conn.RollbackTrans
    MuestraError Err.Number, "Procesar Cambios", Err.Description
End Function

Private Function ProcesarCambiosABN(nTabla As String, cadSelect As String) As Boolean
Dim vSQL As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim Importe As String
Dim Sql As String
Dim Sql1 As String

Dim Albaran As Long
Dim Linea As Long

Dim Codigiva As String
Dim PrecioVenta As Currency
Dim RS1 As ADODB.Recordset
Dim Diferencia As Currency
Dim VarieAnt As String
Dim cantidad As Currency
Dim SqlInsert As String
Dim SqlValues As String
Dim PrecioVta As Currency

    On Error GoTo eProcesarCambios


    ProcesarCambiosABN = False
    
    conn.BeginTrans

    If cadSelect = "" Then cadSelect = "(1=1)"
    
    nTabla = QuitarCaracterACadena(nTabla, "{")
    nTabla = QuitarCaracterACadena(nTabla, "}")

    If cadSelect <> "" Then
        cadSelect = QuitarCaracterACadena(cadSelect, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
        cadSelect = QuitarCaracterACadena(cadSelect, "_1")
    End If

    vSQL = "select variedades.eurdesta, rbodalbaran_variedad.* from " & nTabla
    If cadSelect <> "" Then vSQL = vSQL & " where " & cadSelect

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
        PrecioVenta = DBLet(Rs!EurDesta, "N")
    
        Importe = CalcularImporte(CStr(Rs!cantidad), CStr(PrecioVenta), CStr(Rs!dtolinea), 0, 0, 0)
    
        Sql = "update rbodalbaran_variedad set precioar = " & DBSet(PrecioVenta, "N")
        Sql = Sql & ",importel = " & DBSet(Importe, "N")
        Sql = Sql & " where numalbar = " & DBSet(Rs!numalbar, "N")
        Sql = Sql & " and numlinea = " & DBSet(Rs!numlinea, "N")
        
        conn.Execute Sql
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing
        
    ' una vez está todo calculado con respecto al precio1 de  los litros consumidos,
    ' calculamos y ajustamos los precios de los litros excedidos en negativo al precio 1
    ' y en positivo al precio2
    
    If CargarDatosTemporalABN(nTabla, cadSelect) Then
        ' tenemos cargada la tabla temporal con los datos de entrada y de salida,
        ' procesamos la tabla temporal para grabar las lineas en negativo (precio1) y en positivo
        ' (precio2) sobre el ultimo albaran de cada socio a regular
        
        Sql = "select importe1 codsocio, importe2 codvarie, importe3 entran, importe4 salen, importe5 diferencia "
        Sql = Sql & " from tmpinformes "
        Sql = Sql & " where codusu = " & vUsu.Codigo & " and importe5 > 0 " ' el socio produce más que consume
        Sql = Sql & " order by importe1, importe2 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Sql1 = "delete from tmpinformes2 where codusu = " & DBSet(vUsu.Codigo, "N")
            conn.Execute Sql1
            
            SqlInsert = "insert into tmpinformes2 (codusu, codigo1, importe1, importe2, precio1, precio2) values "
            
            SqlValues = ""
            
            
            Sql1 = "select rbodalbaran_variedad.*, variedades.eurdesta, variedades.eurecole "
            Sql1 = Sql1 & " from (rbodalbaran INNER JOIN rbodalbaran_variedad ON rbodalbaran.numalbar = rbodalbaran_variedad.numalbar) "
            Sql1 = Sql1 & " INNER JOIN variedades ON rbodalbaran_variedad.codvarie = variedades.codvarie "
            Sql1 = Sql1 & " where codsocio = " & DBSet(Rs!Codsocio, "N") & " and "
            Sql1 = Sql1 & " variedades.codclase = " & DBSet(Rs!codvarie, "N")
            If cadSelect <> "" Then Sql1 = Sql1 & " and " & cadSelect
            Sql1 = Sql1 & " order by rbodalbaran.fechaalb desc, rbodalbaran.numalbar desc, rbodalbaran_variedad.numlinea desc "
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Diferencia = DBSet(Rs!Diferencia, "N")
            
            VarieAnt = RS1!codvarie
            cantidad = 0
            
            While Not RS1.EOF And Diferencia > 0
                    
                If Diferencia < DBLet(RS1!cantidad, "N") Then
                    cantidad = Diferencia
                    Diferencia = 0
                Else
                    cantidad = DBLet(RS1!cantidad, "N")
                    Diferencia = Diferencia - DBLet(RS1!cantidad, "N")
                End If
                
                
                
                Sql = "select count(*) from tmpinformes2 where codusu = " & vUsu.Codigo & " and importe1 = " & DBSet(RS1!codvarie, "N")
                Sql = Sql & " and codigo1 = " & DBSet(Rs!Codsocio, "N")
                
                If TotalRegistros(Sql) = 0 Then
                    SqlValues = "(" & vUsu.Codigo & "," & DBSet(Rs!Codsocio, "N") & "," & DBSet(RS1!codvarie, "N") & "," & DBSet(cantidad, "N") & ","
                    SqlValues = SqlValues & DBSet(RS1!EurDesta, "N") & "," & DBSet(RS1!eurecole, "N") & ")"
                
                    conn.Execute SqlInsert & SqlValues
                Else
                    SqlValues = "update tmpinformes2 set importe2 = importe2 + " & DBSet(cantidad, "N")
                    SqlValues = SqlValues & " where codusu = " & vUsu.Codigo & " and codigo1 = " & DBSet(Rs!Codsocio, "N")
                    SqlValues = SqlValues & " and importe1 = " & DBSet(VarieAnt, "N")
                    
                    conn.Execute SqlValues
                End If
            
            
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            
            
            Sql1 = "select max(rbodalbaran.numalbar) "
            Sql1 = Sql1 & " from rbodalbaran INNER JOIN rbodalbaran_variedad ON rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
            Sql1 = Sql1 & " where codsocio = " & DBSet(Rs!Codsocio, "N")
            If cadSelect <> "" Then Sql1 = Sql1 & " and " & cadSelect
            
            Albaran = DevuelveValor(Sql1)
            
            If Albaran <> 0 Then
                Sql1 = "select max(numlinea) from rbodalbaran_variedad where numalbar = " & DBSet(Albaran, "N")
                Linea = DevuelveValor(Sql1)
                
                If Linea <> 0 Then
                   Sql1 = "select * from tmpinformes2 where codusu = " & vUsu.Codigo & " and codigo1 = " & DBSet(Rs!Codsocio, "N")
                   Sql1 = Sql1 & " order by importe1 "
                   
                    Set RS1 = New ADODB.Recordset
                    RS1.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not RS1.EOF
                        Codigiva = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", RS1!importe1, "N")
                    
                    
                        ' en negativo
                        Linea = Linea + 1
                        
                        Importe = CalcularImporte(CStr(RS1!importe2 * (-1)), CStr(DBLet(RS1!Precio1, "N")), CStr(0), 0, 0, 0)
                        
                        Sql1 = "insert into rbodalbaran_variedad (numalbar, numlinea, codvarie, unidades, cantidad, "
                        Sql1 = Sql1 & "precioar, dtolinea, importel, ampliaci, codigiva) values ("
                        Sql1 = Sql1 & DBSet(Albaran, "N") & ","
                        Sql1 = Sql1 & DBSet(Linea, "N") & ","
                        Sql1 = Sql1 & DBSet(RS1!importe1, "N") & "," 'variedad
                        Sql1 = Sql1 & DBSet(RS1!importe2 * (-1), "N") & "," 'cantidad
                        Sql1 = Sql1 & DBSet(RS1!importe2 * (-1), "N") & ","
                        Sql1 = Sql1 & DBSet(RS1!Precio1, "N") & ","
                        Sql1 = Sql1 & "0," ' la filas añadidas no tienen dtolinea
                        Sql1 = Sql1 & DBSet(Importe, "N") & ","
                        Sql1 = Sql1 & "'Regularización de Precios'" & ","
                        Sql1 = Sql1 & DBSet(Codigiva, "N") & ")"
                        
                        conn.Execute Sql1
                        
                        ' en positivo
                        Linea = Linea + 1
                        
                        Importe = CalcularImporte(CStr(RS1!importe2), CStr(DBLet(RS1!Precio2, "N")), CStr(0), 0, 0, 0)
                        
                        Sql1 = "insert into rbodalbaran_variedad (numalbar, numlinea, codvarie, unidades, cantidad, "
                        Sql1 = Sql1 & "precioar, dtolinea, importel, ampliaci, codigiva) values ("
                        Sql1 = Sql1 & DBSet(Albaran, "N") & ","
                        Sql1 = Sql1 & DBSet(Linea + 1, "N") & ","
                        Sql1 = Sql1 & DBSet(RS1!importe1, "N") & "," 'variedad
                        Sql1 = Sql1 & DBSet(RS1!importe2, "N") & "," 'cantidad
                        Sql1 = Sql1 & DBSet(RS1!importe2, "N") & ","
                        Sql1 = Sql1 & DBSet(RS1!Precio2, "N") & ","
                        Sql1 = Sql1 & "0," ' la filas añadidas no tienen dtolinea
                        Sql1 = Sql1 & DBSet(Importe, "N") & ","
                        Sql1 = Sql1 & "'Regularización de Precios'" & ","
                        Sql1 = Sql1 & DBSet(Codigiva, "N") & ")"
                        
                        conn.Execute Sql1
                    
                    
                        RS1.MoveNext
                    Wend
                    
                    Set RS1 = Nothing
                
                End If
            End If
        
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
        
    End If
    
       
    conn.CommitTrans
    ProcesarCambiosABN = True
    Exit Function
    
eProcesarCambios:
    conn.RollbackTrans
    MuestraError Err.Number, "Procesar Cambios", Err.Description
End Function










Private Function ProcesarRepartoGastos(nTabla As String, cadSelect As String, Imporgasto As String) As Boolean
Dim vSQL As String
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim KilosTotal As Long
Dim Importe As Currency
Dim TotalImporte As Currency
Dim NroAlbaran As String
Dim NumF As String

    On Error GoTo eProcesarRepartoGastos

    ProcesarRepartoGastos = False
    
    conn.BeginTrans

    ' eliminamos de los gastos de albaranes todos los registros correspondientes al gasto de liquidacion
    ' de los albaranes seleccionados
    If OpcionListado = 4 Then
        vSQL = "delete from rhisfruta_gastos where codgasto = " & vParamAplic.CodGastoBOD
    Else
        vSQL = "delete from rhisfruta_gastos where codgasto = " & vParamAplic.CodGastoAlmz
    End If
    vSQL = vSQL & " and " & cadSelect
    conn.Execute vSQL

    ' obtenemos los kilos totales sobre los que se va a hacer el prorrateo
    vSQL = "select sum(kilosnet) from rhisfruta " & " where " & cadSelect
    KilosTotal = DevuelveValor(vSQL) ' kilos totales sobre los que prorratearemos los gastos
    
    'proceso de prorrateo de gastos
    vSQL = "select numalbar, kilosnet from  rhisfruta "
    If cadSelect <> "" Then vSQL = vSQL & " where " & cadSelect
    
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    TotalImporte = 0
    NroAlbaran = ""
    While Not Rs.EOF
        NroAlbaran = CStr(DBLet(Rs!numalbar, "N"))
        
        Importe = Round2(Rs!KilosNet * Imporgasto / KilosTotal, 2)
        TotalImporte = TotalImporte + Importe
        
        NumF = SugerirCodigoSiguienteStr("rhisfruta_gastos", "numlinea", "numalbar = " & DBSet(Rs!numalbar, "N"))
        
        Sql2 = "insert into rhisfruta_gastos (numalbar,numlinea,codgasto,importe) values ("
        
        If OpcionListado = 4 Then
            Sql2 = Sql2 & DBSet(NroAlbaran, "N") & "," & DBSet(NumF, "N") & "," & DBSet(vParamAplic.CodGastoBOD, "N") & ","
        Else
            Sql2 = Sql2 & DBSet(NroAlbaran, "N") & "," & DBSet(NumF, "N") & "," & DBSet(vParamAplic.CodGastoAlmz, "N") & ","
        End If
        Sql2 = Sql2 & DBSet(Importe, "N") & ")"
        
        conn.Execute Sql2
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing
        
    'si hay diferencia de importes por redondeo lo introducimos en el ultimo albaran
    If NroAlbaran <> "" And TotalImporte <> CCur(Imporgasto) Then
        Sql = "update rhisfruta_gastos set importe = importe + " & DBSet(Imporgasto - TotalImporte, "N")
        Sql = Sql & " where numalbar = " & NroAlbaran & " and numlinea = " & NumF
        conn.Execute Sql
    End If
        
    conn.CommitTrans
    ProcesarRepartoGastos = True
    Exit Function
    
eProcesarRepartoGastos:
    conn.RollbackTrans
    MuestraError Err.Number, "Procesar Reparto Gastos", Err.Description
End Function


Private Function CadenaAlbaranes(cTabla As String, cWhere As String) As String
'Devuelve una cadena con los albaranes separados por comas
Dim Sql As String
Dim Cad As String
Dim Rs As ADODB.Recordset


    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select numalbar FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    While Not Rs.EOF
        Cad = Cad & DBLet(Rs!numalbar, "N") & ","
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1)
    End If
    CadenaAlbaranes = Cad
    
End Function

Private Function CargarDatosTemporal2(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Porcen As Currency
Dim Grado As Currency

Dim Entradas As Currency
Dim Salidas As Currency
Dim Diferencia As Currency
    
    On Error GoTo eCargarTemporal
    
    CargarDatosTemporal2 = False

    Label2(55).visible = True

    conn.Execute "delete from tmpalmazara where codusu = " & vUsu.Codigo
    
    Sql = "insert into tmpalmazara (codusu, codsocio, entradas, cantidad) "
    Sql = Sql & "select " & vUsu.Codigo & ", rhisfruta.codsocio,  'Entradas', sum(round(rhisfruta.kilosnet * rhisfruta.prestimado / 100,0)) cantidad "
    Sql = Sql & " from rhisfruta, variedades "
    Sql = Sql & " where rhisfruta.codvarie = variedades.codvarie and "
    Sql = Sql & " variedades.codprodu in (select codprodu from productos where codgrupo = 5) "
    
    If cWhere <> "" Then Sql = Sql & " and " & Replace(Replace(Replace(cWhere, "rbodalbaran_variedad", "rhisfruta"), "rbodalbaran", "rhisfruta"), "fechaalb", "fecalbar")
    
    Sql = Sql & " group by 1, 2, 3 "
    Sql = Sql & " union  "
    Sql = Sql & " select " & vUsu.Codigo & ", codsocio, 'Salidas', sum(rbodalbaran_variedad.cantidad) cantidad "
    Sql = Sql & " from rbodalbaran, rbodalbaran_variedad, variedades "
    Sql = Sql & " where  rbodalbaran.numalbar = rbodalbaran_variedad.numalbar and  "
    Sql = Sql & " rbodalbaran_variedad.codvarie = variedades.codvarie and "
    Sql = Sql & " variedades.codprodu in (select codprodu from productos where codgrupo = 5) "
    
    
    If cWhere <> "" Then Sql = Sql & " and " & cWhere
    
    Sql = Sql & " group by 1, 2, 3 "
    Sql = Sql & " order by 1, 2, 3 "


    conn.Execute Sql

    ' una vez insertado en la tabla temporal grabamos la tabla de tmpinformes

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
                                            'socio,   entradas, salidas,  diferencia
    Sql = "insert into tmpinformes (codusu, importe1, importe2, importe3, importe4)  "
    Sql = Sql & " select " & vUsu.Codigo & ", codsocio, 0,0,0 "
    Sql = Sql & " from tmpalmazara where codusu = " & vUsu.Codigo
    Sql = Sql & " group by codusu, codsocio"
    conn.Execute Sql
    
    Sql = "select importe1 from tmpinformes where codusu = " & vUsu.Codigo
    
    pb2.Max = TotalRegistrosConsulta(Sql)
    pb2.visible = True
    pb2.Value = 0
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgresNew pb2, 1
        DoEvents
    
        Entradas = DevuelveValor("select cantidad from tmpalmazara where codusu = " & vUsu.Codigo & " and entradas = 'Entradas' and codsocio = " & DBSet(Rs!importe1, "N"))
        Salidas = DevuelveValor("select cantidad from tmpalmazara where codusu = " & vUsu.Codigo & " and entradas = 'Salidas' and codsocio = " & DBSet(Rs!importe1, "N"))
        Diferencia = Entradas - Salidas
        
        Sql = "update tmpinformes set importe2 = " & DBSet(Entradas, "N")
        Sql = Sql & ", importe3 = " & DBSet(Salidas, "N")
        Sql = Sql & ", importe4 = " & DBSet(Diferencia, "N")
        Sql = Sql & " where codusu = " & vUsu.Codigo & " and importe1 = " & DBSet(Rs!importe1, "N")
        
        conn.Execute Sql
        
        Rs.MoveNext
    Wend
     
    Set Rs = Nothing

    CargarDatosTemporal2 = True
    Label2(55).visible = False
    pb2.visible = False
    Exit Function
    
eCargarTemporal:
    Label2(55).visible = False
    MuestraError "Cargando Datos Temporal2", Err.Description
End Function

