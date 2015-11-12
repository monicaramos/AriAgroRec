VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPOZListado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8340
   Icon            =   "frmPOZListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameReimpresion 
      Height          =   5220
      Left            =   0
      TabIndex        =   112
      Top             =   0
      Width           =   6675
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   134
         Top             =   1350
         Width           =   2070
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   118
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1755
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1740
         MaxLength       =   7
         TabIndex        =   117
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1365
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   122
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   120
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelReimp 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5340
         TabIndex        =   116
         Top             =   4275
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarReimp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   115
         Top             =   4275
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   126
         Top             =   3765
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   124
         Top             =   3390
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   35
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "Text5"
         Top             =   3765
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   34
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text5"
         Top             =   3390
         Width           =   3675
      End
      Begin VB.Label Label6 
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
         TabIndex        =   133
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   900
         TabIndex        =   132
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   131
         Top             =   1395
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
         Index           =   2
         Left            =   495
         TabIndex        =   130
         Top             =   1125
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   465
         TabIndex        =   129
         Top             =   2115
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   825
         TabIndex        =   128
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   825
         TabIndex        =   127
         Top             =   2775
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1485
         Picture         =   "frmPOZListado.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   2415
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   1485
         Picture         =   "frmPOZListado.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   855
         TabIndex        =   125
         Top             =   3405
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   870
         TabIndex        =   123
         Top             =   3780
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
         Index           =   11
         Left            =   510
         TabIndex        =   121
         Top             =   3165
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1470
         MouseIcon       =   "frmPOZListado.frx":0122
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3765
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1470
         MouseIcon       =   "frmPOZListado.frx":0274
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3390
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   5
         Left            =   3180
         TabIndex        =   119
         Top             =   1110
         Width           =   1815
      End
   End
   Begin VB.Frame FrameReciboConsumo 
      Height          =   6285
      Left            =   0
      TabIndex        =   57
      Top             =   -30
      Width           =   6945
      Begin VB.Frame Frame99 
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   240
         TabIndex        =   74
         Top             =   2850
         Width           =   6375
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   0
            TabIndex        =   169
            Top             =   1590
            Width           =   6525
            Begin VB.TextBox txtcodigo 
               Height          =   285
               Index           =   48
               Left            =   1560
               MaxLength       =   40
               MultiLine       =   -1  'True
               TabIndex        =   170
               Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
               Top             =   210
               Width           =   4725
            End
            Begin VB.Image imgAyuda 
               Height          =   240
               Index           =   2
               Left            =   1230
               MousePointer    =   4  'Icon
               Tag             =   "-1"
               ToolTipText     =   "Ayuda"
               Top             =   210
               Width           =   240
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Concepto"
               ForeColor       =   &H00972E0B&
               Height          =   225
               Index           =   35
               Left            =   330
               TabIndex        =   171
               Top             =   180
               Width           =   690
            End
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   62
            Top             =   150
            Width           =   1005
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   2
            Left            =   4260
            TabIndex        =   63
            Top             =   0
            Width           =   1965
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Recibo"
            Height          =   195
            Index           =   3
            Left            =   4260
            TabIndex        =   64
            Top             =   330
            Width           =   1995
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   1155
            Left            =   270
            TabIndex        =   172
            Top             =   510
            Width           =   3555
            Begin VB.TextBox txtcodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   1290
               MaxLength       =   10
               TabIndex        =   176
               Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
               Top             =   480
               Width           =   1005
            End
            Begin VB.TextBox txtcodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   1290
               MaxLength       =   10
               TabIndex        =   175
               Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
               Top             =   780
               Width           =   1005
            End
            Begin VB.TextBox txtcodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   4
               Left            =   2400
               MaxLength       =   10
               TabIndex        =   174
               Top             =   480
               Width           =   1005
            End
            Begin VB.TextBox txtcodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   5
               Left            =   2400
               MaxLength       =   10
               TabIndex        =   173
               Top             =   780
               Width           =   1005
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Rango Consumo"
               ForeColor       =   &H00972E0B&
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   179
               Top             =   180
               Width           =   1185
            End
            Begin VB.Label Label2 
               Caption         =   "Hasta m3"
               Height          =   195
               Index           =   5
               Left            =   1290
               TabIndex        =   178
               Top             =   300
               Width           =   945
            End
            Begin VB.Label Label2 
               Caption         =   "Precio"
               Height          =   195
               Index           =   7
               Left            =   2430
               TabIndex        =   177
               Top             =   300
               Width           =   945
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recibo"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   3
            Left            =   330
            TabIndex        =   75
            Top             =   -30
            Width           =   1005
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1305
            Picture         =   "frmPOZListado.frx":03C6
            Top             =   150
            Width           =   240
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   61
         Top             =   2370
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   60
         Top             =   1980
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   12
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   59
         Top             =   1380
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   11
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   58
         Top             =   1020
         Width           =   960
      End
      Begin VB.CommandButton CmdAceptarRecCons 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   4530
         TabIndex        =   65
         Top             =   5610
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5640
         TabIndex        =   66
         Top             =   5595
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   540
         TabIndex        =   76
         Top             =   5220
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1530
         Picture         =   "frmPOZListado.frx":0451
         Top             =   2370
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1530
         Picture         =   "frmPOZListado.frx":04DC
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   22
         Left            =   570
         TabIndex        =   73
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   930
         TabIndex        =   72
         Top             =   2370
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   930
         TabIndex        =   71
         Top             =   2010
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Generación de Recibos de Consumo"
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
         Left            =   540
         TabIndex        =   70
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hidrante"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   16
         Left            =   540
         TabIndex        =   69
         Top             =   870
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   15
         Left            =   900
         TabIndex        =   68
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   900
         TabIndex        =   67
         Top             =   1110
         Width           =   465
      End
   End
   Begin VB.Frame FrameReciboTalla 
      Height          =   5085
      Left            =   0
      TabIndex        =   235
      Top             =   0
      Width           =   6945
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen"
         Height          =   255
         Index           =   7
         Left            =   4590
         TabIndex        =   262
         Top             =   1950
         Width           =   1965
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Recibo"
         Height          =   195
         Index           =   6
         Left            =   4590
         TabIndex        =   261
         Top             =   2310
         Width           =   1995
      End
      Begin VB.Frame FrameBonif 
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   765
         Left            =   450
         TabIndex        =   252
         Top             =   3720
         Width           =   6255
         Begin VB.CheckBox Check1 
            Caption         =   "Sólo efectos"
            Height          =   195
            Index           =   8
            Left            =   4980
            TabIndex        =   272
            Top             =   420
            Width           =   1695
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   78
            Left            =   1470
            MaxLength       =   10
            TabIndex        =   244
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   77
            Left            =   3720
            MaxLength       =   10
            TabIndex        =   245
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bonificación"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   71
            Left            =   150
            TabIndex        =   256
            Top             =   390
            Width           =   870
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   1
            Left            =   1140
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   390
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Recargo"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   70
            Left            =   2850
            TabIndex        =   255
            Top             =   390
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   69
            Left            =   2460
            TabIndex        =   254
            Top             =   390
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   68
            Left            =   4710
            TabIndex        =   253
            Top             =   390
            Width           =   120
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5790
         TabIndex        =   147
         Top             =   4455
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   246
         Top             =   4470
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   75
         Left            =   1995
         MaxLength       =   10
         TabIndex        =   239
         Top             =   1455
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   74
         Left            =   1995
         MaxLength       =   10
         TabIndex        =   238
         Top             =   1110
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   73
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   240
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   74
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   237
         Text            =   "Text5"
         Top             =   1110
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   75
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   236
         Text            =   "Text5"
         Top             =   1470
         Width           =   3675
      End
      Begin MSComctlLib.ProgressBar pb4 
         Height          =   255
         Left            =   480
         TabIndex        =   263
         Top             =   4110
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame FrameCuota 
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         Height          =   1605
         Left            =   300
         TabIndex        =   257
         Top             =   2430
         Width           =   6525
         Begin VB.TextBox txtNombre 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   5430
            Locked          =   -1  'True
            TabIndex        =   267
            Text            =   "Text5"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
            Height          =   435
            Index           =   76
            Left            =   1590
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   243
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
            Top             =   1110
            Width           =   4815
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   72
            Left            =   2220
            MaxLength       =   10
            TabIndex        =   241
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   180
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   66
            Left            =   2220
            MaxLength       =   10
            TabIndex        =   242
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Precio Braza"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   74
            Left            =   4470
            TabIndex        =   266
            Top             =   630
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "€/hanegada"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   73
            Left            =   3240
            TabIndex        =   265
            Top             =   630
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "€/hanegada"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   72
            Left            =   3240
            TabIndex        =   264
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   67
            Left            =   270
            TabIndex        =   260
            Top             =   1020
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cuota Amortizacion Canal"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   59
            Left            =   270
            TabIndex        =   259
            Top             =   210
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cuota Talla Ordinaria"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   58
            Left            =   270
            TabIndex        =   258
            Top             =   600
            Width           =   1485
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Comprobando"
         Height          =   195
         Index           =   78
         Left            =   480
         TabIndex        =   273
         Top             =   4440
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   66
         Left            =   1020
         TabIndex        =   251
         Top             =   1110
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   65
         Left            =   1020
         TabIndex        =   250
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   64
         Left            =   540
         TabIndex        =   249
         Top             =   870
         Width           =   405
      End
      Begin VB.Label Label12 
         Caption         =   "Generación Recibos de Tallas"
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
         TabIndex        =   248
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   63
         Left            =   570
         TabIndex        =   247
         Top             =   1980
         Width           =   1005
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   1680
         Picture         =   "frmPOZListado.frx":0567
         Top             =   1980
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1680
         MouseIcon       =   "frmPOZListado.frx":05F2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1680
         MouseIcon       =   "frmPOZListado.frx":0744
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1470
         Width           =   240
      End
   End
   Begin VB.Frame FrameCartaTallas 
      Height          =   3885
      Left            =   30
      TabIndex        =   218
      Top             =   0
      Width           =   6945
      Begin VB.TextBox txtNombre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   268
         Text            =   "Text5"
         Top             =   3270
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   68
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   234
         Text            =   "Text5"
         Top             =   1470
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   67
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   233
         Text            =   "Text5"
         Top             =   1110
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   71
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   223
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
         Top             =   2850
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   70
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   222
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
         Top             =   2430
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   69
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   221
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   67
         Left            =   1995
         MaxLength       =   10
         TabIndex        =   219
         Top             =   1110
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   68
         Left            =   1995
         MaxLength       =   10
         TabIndex        =   220
         Top             =   1455
         Width           =   960
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   4500
         TabIndex        =   224
         Top             =   3210
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5610
         TabIndex        =   225
         Top             =   3195
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "€/hanegada"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   77
         Left            =   3540
         TabIndex        =   271
         Top             =   2460
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "€/hanegada"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   76
         Left            =   3540
         TabIndex        =   270
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio Braza"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   75
         Left            =   570
         TabIndex        =   269
         Top             =   3300
         Width           =   900
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1680
         MouseIcon       =   "frmPOZListado.frx":0896
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1470
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1680
         MouseIcon       =   "frmPOZListado.frx":09E8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuota Talla Ordinaria"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   62
         Left            =   570
         TabIndex        =   232
         Top             =   2850
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuota Amortizacion Canal"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   61
         Left            =   570
         TabIndex        =   231
         Top             =   2460
         Width           =   1815
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1680
         Picture         =   "frmPOZListado.frx":0B3A
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   60
         Left            =   570
         TabIndex        =   230
         Top             =   1980
         Width           =   1005
      End
      Begin VB.Label Label11 
         Caption         =   "Carta de Tallas a Socios"
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
         TabIndex        =   229
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   57
         Left            =   540
         TabIndex        =   228
         Top             =   870
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   56
         Left            =   1020
         TabIndex        =   227
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   55
         Left            =   1020
         TabIndex        =   226
         Top             =   1110
         Width           =   465
      End
   End
   Begin VB.Frame FrameEtiquetasContadores 
      Height          =   3885
      Left            =   0
      TabIndex        =   157
      Top             =   0
      Width           =   6945
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   47
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   160
         Top             =   1920
         Width           =   4770
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   162
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|##,###,##0||"
         Top             =   2850
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   45
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   158
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   1050
         Width           =   4770
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   46
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   159
         Top             =   1470
         Width           =   4770
      End
      Begin VB.CommandButton CmdAceptarEtiq 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4500
         TabIndex        =   164
         Top             =   3210
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelEtiq 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5610
         TabIndex        =   166
         Top             =   3195
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Línea 3"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   34
         Left            =   570
         TabIndex        =   168
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Línea 2"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   39
         Left            =   570
         TabIndex        =   167
         Top             =   1500
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Nro.Etiquetas"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   38
         Left            =   570
         TabIndex        =   165
         Top             =   2550
         Width           =   1050
      End
      Begin VB.Label Label9 
         Caption         =   "Etiquetas Contadores"
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
         TabIndex        =   163
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Línea 1"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   36
         Left            =   570
         TabIndex        =   161
         Top             =   1080
         Width           =   555
      End
   End
   Begin VB.Frame FrameComprobacion 
      Height          =   3885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6945
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5610
         TabIndex        =   6
         Top             =   3195
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   4500
         TabIndex        =   5
         Top             =   3210
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   19
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1455
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   18
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1110
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2325
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1980
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   29
         Left            =   900
         TabIndex        =   13
         Top             =   1110
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   900
         TabIndex        =   12
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hidrante"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   540
         TabIndex        =   11
         Top             =   870
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Informe de Comprobación de Lecturas"
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
         TabIndex        =   10
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   930
         TabIndex        =   9
         Top             =   2010
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   930
         TabIndex        =   8
         Top             =   2370
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   570
         TabIndex        =   7
         Top             =   1800
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1545
         Picture         =   "frmPOZListado.frx":0BC5
         Top             =   2325
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1560
         Picture         =   "frmPOZListado.frx":0C50
         Top             =   1980
         Width           =   240
      End
   End
   Begin VB.Frame FrameRectificacion 
      Height          =   4680
      Left            =   0
      TabIndex        =   184
      Top             =   0
      Width           =   6675
      Begin VB.Frame Frame9 
         Caption         =   "Datos para Selección"
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
         Height          =   1995
         Left            =   240
         TabIndex        =   188
         Top             =   870
         Width           =   6315
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   55
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   192
            Top             =   1350
            Width           =   1065
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   4080
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   197
            Top             =   450
            Width           =   2070
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   52
            Left            =   1620
            MaxLength       =   7
            TabIndex        =   190
            Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
            Top             =   450
            Width           =   1035
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   56
            Left            =   1620
            MaxLength       =   6
            TabIndex        =   191
            Top             =   900
            Width           =   1065
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   56
            Left            =   2775
            Locked          =   -1  'True
            TabIndex        =   189
            Text            =   "Text5"
            Top             =   900
            Width           =   3405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hidrante"
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
            Index           =   7
            Left            =   300
            TabIndex        =   201
            Top             =   1350
            Width           =   615
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
            Index           =   21
            Left            =   300
            TabIndex        =   200
            Top             =   465
            Width           =   870
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
            Index           =   27
            Left            =   300
            TabIndex        =   199
            Top             =   900
            Width           =   375
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   1320
            MouseIcon       =   "frmPOZListado.frx":0CDB
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   900
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Factura"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   28
            Left            =   2850
            TabIndex        =   198
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   51
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   193
         Top             =   3060
         Width           =   1065
      End
      Begin VB.CommandButton CmdAceptarRectif 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   195
         Top             =   3915
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelRectif 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5340
         TabIndex        =   196
         Top             =   3915
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   54
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   194
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3480
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lectura"
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
         Index           =   20
         Left            =   510
         TabIndex        =   187
         Top             =   3060
         Width           =   540
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   1560
         Picture         =   "frmPOZListado.frx":0E2D
         ToolTipText     =   "Buscar fecha"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   22
         Left            =   510
         TabIndex        =   186
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Rectificación de Lecturas"
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
         Left            =   465
         TabIndex        =   185
         Top             =   405
         Width           =   5160
      End
   End
   Begin VB.Frame FrameReciboContador 
      Height          =   7725
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   8235
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   33
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   110
         Top             =   6270
         Width           =   1185
      End
      Begin VB.Frame Frame4 
         Caption         =   "Artículos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2265
         Left            =   270
         TabIndex        =   107
         Top             =   3900
         Width           =   7815
         Begin VB.TextBox txtcodigo 
            Height          =   315
            Index           =   31
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   95
            Text            =   "frmPOZListado.frx":0EB8
            Top             =   1620
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   96
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   1620
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
            Height          =   315
            Index           =   29
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   93
            Text            =   "frmPOZListado.frx":0F01
            Top             =   1260
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   94
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   1260
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
            Height          =   315
            Index           =   27
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   91
            Text            =   "frmPOZListado.frx":0F4A
            Top             =   900
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   28
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   92
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
            Height          =   315
            Index           =   25
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   89
            Text            =   "frmPOZListado.frx":0F93
            Top             =   540
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   90
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   540
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   6420
            TabIndex        =   109
            Top             =   270
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   240
            TabIndex        =   108
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mano Obra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1005
         Left            =   300
         TabIndex        =   104
         Top             =   2670
         Width           =   7785
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   6390
            MaxLength       =   10
            TabIndex        =   88
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
            Height          =   315
            Index           =   20
            Left            =   210
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   87
            Text            =   "frmPOZListado.frx":0FDC
            Top             =   540
            Width           =   6105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   6390
            TabIndex        =   106
            Top             =   270
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   210
            TabIndex        =   105
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   6810
         TabIndex        =   99
         Top             =   7095
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarRecCont 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5700
         TabIndex        =   97
         Top             =   7110
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   85
         Top             =   1440
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   84
         Top             =   1080
         Width           =   960
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3165
         Left            =   210
         TabIndex        =   80
         Top             =   2010
         Width           =   7815
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   86
            Top             =   150
            Width           =   1005
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   5
            Left            =   4290
            TabIndex        =   82
            Top             =   60
            Width           =   1965
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Recibo"
            Height          =   195
            Index           =   4
            Left            =   4290
            TabIndex        =   81
            Top             =   420
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recibo"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   11
            Left            =   330
            TabIndex        =   83
            Top             =   -30
            Width           =   1005
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   1350
            Picture         =   "frmPOZListado.frx":1025
            Top             =   150
            Width           =   240
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "Text5"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "Text5"
         Top             =   1110
         Width           =   3375
      End
      Begin MSComctlLib.ProgressBar Pb3 
         Height          =   255
         Left            =   510
         TabIndex        =   98
         Top             =   6750
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe  Recibo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   12
         Left            =   4890
         TabIndex        =   111
         Top             =   6300
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   30
         Left            =   900
         TabIndex        =   103
         Top             =   1110
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   900
         TabIndex        =   102
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   13
         Left            =   540
         TabIndex        =   101
         Top             =   870
         Width           =   405
      End
      Begin VB.Label Label5 
         Caption         =   "Generación de Recibos Contadores"
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
         TabIndex        =   100
         Top             =   300
         Width           =   5925
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1590
         MouseIcon       =   "frmPOZListado.frx":10B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1590
         MouseIcon       =   "frmPOZListado.frx":1202
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1080
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameFacturasHidrante 
      Height          =   6030
      Left            =   0
      TabIndex        =   135
      Top             =   0
      Width           =   6675
      Begin VB.Frame Frame7 
         Caption         =   "Ordenado por"
         Enabled         =   0   'False
         ForeColor       =   &H00972E0B&
         Height          =   615
         Left            =   480
         TabIndex        =   181
         Top             =   5130
         Width           =   3285
         Begin VB.OptionButton Option3 
            Caption         =   "Nro.Factura"
            Height          =   195
            Left            =   1800
            TabIndex        =   183
            Top             =   270
            Width           =   1305
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Socio"
            Height          =   195
            Left            =   210
            TabIndex        =   182
            Top             =   270
            Value           =   -1  'True
            Width           =   1305
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Resumen Facturación"
         Height          =   285
         Left            =   510
         TabIndex        =   143
         Top             =   4740
         Width           =   2175
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   50
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   141
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3825
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   140
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3450
         Width           =   1050
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   4260
         Width           =   2100
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   40
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   148
         Text            =   "Text5"
         Top             =   1230
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   41
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   146
         Text            =   "Text5"
         Top             =   1605
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   136
         Top             =   1230
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   137
         Top             =   1605
         Width           =   830
      End
      Begin VB.CommandButton cmdAceptarListFact 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4230
         TabIndex        =   144
         Top             =   5355
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelList 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5400
         TabIndex        =   145
         Top             =   5340
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   138
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   139
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1050
      End
      Begin VB.Label Label4 
         Caption         =   "Nro. Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   6
         Left            =   510
         TabIndex        =   180
         Top             =   3180
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   156
         Top             =   4110
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":1354
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":14A6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
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
         Index           =   19
         Left            =   510
         TabIndex        =   155
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   870
         TabIndex        =   154
         Top             =   1620
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   855
         TabIndex        =   153
         Top             =   1245
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1485
         Picture         =   "frmPOZListado.frx":15F8
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   1485
         Picture         =   "frmPOZListado.frx":1683
         ToolTipText     =   "Buscar fecha"
         Top             =   2415
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   825
         TabIndex        =   152
         Top             =   2775
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   825
         TabIndex        =   151
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   8
         Left            =   510
         TabIndex        =   150
         Top             =   2115
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Facturas por Hidrante"
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
         TabIndex        =   149
         Top             =   315
         Width           =   5160
      End
   End
   Begin VB.Frame FrameReciboMantenimiento 
      Height          =   7005
      Left            =   30
      TabIndex        =   26
      Top             =   0
      Width           =   6945
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   36
         Top             =   3570
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   64
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   35
         Top             =   3540
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   63
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   30
         Top             =   1950
         Width           =   1080
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   62
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   29
         Text            =   "      "
         Top             =   1950
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   60
         Left            =   4080
         MaxLength       =   25
         TabIndex        =   34
         Text            =   "      "
         Top             =   3000
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   59
         Left            =   1845
         MaxLength       =   25
         TabIndex        =   33
         Top             =   3000
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   58
         Left            =   4080
         MaxLength       =   6
         TabIndex        =   32
         Text            =   "      "
         Top             =   2460
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   57
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   31
         Top             =   2460
         Width           =   960
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Text5"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text5"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   210
         TabIndex        =   49
         Top             =   3990
         Width           =   6375
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   61
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   40
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   990
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   53
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   39
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   990
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Recibo"
            Height          =   195
            Index           =   0
            Left            =   4290
            TabIndex        =   51
            Top             =   420
            Width           =   1995
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   1
            Left            =   4290
            TabIndex        =   50
            Top             =   60
            Width           =   1965
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   38
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   570
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
            Height          =   435
            Index           =   9
            Left            =   1650
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
            Top             =   1530
            Width           =   4725
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   37
            Top             =   150
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   54
            Left            =   4890
            TabIndex        =   217
            Top             =   1020
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   53
            Left            =   2640
            TabIndex        =   216
            Top             =   1020
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Recargo"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   46
            Left            =   3030
            TabIndex        =   209
            Top             =   1020
            Width           =   615
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   1320
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bonificación"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   37
            Left            =   330
            TabIndex        =   202
            Top             =   1020
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   9
            Left            =   330
            TabIndex        =   54
            Top             =   1440
            Width           =   690
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   1350
            Picture         =   "frmPOZListado.frx":170E
            Top             =   150
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recibo"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   10
            Left            =   330
            TabIndex        =   53
            Top             =   -30
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Euros/Acción"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   6
            Left            =   330
            TabIndex        =   52
            Top             =   570
            Width           =   975
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   27
         Top             =   1080
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   28
         Top             =   1440
         Width           =   960
      End
      Begin VB.CommandButton CmdAceptarRecMto 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4560
         TabIndex        =   42
         Top             =   6450
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   5640
         TabIndex        =   43
         Top             =   6435
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   255
         Left            =   540
         TabIndex        =   44
         Top             =   6090
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   52
         Left            =   930
         TabIndex        =   215
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   51
         Left            =   3270
         TabIndex        =   214
         Top             =   3570
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alta"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   50
         Left            =   540
         TabIndex        =   213
         Top             =   3330
         Width           =   765
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   3810
         Picture         =   "frmPOZListado.frx":1799
         Top             =   3570
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   1560
         Picture         =   "frmPOZListado.frx":1824
         Top             =   3540
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hidrante"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   49
         Left            =   540
         TabIndex        =   212
         Top             =   1740
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   48
         Left            =   900
         TabIndex        =   211
         Top             =   1980
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   47
         Left            =   3270
         TabIndex        =   210
         Top             =   1980
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   45
         Left            =   3270
         TabIndex        =   208
         Top             =   3030
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   44
         Left            =   900
         TabIndex        =   207
         Top             =   3030
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   43
         Left            =   540
         TabIndex        =   206
         Top             =   2790
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   42
         Left            =   3270
         TabIndex        =   205
         Top             =   2490
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   41
         Left            =   900
         TabIndex        =   204
         Top             =   2490
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Polígono"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   40
         Left            =   540
         TabIndex        =   203
         Top             =   2250
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":18AF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":1A01
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Generación de Recibos Mantenimiento"
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
         TabIndex        =   48
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   540
         TabIndex        =   47
         Top             =   870
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   900
         TabIndex        =   46
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   900
         TabIndex        =   45
         Top             =   1110
         Width           =   465
      End
   End
   Begin VB.Frame FrameTomaLectura 
      Height          =   3795
      Left            =   0
      TabIndex        =   14
      Top             =   30
      Width           =   6105
      Begin VB.Frame Frame1 
         Caption         =   "Ordenado por"
         ForeColor       =   &H00972E0B&
         Height          =   1125
         Left            =   450
         TabIndex        =   23
         Top             =   2100
         Width           =   2205
         Begin VB.OptionButton Option1 
            Caption         =   "Nro.Orden"
            Height          =   315
            Index           =   1
            Left            =   300
            TabIndex        =   25
            Top             =   660
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Contador"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   24
            Top             =   330
            Width           =   1725
         End
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   1
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1665
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   0
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1260
         Width           =   960
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   18
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   20
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Listado de Toma de Lectura"
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
         TabIndex        =   22
         Top             =   450
         Width           =   5250
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hidrante"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   21
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   19
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   17
         Top             =   1320
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmPOZListado"
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
    ' 1 .- Listado Toma de Lectura de Contador
    ' 2 .- Listado de Comprobacion de lecturas
    
    ' 3 .- Generación Recibos de Consumo (Facturacion de consumo)
    ' 4 .- Generación Recibos de Mantenimiento (Factura de Mantenimiento)
    ' 5 .- Generacion Recibos de Contadores ( Factura de Contadores )
    
    ' 6 .- Reimpresion de recibos de pozos
    ' 7 .- Listado de consumo por hidrante
    
    ' 8 .- Etiquetas contadores
    ' 9 .- Facturas rectificativas
    
    ' 10.- Listado de tallas, recibos de talla (solo para Escalona)
    ' 11.- Generacion de recibos de talla
    ' 12.- Cálculo de bonificacion de Recibos de Talla (solo para Escalona)
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmSoc  As frmManSocios 'mantenimiento de socios
Attribute frmSoc.VB_VarHelpID = -1
 
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'ayuda de hidrantes por socio
Attribute frmMens.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------




Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String
Dim vSeccion As CSeccion


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




Private Sub Check2_Click()
    Me.Frame7.Enabled = (Check2.Value = 1)
End Sub

Private Sub CmdAceptarRectif_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

Dim Consumo As Long
    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Tabla = "rpozos"

    cadSelect = " rpozos.hidrante = " & DBSet(txtcodigo(55).Text, "T")     ' Hidrante
    
    
    '[Monica]23/09/2011: de momento solo rectifico las facturas de quatretonda
    Select Case vParamAplic.Cooperativa
        '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
        Case 8, 10 ' UTXERA
    
    
        Case 7 ' Quatretonda
            Dim b As Boolean
            
            Consumo = 0
            b = CalculoConsumoHidrante(txtcodigo(55).Text, txtcodigo(51).Text, Consumo)
             
            If b Then
                Check1(2).Value = 1
                Check1(3).Value = 1
                ProcesoFacturacionConsumo Tabla, cadSelect, txtcodigo(54).Text, Consumo, True
            End If
        
        Case Else ' MALLAES
    
    
    End Select

End Sub

Private Sub CmdAceptarEtiq_Click()
Dim campo As String
Dim Tabla As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim Sql As String
Dim Sql2 As String
Dim I As Long

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    'si es listado de CARTAS/eMAIL a socios comprobar que se ha seleccionado
    'una carta para imprimir
    'ETIQUETAS
    cadParam = "|"

    'Nombre fichero .rpt a Imprimir
    nomRPT = "TurPOZEtiqContador.rpt"
    cadTitulo = "Etiquetas de Contadores" '"Etiquetas de Contadores"
    

    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 50 'Impresion de Etiquetas de contadores
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    
    conSubRPT = False
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H Seccion
    '--------------------------------------------
        
    'Parametro Linea 1
    If txtcodigo(45).Text <> "" Then
        cadParam = cadParam & "pLinea1="" " & txtcodigo(45).Text & """|"
    Else
        cadParam = cadParam & "pLinea1=""""|"
    End If
    numParam = numParam + 1
    
    'Parametro Linea 2
    If txtcodigo(46).Text <> "" Then
        cadParam = cadParam & "pLinea2="" " & txtcodigo(46).Text & """|"
    Else
        cadParam = cadParam & "pLinea2=""""|"
    End If
    numParam = numParam + 1
    
    'Parametro Linea 3
    If txtcodigo(47).Text <> "" Then
        cadParam = cadParam & "pLinea3="" " & txtcodigo(47).Text & """|"
    Else
        cadParam = cadParam & "pLinea3=""""|"
    End If
    numParam = numParam + 1
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = ""
    For I = 1 To CLng(txtcodigo(44).Text)
        Sql = Sql & "(" & vUsu.Codigo & "," & I & "),"
    Next I
    
    Sql2 = "insert into tmpinformes (codusu,codigo1) values "
    Sql2 = Sql2 & Mid(Sql, 1, Len(Sql) - 1) ' quitamos la ultima coma
    
    conn.Execute Sql2
    
    cadFormula = ""
    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu}=" & vUsu.Codigo) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{tmpinformes.codusu}=" & vUsu.Codigo) Then Exit Sub
    
    Tabla = "tmpinformes"
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    LlamarImprimir
    

End Sub

Private Sub cmdCancelEtiq_Click()
    Unload Me
End Sub

Private Sub CmdCancelList_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String
Dim Sql3 As String


    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    Select Case Index
        Case 0 ' Listado de Toma de lectura de contador
            'D/H Hidrante
            cDesde = Trim(txtcodigo(0).Text)
            cHasta = Trim(txtcodigo(1).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rpozos.hidrante}"
                TipCod = "T"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHHidrante=""") Then Exit Sub
            End If
            
            cTabla = Tabla
            
            If Me.Option1(0).Value Then
                cadParam = cadParam & "pOrden={rpozos.hidrante}|"
                cadParam = cadParam & "pDescOrden=""Ordenado por Hidrante""|"
                cadParam = cadParam & "pOrden1={rpozos.hidrante}|"
                
            End If
            If Me.Option1(1).Value Then
                cadParam = cadParam & "pOrden={rpozos.nroorden}|"
                cadParam = cadParam & "pDescOrden=""Ordenado por Nro.Orden""|"
                cadParam = cadParam & "pOrden1={rpozos.hidrante}|"
            End If
            numParam = numParam + 3
            
            If Not AnyadirAFormula(cadFormula, "isnull({rpozos.fechabaja})") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{rpozos.fechabaja} is null") Then Exit Sub
            
            
            
            indRPT = 44 'listado de toma de lecturas de pozos
            ConSubInforme = False
            cadTitulo = "Listado de Toma de Lecturas"
        
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = nomDocu

            If HayRegParaInforme(cTabla, cadSelect) Then
                LlamarImprimir
            End If
    
        Case 1  ' opcionlistado = 2 --> informe de comprobacion
            '======== FORMULA  ====================================
            'D/H Hidrante
            cDesde = Trim(txtcodigo(18).Text)
            cHasta = Trim(txtcodigo(19).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rpozos.hidrante}"
                TipCod = "T"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHHidrante=""") Then Exit Sub
            End If
            
            'D/H fecha
            cDesde = Trim(txtcodigo(16).Text)
            cHasta = Trim(txtcodigo(17).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rpozos.fech_act}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If

            If Not AnyadirAFormula(cadFormula, "isnull({rpozos.fechabaja})") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{rpozos.fechabaja} is null") Then Exit Sub


            indRPT = 45
            ConSubInforme = False
            cadTitulo = "Comprobación de Lecturas"
        
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = nomDocu
            
            If vParamAplic.Cooperativa = 7 Then
                If CargarTemporal(Tabla, cadSelect) Then
                    If HayRegParaInforme("tmpinformes", "tmpinformes.codusu = " & vUsu.Codigo) Then
                        cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
                        numParam = numParam + 1
                        ConSubInforme = True
                        LlamarImprimir
                    End If
                End If
            Else
                If HayRegParaInforme(Tabla, cadSelect) Then
                    LlamarImprimir
                End If
            End If
    
        Case 2  ' opcionlistado = 10 --> cartas de tallas
            '======== FORMULA  ====================================
            'D/H Socio
            cDesde = Trim(txtcodigo(67).Text)
            cHasta = Trim(txtcodigo(68).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rrecibpozos.codsocio}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
            End If
            
            'Fecha de Recibo
            If Not AnyadirAFormula(cadFormula, "{rrecibpozos.fecfactu}=date('" & txtcodigo(69).Text & "')") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{rrecibpozos.fecfactu}=" & DBSet(txtcodigo(69).Text, "F")) Then Exit Sub
            
            If Not AnyadirAFormula(cadFormula, "{rrecibpozos.codtipom}='TAL'") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{rrecibpozos.codtipom}='TAL'") Then Exit Sub

            cadParam = cadParam & "pCuoAmor=" & TransformaComasPuntos(ImporteSinFormato(ComprobarCero(txtcodigo(70).Text))) & "|"
            cadParam = cadParam & "pCuoTalla=" & TransformaComasPuntos(ImporteSinFormato(ComprobarCero(txtcodigo(71).Text))) & "|"
            numParam = numParam + 2

            indRPT = 86
            ConSubInforme = False
            cadTitulo = "Carta de tallas a Socios"
        
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = nomDocu
            
            If HayRegParaInforme(Tabla, cadSelect) Then
                ConSubInforme = True
                LlamarImprimir
            End If
    
    
        Case 3 ' opcionlistado = 11 generacion de recibos de talla
               ' opcionlistado = 12 calculo de bonificacion de recibos de talla
            '======== FORMULA  ====================================
            If OpcionListado = 11 Then
                'D/H Socio
                cDesde = Trim(txtcodigo(74).Text)
                cHasta = Trim(txtcodigo(75).Text)
                nDesde = ""
                nHasta = ""
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{rsocios.codsocio}"
                    TipCod = "N"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
                End If
                
                vSQL = ""
                If txtcodigo(74).Text <> "" Then vSQL = vSQL & " and rsocios.codsocio >= " & DBSet(txtcodigo(74).Text, "N")
                If txtcodigo(75).Text <> "" Then vSQL = vSQL & " and rsocios.codsocio <= " & DBSet(txtcodigo(75).Text, "N")
            
            
                '[Monica]19/09/2012: se factura al propietario de los campos
                Tabla = "rcampos INNER JOIN rsocios ON rcampos.codpropiet = rsocios.codsocio "
                Tabla = "(" & Tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
            
                If Not AnyadirAFormula(cadSelect, "{rsituacion.bloqueo} = 0") Then Exit Sub
'[Monica]03/08/2012: he quitado esto pq en la generacion de facturas de talla va solo por campos CORREGIDO
'                If vSQL <> "" And txtcodigo(74).Text = txtcodigo(75).Text Then
'                    Set frmMens = New frmMensajes
'
'                    frmMens.OpcionMensaje = 37
'                    frmMens.cadWhere = vSQL
'                    frmMens.Show vbModal
'
'                    Set frmMens = Nothing
'                End If
                
                ProcesoFacturacionTallaESCALONA Tabla, cadSelect
            Else
                'D/H Socio
                cDesde = Trim(txtcodigo(74).Text)
                cHasta = Trim(txtcodigo(75).Text)
                nDesde = ""
                nHasta = ""
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{rsocios.codsocio}"
                    TipCod = "N"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
                End If
                
                '[Monica]19/09/2012: se actualiza la factura del propietario de los campos
                Tabla = "rcampos INNER JOIN rsocios ON rcampos.codpropiet = rsocios.codsocio "
                Tabla = "(" & Tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
                Tabla = "(" & Tabla & ") INNER JOIN rrecibpozos ON rsocios.codsocio = rrecibpozos.codsocio and rrecibpozos.codtipom = 'TAL' "
                
                If Not AnyadirAFormula(cadSelect, "{rrecibpozos.fecfactu} = " & DBSet(txtcodigo(73).Text, "F")) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rrecibpozos.fecfactu} = date(" & txtcodigo(73).Text & ")") Then Exit Sub
                
                
                If Check1(8).Value Then
                    Sql3 = "{rsocios.codbanco} <> '8888888888' and not {rsocios.codbanco} is null"
                    If Not AnyadirAFormula(cadSelect, Sql3) Then Exit Sub
                    Sql3 = "{rsocios.codbanco} <> '8888888888' and not isnull({rsocios.codbanco})"
                    If Not AnyadirAFormula(cadSelect, Sql3) Then Exit Sub
                End If
                    
                ProcesoFacturacionTallaESCALONA Tabla, cadSelect
            
            End If
    End Select
End Sub

Private Function CargarTemporal(cTabla As String, cwhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Sql2 As String

    On Error GoTo eCargarTemporal
    
    CargarTemporal = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select " & vUsu.Codigo & ", codpozo, sum(consumo), sum(nroacciones) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    Sql = Sql & " group by 1, 2 "
    Sql = Sql & " order by 1, 2"
    
    Sql2 = "insert into tmpinformes (codusu, codigo1, importe1, importe2) "
    Sql2 = Sql2 & Sql
    conn.Execute Sql2
    
    CargarTemporal = True
    Exit Function

eCargarTemporal:
    MuestraError Err.Description, "Cargar Temporal", Err.Description
End Function




Private Sub cmdAceptarListFact_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim CodTipom As String


InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    Select Case Combo1(1).ListIndex
        Case 0 ' todos
        
        Case 1
            Tipos = "{rrecibpozos.codtipom} = 'RCP'"
            If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
            If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
        Case 2
            Tipos = "{rrecibpozos.codtipom} = 'RMP'"
            If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
            If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
        Case 3
            Tipos = "{rrecibpozos.codtipom} = 'RVP'"
            If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
            If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
        Case 4
            Tipos = "{rrecibpozos.codtipom} = 'TAL'"
            If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
            If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
    End Select
     
    
    'D/H Socio
    cDesde = Trim(txtcodigo(40).Text)
    cHasta = Trim(txtcodigo(41).Text)
    nDesde = txtNombre(40).Text
    nHasta = txtNombre(41).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(42).Text)
    cHasta = Trim(txtcodigo(43).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    
    '[Monica]26/08/2011: añadido el nro de factura
    'D/H Nro Factura
    cDesde = Trim(txtcodigo(49).Text)
    cHasta = Trim(txtcodigo(50).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFactura= """) Then Exit Sub
    End If
    
    If HayRegistros(Tabla, cadSelect) Then
        indRPT = 48
        ConSubInforme = False
        cadTitulo = "Facturas por Hidrante"
        
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
          
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
          
        '[Monica]26/08/2011: nuevo report que equivale al resumen de facturas de la facturacion
        If Check2.Value = 1 Then
            If Option2.Value Then
                cadTitulo = "Resumen Facturación por Socio"
                cadNombreRPT = Replace(cadNombreRPT, "RecibHidrante", "ResumFactSocio") ' agrupado por socio
            Else
                cadTitulo = "Resumen Facturación"
                cadNombreRPT = Replace(cadNombreRPT, "RecibHidrante", "ResumFactura") ' agrupado por factura
            End If
            numParam = numParam + 1
        End If
          
        'Nombre fichero .rpt a Imprimir
        
        LlamarImprimir
    End If



End Sub

Private Sub CmdAceptarRecCons_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    ' opcionlistado = 3 --> generacion de recibos de consumo
        
    '======== FORMULA  ====================================
    'D/H Hidrante
    cDesde = Trim(txtcodigo(11).Text)
    cHasta = Trim(txtcodigo(12).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.hidrante}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHHidrante=""") Then Exit Sub
    End If
    
    'D/H fecha
    cDesde = Trim(txtcodigo(13).Text)
    cHasta = Trim(txtcodigo(15).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.fech_act}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
'08/09/2010 : no va a ser el que tenga lectura a cero sino el que no tenga fecha de baja
'    If Not AnyadirAFormula(cadSelect, "{rpozos.lect_act} > 0") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{rpozos.fechabaja} is null") Then Exit Sub
    
    Tabla = Tabla & " INNER JOIN rsocios ON rpozos.codsocio = rsocios.codsocio "
    
    Select Case vParamAplic.Cooperativa
        '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
        Case 8, 10 ' UTXERA
            '[Monica]24/10/2011: dejamos que entren las facturas con lectura 0 solo para actualizar contadores
            '                    ponemos ({rpozos.consumo}) >= 0 antes ({rpozos.consumo}) > 0
            cadSelect = cadSelect & " and {rpozos.fech_act} is not null and {rpozos.lect_act} is not null and ({rpozos.consumo}) >= 0 "
        
            '[Monica]27/08/2012: en escalona dejamos unicamente los socios no bloqueados
            If vParamAplic.Cooperativa = 10 Then
                Tabla = "(" & Tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
            
                If Not AnyadirAFormula(cadSelect, "{rsituacion.bloqueo} = 0") Then Exit Sub
            End If
        
        
            ' un recibo por hidrante
            ProcesoFacturacionConsumoUTXERA Tabla, cadSelect
    
        Case 7 ' Quatretonda
            '[Monica] 11/07/2011: tiene hidrantes que no son contadores y a los que solo se les facturan las acciones
            '                   : por lo tanto quito la condicion de la fecha de lectura actual
            'cadSelect = cadSelect & " and {rpozos.fech_act} is not null and {rpozos.lect_act} is not null "
        
            ProcesoFacturacionConsumo Tabla, cadSelect, txtcodigo(14).Text, 0, False
    
        Case Else ' MALLAES
            cadSelect = cadSelect & " and {rpozos.fech_act} is not null and {rpozos.lect_act} is not null "
        
            ProcesoFacturacionConsumo Tabla, cadSelect, txtcodigo(14).Text, 0, False
    End Select
       
End Sub

Private Sub CmdAceptarRecCont_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


    InicializarVbles

    If Not DatosOk Then Exit Sub


    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtcodigo(23).Text)
    cHasta = Trim(txtcodigo(24).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            Codigo = "{rpozos.codsocio}"
        Else
            Codigo = "{rsocios_pozos.codsocio}"
        End If
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If


'10/06/2011 : facturamos unicamente los hidrantes que no tienen fecha de baja
''09/09/2010 : solo socios que no tengan fecha de baja
'    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{rpozos.fechabaja} is null") Then Exit Sub
    
    '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Tabla = "rpozos INNER JOIN rsocios ON rsocios_pozos.codsocio = rsocios.codsocio "
    Else
        Tabla = "(rsocios_pozos INNER JOIN rsocios ON rsocios_pozos.codsocio = rsocios.codsocio) INNER JOIN rpozos ON rsocios.codsocio = rpozos.codsocio "
    End If
    
    ProcesoFacturacionContadores Tabla, cadSelect

End Sub

Private Sub CmdAceptarRecMto_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String


    InicializarVbles

    If Not DatosOk Then Exit Sub


    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtcodigo(6).Text)
    cHasta = Trim(txtcodigo(7).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            Codigo = "{rpozos.codsocio}"
        Else
            Codigo = "{rsocios_pozos.codsocio}"
        End If
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If

    vSQL = ""
    If txtcodigo(6).Text <> "" Then vSQL = vSQL & " and rpozos.codsocio >= " & DBSet(txtcodigo(6).Text, "N")
    If txtcodigo(7).Text <> "" Then vSQL = vSQL & " and rpozos.codsocio <= " & DBSet(txtcodigo(7).Text, "N")

    'D/H hidrante
    cDesde = Trim(txtcodigo(62).Text)
    cHasta = Trim(txtcodigo(63).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.hidrante}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHHidrante=""") Then Exit Sub
    End If

    'D/H Poligono
    cDesde = Trim(txtcodigo(57).Text)
    cHasta = Trim(txtcodigo(58).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.poligono}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPoligono=""") Then Exit Sub
    End If

    'D/H Parcela
    cDesde = Trim(txtcodigo(59).Text)
    cHasta = Trim(txtcodigo(60).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.parcelas}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHParcela=""") Then Exit Sub
    End If

    'D/H fecha
    cDesde = Trim(txtcodigo(64).Text)
    cHasta = Trim(txtcodigo(65).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.fechaalta}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If



    ' opcionlistado = 4 --> generacion de recibos de mantenimiento

'09/09/2010 : solo socios que no tengan fecha de baja
'    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{rpozos.fechabaja} is null") Then Exit Sub
    
    '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Tabla = "rpozos INNER JOIN rsocios ON rpozos.codsocio = rsocios.codsocio "
        
        If vParamAplic.Cooperativa = 10 Then
            Tabla = "(" & Tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
    
            If Not AnyadirAFormula(cadSelect, "{rsituacion.bloqueo} = 0") Then Exit Sub
            
        End If
    Else
        Tabla = "rsocios_pozos INNER JOIN rsocios ON rsocios_pozos.codsocio = rsocios.codsocio "
        Tabla = "(" & Tabla & ") INNER JOIN rpozos ON rsocios_pozos.codsocio = rpozos.codsocio "
    End If

    '[Monica]08/05/2012: solo para Utxera y Escalona pq en turis se va a rsocios_pozos
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        If vSQL <> "" And txtcodigo(6).Text = txtcodigo(7).Text Then
            Set frmMens = New frmMensajes
        
            frmMens.OpcionMensaje = 37
            frmMens.cadWhere = vSQL
            frmMens.Show vbModal
        
            Set frmMens = Nothing
        End If
    End If

    Select Case vParamAplic.Cooperativa
        '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
        Case 8 ' UTXERA
            ProcesoFacturacionMantenimientoUTXERA Tabla, cadSelect
    
        Case 10 ' ESCALONA
            ProcesoFacturacionMantenimientoESCALONA Tabla, cadSelect
    
        Case Else
            ProcesoFacturacionMantenimiento Tabla, cadSelect
    End Select


End Sub

Private Sub cmdAceptarReimp_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim CodTipom As String


InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    Tipos = Mid(Combo1(0).Text, 1, 3)
    CodTipom = Tipos
    
    If Tipos = "" Then
        MsgBox "Debe seleccionar al menos un tipo de factura.", vbExclamation
        Exit Sub
    Else
        ' quitamos la ultima coma
        Tipos = "{rrecibpozos.codtipom} = '" & Tipos & "'"
        If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
    End If
    
    'D/H Socio
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
    
    'D/H Nro de Factura
    cDesde = Trim(txtcodigo(38).Text)
    cHasta = Trim(txtcodigo(39).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rrecibpozos.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(36).Text)
    cHasta = Trim(txtcodigo(37).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    If HayRegistros(Tabla, cadSelect) Then
        Select Case CodTipom
            Case "RCP"
                indRPT = 46 'Impresion de Recibo de Consumo
                If vParamAplic.Cooperativa = 7 Or vParamAplic.Cooperativa = 1 Then
                    ConSubInforme = True
                Else
                    ConSubInforme = False
                End If
                cadTitulo = "Reimpresión de Recibos Consumo"
            Case "RMP"
                indRPT = 47
                ConSubInforme = True
                cadTitulo = "Reimpresión de Recibos Mantenimiento"
            Case "RVP"
                indRPT = 47
                ConSubInforme = False
                cadTitulo = "Reimpresión de Recibos Contadores"
            Case "TAL"
                indRPT = 47
                ConSubInforme = True
                cadTitulo = "Reimpresión de Recibos Talla"
        End Select
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
          
        If CodTipom = "TAL" Then nomDocu = Replace(nomDocu, "Mto.", "Tal.")
  
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
          
        'Nombre fichero .rpt a Imprimir
        
        LlamarImprimir
        
        If frmVisReport.EstaImpreso Then
'            ActualizarRegistros "rfactsoc", cadSelect
        End If
    End If

End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub CmdCancelRectif_Click()
    Unload Me
End Sub

Private Sub cmdCancelReimp_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
Dim NRegs As Long

    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 1 ' Listado de Toma de Lectura
                PonerFoco txtcodigo(0)
            Case 2  ' Listado de comprobacion de lecturas
                PonerFoco txtcodigo(18)
            Case 3 ' generacion de facturas de consumo
                PonerFoco txtcodigo(11)
            Case 4 ' generacion de facturas de mantenimiento
                PonerFoco txtcodigo(6)
            Case 5 ' generacion de facturas de contadores
                PonerFoco txtcodigo(23)
            Case 6 ' reimpresion de recibos
                PonerFoco txtcodigo(38)
            Case 7 ' informe de facturas por hidrante
                PonerFoco txtcodigo(40)
            Case 8 ' etiquetas contadores
                PonerFoco txtcodigo(45)
                
                txtcodigo(45).Text = "AGUA CON CUPO XXXM3/HG/MES"
                txtcodigo(46).Text = "DIA:"
                txtcodigo(47).Text = "LECTURA:"
                
                NRegs = DevuelveValor("select count(*) from rpozos")
                txtcodigo(44).Text = Format(NRegs, "###,###,##0")
                
            Case 9 ' rectificacion de facturas
                txtcodigo(54).Text = Format(Now, "dd/mm/yyyy")
                PonerFoco txtcodigo(52)
                
            Case 10 ' informe de tallas (recibos de mantenimiento de Escalona)
                PonerFoco txtcodigo(67)
                
            Case 11, 12 'recibos y bonificacion de talla
                PonerFoco txtcodigo(74)
                
                If OpcionListado = 12 Then ConexionConta

                txtcodigo(73).Text = Format(Now, "dd/mm/yyyy")
                
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
    
    
    For H = 0 To 7
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 9 To 13
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 0 To imgAyuda.Count - 1
        imgAyuda(H).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next H
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    Me.FrameTomaLectura.visible = False
    Me.FrameComprobacion.visible = False
    Me.FrameReciboMantenimiento.visible = False
    Me.FrameReciboConsumo.visible = False
    Me.FrameReciboContador.visible = False
    Me.FrameFacturasHidrante.visible = False
    Me.FrameEtiquetasContadores.visible = False
    Me.FrameReimpresion.visible = False
    Me.FrameRectificacion.visible = False
    Me.FrameCartaTallas.visible = False
    Me.FrameReciboTalla.visible = False
    
    
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
        'LISTADOS DE MANTENIMIENTOS BASICOS
        '---------------------
        Case 1 ' Informe de Toma de Lectura
            FrameTomaLecturaVisible True, H, W
            indFrame = 0
            Tabla = "rpozos"
            Me.Option1(0).Value = True
            
        Case 2 ' Informe de Comprobacion de lecturas
            FrameComprobacionVisible True, H, W
            indFrame = 0
            Tabla = "rpozos"
            Label7.Caption = "Informe de Comprobación de Lecturas"
            Me.Pb1.visible = False
        
        Case 3 ' generacion de recibos de consumo
            FrameReciboConsumoVisible True, H, W
            indFrame = 0
            Tabla = "rpozos"
            txtcodigo(14).Text = Format(Now, "dd/mm/yyyy")
            
           
            
            Frame6.Enabled = (vParamAplic.Cooperativa = 7 Or vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            Frame6.visible = (vParamAplic.Cooperativa = 7 Or vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            Frame5.Enabled = Not (vParamAplic.Cooperativa = 7)
            Frame5.visible = Not (vParamAplic.Cooperativa = 7)
            
            If vParamAplic.Cooperativa = 7 Then
                Frame6.Top = 510
            End If
            
            
            '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                txtcodigo(2).Text = Format(DevuelveValor("select hastametcub1 from rtipopozos where codpozo = 1"), "0000000")
                txtcodigo(3).Text = Format(DevuelveValor("select hastametcub2 from rtipopozos where codpozo = 1"), "0000000")
                txtcodigo(4).Text = Format(DevuelveValor("select precio1 from rtipopozos where codpozo = 1"), "###,##0.0000")
                txtcodigo(5).Text = Format(DevuelveValor("select precio2 from rtipopozos where codpozo = 1"), "###,##0.0000")
            Else
                txtcodigo(2).Text = Format(vParamAplic.Consumo1POZ, "0000000")
                txtcodigo(3).Text = Format(vParamAplic.Consumo2POZ, "0000000")
                txtcodigo(4).Text = Format(vParamAplic.Precio1POZ, "###,##0.00")
                txtcodigo(5).Text = Format(vParamAplic.Precio2POZ, "###,##0.00")
            End If
            
            Me.Pb1.visible = False
        
        Case 4 ' Generacion de recibos de mantenimiento
            FrameReciboMantenimientoVisible True, H, W
            indFrame = 0
            Tabla = "rsocios_pozos"
            txtcodigo(10).Text = Format(Now, "dd/mm/yyyy")
            Me.Pb2.visible = False
            
            'Si es Escalona el concepto tiene que caber en textcsb33(40 posiciones)
            If vParamAplic.Cooperativa = 10 Then
                txtcodigo(9).MaxLength = 40
                txtcodigo(8).Text = Format(DevuelveValor("select imporcuotahda from rtipopozos where codpozo = 1"), "###,##0.0000")
                Label2(6).Caption = "Euros/Hanegada"
                Check1(0).Value = 1
                Check1(1).Value = 1
            End If
            
        Case 5 ' Generacion de recibos de contadores
            FrameReciboContadorVisible True, H, W
            indFrame = 0
            Tabla = "rsocios_pozos"
            txtcodigo(22).Text = Format(Now, "dd/mm/yyyy")
            Me.Pb3.visible = False
        
        Case 6 ' Reimpresion de recibos de pozos
            FrameReimpresionVisible True, H, W
            Tabla = "rrecibpozos"
            CargaCombo
            Combo1(0).ListIndex = 0
            
        Case 7 ' Informe de recibos por hidrante
            FrameFacturasHidranteVisible True, H, W
            Tabla = "rrecibpozos"
            CargaCombo
            Combo1(1).ListIndex = 0
        
        Case 8 ' Etiquetas contadores
            FrameEtiquetasContadoresVisible True, H, W
            Tabla = "tmpinformes"
        
        Case 9 ' Rectificacion de lecturas
            FrameRectificacionVisible True, H, W
            Tabla = "rrecibpozos"
            CargaCombo
            Combo1(2).ListIndex = 0
        
        Case 10 ' Informe de Tallas recibos de Mto (solo visible para Escalona)
            FrameCartaTallasVisible True, H, W
            indFrame = 0
            Tabla = "rrecibpozos"
        
        Case 11 ' Generacion de recibos de talla (solo para Escalona)
            FrameReciboTallaVisible True, H, W
            indFrame = 0
            Tabla = "rrecibpozos"
            Me.pb4.visible = False
            Check1(6).Value = 1
            Check1(7).Value = 1
        
        Case 12 ' Calculo de bonificacion de recibos de talla
            FrameReciboTallaVisible True, H, W
            indFrame = 0
            Tabla = "rrecibpozos"
            Me.pb4.visible = False
            Check1(6).Value = 1
            Check1(7).Value = 1
        
    End Select
    'Esto se consigue poniendo el cancel en el opcion k corresponda
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
    txtcodigo(CByte(imgFecha(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {rpozos.hidrante} in (" & CadenaSeleccion & ")"
        Sql2 = " {rpozos.hidrante} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {rpozos.hidrante} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Sólo podemos poner un porcentaje de bonificación o un porcentaje de" & vbCrLf & _
                      "recargo pero no ambos a la vez. " & vbCrLf & vbCrLf
                                            
        Case 1
           ' "____________________________________________________________"
            vCadena = "Sólo podemos poner un porcentaje de bonificación o un porcentaje de" & vbCrLf & _
                      "recargo pero no ambos a la vez. " & vbCrLf & vbCrLf
    
        Case 2
           ' "____________________________________________________________"
            vCadena = "Concepto que se imprime en el recibo en caso de que tenga valor." & vbCrLf & _
                      "" & vbCrLf & vbCrLf
                                            
    
    
    
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"

End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 ' Socios
            AbrirFrmSocios (Index + 23)
        Case 2, 3 ' Socios
            AbrirFrmSocios (Index + 32)
        Case 4, 5 ' Socios
            AbrirFrmSocios (Index + 36)
        Case 6, 7  'Socios
            AbrirFrmSocios (Index)
        Case 9 'socios
            AbrirFrmSocios (Index + 47)
        Case 10, 11 ' socios
            AbrirFrmSocios (Index + 57)
        Case 12, 13
            AbrirFrmSocios (Index + 62)
    End Select
    PonerFoco txtcodigo(indCodigo)
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
        Case 0
            indice = 14
        Case 1
            indice = 10
        Case 0, 2, 3
            indice = Index + 14
        Case 4
            indice = 22
        Case 5
            indice = 13
        Case 6
            indice = 15
        Case 7, 8
            indice = Index + 29
        Case 9, 10
            indice = Index + 33
        Case 11
            indice = Index + 43
        Case 12, 13
            indice = Index + 52
        Case 14
            indice = 73
        Case 15
            indice = 69
    End Select

    imgFecha(0).Tag = indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(indice).Text <> "" Then frmC.NovaData = txtcodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(CByte(imgFecha(0).Tag)) '<===
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
            Case 10: KEYFecha KeyAscii, 1 'fecha
            Case 14: KEYFecha KeyAscii, 0 'fecha
            
            Case 13: KEYFecha KeyAscii, 5 'fecha desde
            Case 15: KEYFecha KeyAscii, 6 'fecha hasta
            
            Case 16: KEYFecha KeyAscii, 2 'fecha desde
            Case 17: KEYFecha KeyAscii, 3 'fecha hasta
            
            Case 36: KEYFecha KeyAscii, 7 'fecha desde
            Case 37: KEYFecha KeyAscii, 8 'fecha hasta
            
            Case 42: KEYFecha KeyAscii, 9 'fecha desde
            Case 43: KEYFecha KeyAscii, 10 'fecha hasta
            
            Case 22: KEYFecha KeyAscii, 4 'fecha
            
            Case 6: KEYBusqueda KeyAscii, 6 ' socio desde
            Case 7: KEYBusqueda KeyAscii, 7 ' socio hasta
            
            Case 34: KEYBusqueda KeyAscii, 2 ' socio desde
            Case 35: KEYBusqueda KeyAscii, 3 ' socio hasta
            
            Case 23: KEYBusqueda KeyAscii, 0 ' socio desde
            Case 24: KEYBusqueda KeyAscii, 1 ' socio hasta
            
            Case 40: KEYBusqueda KeyAscii, 4 ' socio desde
            Case 41: KEYBusqueda KeyAscii, 5 ' socio hasta
            
            Case 56: KEYBusqueda KeyAscii, 9 ' socio factura
            Case 54: KEYFecha KeyAscii, 11 'fecha factura
            
            Case 64: KEYFecha KeyAscii, 12 'fecha desde
            Case 65: KEYFecha KeyAscii, 13 'fecha hasta
            
            Case 67: KEYBusqueda KeyAscii, 10 ' socio desde
            Case 68: KEYBusqueda KeyAscii, 11 ' socio hasta
            
            Case 69: KEYFecha KeyAscii, 15 'fecha
            
            Case 73: KEYFecha KeyAscii, 14 'fecha
            
            Case 74: KEYBusqueda KeyAscii, 12 ' socio desde
            Case 75: KEYBusqueda KeyAscii, 13 ' socio hasta
            
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
Dim cad As String, cadTipo As String 'tipo cliente
Dim Precio As Currency

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1, 18, 19 ' Nro.hidrantes
    
        Case 10, 13, 14, 15, 16, 17, 22, 36, 37, 42, 43, 54, 64, 65, 69, 73 'FECHAS
            If txtcodigo(Index).Text <> "" Then
                If PonerFormatoFecha(txtcodigo(Index)) Then
                End If
            End If
            
        Case 2, 3 ' rangos de consumo
            PonerFormatoEntero txtcodigo(Index)
            
        Case 4, 5 'precios para los rangos de consumo
            PonerFormatoDecimal txtcodigo(Index), 7

        Case 6, 7, 23, 24, 34, 35, 40, 41, 56, 67, 68, 74, 75 'socios
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
        
        Case 8 ' euros/accion
            '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                PonerFormatoDecimal txtcodigo(Index), 7
            Else
                PonerFormatoDecimal txtcodigo(Index), 3
            End If

        Case 70, 71 ' cuota amortizacion y de talla ordinaria
            PonerFormatoDecimal txtcodigo(Index), 3
            Precio = Round2((CCur(ImporteSinFormato(ComprobarCero(txtcodigo(70).Text))) + CCur(ImporteSinFormato(ComprobarCero(txtcodigo(71).Text)))) / 200, 4)
            txtNombre(1).Text = Format(Precio, "##,##0.0000")

        Case 72, 66 ' cuota amortizacion y de talla ordinaria
            PonerFormatoDecimal txtcodigo(Index), 3
            
            Precio = Round2((CCur(ImporteSinFormato(ComprobarCero(txtcodigo(72).Text))) + CCur(ImporteSinFormato(ComprobarCero(txtcodigo(66).Text)))) / 200, 4)
            txtNombre(0).Text = Format(Precio, "##,##0.0000")
        
        Case 53 ' bonificacion
            PonerFormatoDecimal txtcodigo(Index), 4
            If ComprobarCero(txtcodigo(53).Text) = 0 Then
                'el recargo es el siguiente campo
                PonerFoco txtcodigo(61)
            Else
                'el concepto es el siguiente campo
                PonerFoco txtcodigo(9)
            End If

        Case 61 ' recargo
            PonerFormatoDecimal txtcodigo(Index), 4

        Case 21, 26, 28, 30, 32 ' Importes de recibo de contadores
            PonerFormatoDecimal txtcodigo(Index), 3
            CalcularTotales
        
        Case 44 ' numero de etiquetas
            PonerFormatoEntero txtcodigo(Index)

        Case 51 ' lectura
            PonerFormatoEntero txtcodigo(Index)
            
        Case 52 ' nro de factura
            PonerFormatoEntero txtcodigo(Index)
            
    End Select
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

Private Sub FrameTomaLecturaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para la impresion de toma de lectura
    Me.FrameTomaLectura.visible = visible
    If visible = True Then
        Me.FrameTomaLectura.Top = -90
        Me.FrameTomaLectura.Left = 0
        Me.FrameTomaLectura.Height = 3795
        Me.FrameTomaLectura.Width = 6105
        W = Me.FrameTomaLectura.Width
        H = Me.FrameTomaLectura.Height
    End If
End Sub

Private Sub FrameComprobacionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameComprobacion.visible = visible
    If visible = True Then
        Me.FrameComprobacion.Top = -90
        Me.FrameComprobacion.Left = 0
        Me.FrameComprobacion.Height = 3885
        Me.FrameComprobacion.Width = 6945
        W = Me.FrameComprobacion.Width
        H = Me.FrameComprobacion.Height
    End If
End Sub

Private Sub FrameCartaTallasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameCartaTallas.visible = visible
    If visible = True Then
        Me.FrameCartaTallas.Top = -90
        Me.FrameCartaTallas.Left = 0
        Me.FrameCartaTallas.Height = 3885
        Me.FrameCartaTallas.Width = 6945
        W = Me.FrameCartaTallas.Width
        H = Me.FrameCartaTallas.Height
    End If
End Sub

Private Sub FrameReciboTallaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameReciboTalla.visible = visible
    If visible = True Then
        Me.FrameReciboTalla.Top = -90
        Me.FrameReciboTalla.Left = 0
        Me.FrameReciboTalla.Height = 5295
        Me.FrameReciboTalla.Width = 6945
        W = Me.FrameReciboTalla.Width
        H = Me.FrameReciboTalla.Height
        
        If OpcionListado = 11 Then ' generacion de recibos de cuotas
            Me.FrameCuota.visible = True
            Me.FrameCuota.Enabled = True
            Me.FrameBonif.visible = False
            Me.FrameBonif.Enabled = False
        Else
            Label12.Caption = "Cálculo Bonificación Recibos Talla"
            Me.FrameCuota.visible = False
            Me.FrameCuota.Enabled = False
            Me.FrameBonif.visible = True
            Me.FrameBonif.Enabled = True
            
            Me.FrameBonif.Top = 2480
        End If
    End If
End Sub




Private Sub FrameReciboMantenimientoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameReciboMantenimiento.visible = visible
    If visible = True Then
        Me.FrameReciboMantenimiento.Top = -90
        Me.FrameReciboMantenimiento.Left = 0
        Me.FrameReciboMantenimiento.Height = 7005
        Me.FrameReciboMantenimiento.Width = 6945
        W = Me.FrameReciboMantenimiento.Width
        H = Me.FrameReciboMantenimiento.Height
    End If
End Sub

Private Sub FrameReciboConsumoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameReciboConsumo.visible = visible
    If visible = True Then
        Me.FrameReciboConsumo.Top = -90
        Me.FrameReciboConsumo.Left = 0
        Me.FrameReciboConsumo.Height = 6285 '5655
        Me.FrameReciboConsumo.Width = 6945
        W = Me.FrameReciboConsumo.Width
        H = Me.FrameReciboConsumo.Height
    End If
End Sub

Private Sub FrameReciboContadorVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameReciboContador.visible = visible
    If visible = True Then
        Me.FrameReciboContador.Top = -90
        Me.FrameReciboContador.Left = 0
        Me.FrameReciboContador.Height = 7725
        Me.FrameReciboContador.Width = 8235
        W = Me.FrameReciboContador.Width
        H = Me.FrameReciboContador.Height
    End If
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


Private Sub FrameFacturasHidranteVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameFacturasHidrante.visible = visible
    If visible = True Then
        Me.FrameFacturasHidrante.Top = -90
        Me.FrameFacturasHidrante.Left = 0
        Me.FrameFacturasHidrante.Height = 6030 '4230
        Me.FrameFacturasHidrante.Width = 6675
        W = Me.FrameFacturasHidrante.Width
        H = Me.FrameFacturasHidrante.Height
    End If
End Sub


Private Sub FrameEtiquetasContadoresVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameEtiquetasContadores.visible = visible
    If visible = True Then
        Me.FrameEtiquetasContadores.Top = -90
        Me.FrameEtiquetasContadores.Left = 0
        Me.FrameEtiquetasContadores.Height = 3885
        Me.FrameEtiquetasContadores.Width = 6945
        W = Me.FrameEtiquetasContadores.Width
        H = Me.FrameEtiquetasContadores.Height
    End If
End Sub


Private Sub FrameRectificacionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameRectificacion.visible = visible
    If visible = True Then
        Me.FrameRectificacion.Top = -90
        Me.FrameRectificacion.Left = 0
        Me.FrameRectificacion.Height = 4680
        Me.FrameRectificacion.Width = 6675
        W = Me.FrameRectificacion.Width
        H = Me.FrameRectificacion.Height
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
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .ConSubInforme = True ' ConSubInforme
        .Opcion = OpcionListado
        .Show vbModal
    End With
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


Private Sub ProcesoFacturacionConsumo(nTabla As String, cadSelect As String, FecFac As String, Consumo As Long, EsRectificativa As Boolean)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadHasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim NRegs As Long

Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Contadores"
                    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        'comprobamos que los tipos de iva existen en la contabilidad de horto
                
        NRegs = TotalFacturasSocios(nTabla, cadSelect)
        If NRegs <> 0 Then
                Me.Pb1.visible = True
                Me.Pb1.Max = NRegs
                Me.Pb1.Value = 0
                Me.Refresh
                Mens = "Proceso Facturación Consumo: " & vbCrLf & vbCrLf
                If vParamAplic.Cooperativa = 7 Then ' QUATRETONDA
                    b = FacturacionConsumoQUATRETONDA(nTabla, cadSelect, FecFac, Me.Pb1, Mens, Consumo, EsRectificativa)
                Else ' MALLAES
                    b = FacturacionConsumo(nTabla, cadSelect, FecFac, Me.Pb1, Mens)
                End If
                If b Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                   
                    'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                    If Me.Check1(2).Value Then
                        cadFormula = ""
                        cadParam = cadParam & "pFecFac= """ & txtcodigo(14).Text & """|"
                        numParam = numParam + 1
                        cadParam = cadParam & "pTitulo= ""Resumen Facturación de Contadores""|"
                        numParam = numParam + 1
                        
'                        FecFac = CDate(txtcodigo(14).Text)
                        cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                        ConSubInforme = False
                        
                        If vParamAplic.Cooperativa = 7 Then
                            Dim vPorcIva As String
                            
                            '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
                            indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
                            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                            cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'                            cadNombreRPT = "rResumFacturasPOZQua.rpt"
                            
                            Dim vSeccion As CSeccion
                            Set vSeccion = New CSeccion
                            If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
                                If Not vSeccion.AbrirConta Then
                                    Exit Sub
                                End If
                            End If
    
                            vPorcIva = ""
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
                            
                            cadParam = cadParam & "pPorcIva=" & vPorcIva & "|"
                            numParam = numParam + 1
                        
                            vSeccion.CerrarConta
                            Set vSeccion = Nothing
                        
                        
                        End If
                        
                        LlamarImprimir
                    End If
                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
                    If Me.Check1(3).Value Then
                        cadFormula = ""
                        cadSelect = ""
                        'Nº Factura
                        cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(0) & "])"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        'Fecha de Factura
'                        FecFac = CDate(txtcodigo(14).Text)
                        cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        indRPT = 46 'Impresion de recibos de consumo de contadores de pozos
                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        'Nombre fichero .rpt a Imprimir
                        cadTitulo = "Reimpresión de Facturas de Contadores"
                        ConSubInforme = True

                        LlamarImprimir

                        If frmVisReport.EstaImpreso Then
'                            ActualizarRegistrosFac "rrecibpozos", cadSelect
                        End If
                    End If
                    'SALIR DE LA FACTURACION DE RECIBOS DE CONTADORES
                    cmdCancel_Click (1)
                Else
                    
                    MsgBox "Error en el proceso" & vbCrLf & Mens, vbExclamation
                    
                    'SALIR DE LA FACTURACION DE RECIBOS DE CONTADORES
                    cmdCancel_Click (1)
                End If
            Else
                MsgBox "No hay contadores a facturar.", vbExclamation
            End If
    End If
End Sub


Private Function FacturacionConsumo(cTabla As String, cwhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

Dim Neto As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vTipoMov As CTiposMov
Dim NumFactu As Long
Dim ImpoIva As Currency
Dim BaseImpo As Currency
Dim TotalFac As Currency


Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long

Dim NumLin As Long
Dim DiferenciaDias As Long

' Calculo de tramos globales a todos los hidrantes de un socio

Dim RsFacturas As ADODB.Recordset
Dim ConsumoTramo1 As Long
Dim ConsumoTramo2 As Long
Dim vConsumo1 As Long
Dim vConsumo2 As Long

    On Error GoTo eFacturacion

    FacturacionConsumo = False
    
    tipoMov = "RCP"
    
    conn.BeginTrans
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rpozos.codsocio,hidrante,nroorden,codparti,poligono,parcelas,hanegada,lect_ant,lect_act,fech_ant,fech_act"
    Sql = Sql & " FROM  " & cTabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by rpozos.codsocio, rpozos.hidrante "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    b = True
    
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        AntSocio = CStr(DBLet(RS!CodSocio, "N"))
        ActSocio = CStr(DBLet(RS!CodSocio, "N"))

        NumFactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            NumFactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                NumFactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe

        NumLin = 0
    End If
    
    
    While Not RS.EOF And b
        HayReg = True
        
        ActSocio = RS!CodSocio
        
        If ActSocio <> AntSocio Then
        
            Sql2 = "select sum(acciones) acciones from rsocios_pozos where codsocio = " & DBSet(AntSocio, "N") 'antes act
            Acciones = DevuelveValor(Sql2)
                                                                            
            Sql2 = "select sum(lect_act - lect_ant) consumo, round(sum(datediff(fech_act, fech_ant)) / count(*),0) dias"
            Sql2 = Sql2 & " from " & cTabla
            If cwhere <> "" Then
                Sql2 = Sql2 & " WHERE " & cwhere & " and "
            Else
                Sql2 = Sql2 & " WHERE "
            End If
            Sql2 = Sql2 & " rpozos.codsocio = " & DBSet(AntSocio, "N") ' antes act
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            If DBLet(Rs2!Consumo, "N") <> 0 And Acciones <> 0 Then
                ConsumoHan = Round2(((DBLet(Rs2!Consumo, "N") / Acciones) * 30) / DBLet(Rs2!Dias, "N"), 0)
            End If
        
            Consumo1 = 0
            Consumo2 = 0
        
            If ConsumoHan < CLng(txtcodigo(3).Text) Then
                If ConsumoHan < CLng(txtcodigo(2).Text) Then
                    Consumo1 = DBLet(Rs2!Consumo, "N")
                    Consumo2 = 0
                Else
                    Consumo1 = CLng(txtcodigo(2).Text) * (Acciones / 30 * DBLet(Rs2!Dias, "N"))
                    Consumo2 = DBLet(Rs2!Consumo) - Consumo1
                End If
            End If
            
            Set Rs2 = Nothing
            
            '[Monica]28/10/2011: añadido el recalculo de tramos de los contadores de la factura
            Sql = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
            
            Set RsFacturas = New ADODB.Recordset
            RsFacturas.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
            ConsumoTramo1 = Consumo1
            ConsumoTramo2 = Consumo2
            
            While Not RsFacturas.EOF
                If DBLet(RsFacturas!Consumo, "N") < ConsumoTramo1 Then
                    vConsumo1 = DBLet(RsFacturas!Consumo, "N")
                    vConsumo2 = 0
                    ConsumoTramo1 = ConsumoTramo1 - DBLet(RsFacturas!Consumo, "N")
                Else
                    vConsumo2 = DBLet(RsFacturas!Consumo, "N") - ConsumoTramo1
                    vConsumo1 = DBLet(RsFacturas!Consumo, "N") - vConsumo2
                    ConsumoTramo2 = ConsumoTramo2 - vConsumo2
                    If ConsumoTramo1 > 0 Then ConsumoTramo1 = ConsumoTramo1 - vConsumo1
                End If
            
                TotalFac = Round2(vConsumo1 * CCur(ImporteSinFormato(txtcodigo(4).Text)), 2) + _
                           Round2(vConsumo2 * CCur(ImporteSinFormato(txtcodigo(5).Text)), 2) + _
                           vParamAplic.CuotaPOZ
            
                Sql = "update rrecibpozos set consumo1 = " & DBSet(vConsumo1, "N") & ", consumo2 = " & DBSet(vConsumo2, "N")
                Sql = Sql & ", baseimpo = " & DBSet(TotalFac, "N") & ", totalfact = " & DBSet(TotalFac, "N")
                Sql = Sql & " where codtipom = 'RCP' and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
                Sql = Sql & " and numlinea = " & DBSet(RsFacturas!numlinea, "N")
                
                conn.Execute Sql
            
                RsFacturas.MoveNext
            Wend
            
            Set RsFacturas = Nothing
            
            
            Sql = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
            
            Set RsFacturas = New ADODB.Recordset
            RsFacturas.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            While Not RsFacturas.EOF
                Sql = "update rpozos set "
                Sql = Sql & " lect_ant = lect_act "
                Sql = Sql & ", fech_ant = fech_act "
                Sql = Sql & ", consumo = 0 "
                Sql = Sql & " WHERE hidrante = " & DBSet(RsFacturas!Hidrante, "T")
                
                conn.Execute Sql
                
                RsFacturas.MoveNext
            Wend
        
            Set RsFacturas = Nothing
                
            'fin añadido
            
            AntSocio = ActSocio
            
            If b Then b = InsertResumen(tipoMov, CStr(NumFactu))
           
            If b Then b = vTipoMov.IncrementarContador(tipoMov)
            
            NumFactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                NumFactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    NumFactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            NumLin = 0
        End If
            
        ConsumoHidrante = DBLet(RS!lect_act, "N") - DBLet(RS!lect_ant, "N")
        Consumo = ConsumoHidrante
        ConsTra1 = 0
        ConsTra2 = 0
            
        If Consumo1 >= Consumo Then
            ConsTra1 = Consumo
            Consumo1 = Consumo1 - ConsTra1
        Else
            ConsTra1 = Consumo1
            Consumo = Consumo - ConsTra1
            If Consumo2 >= Consumo Then
                ConsTra2 = Consumo
                Consumo2 = Consumo2 - ConsTra2
            End If
        End If
        
        TotalFac = Round2(ConsTra1 * CCur(ImporteSinFormato(txtcodigo(4).Text)), 2) + _
                   Round2(ConsTra2 * CCur(ImporteSinFormato(txtcodigo(5).Text)), 2) + _
                   vParamAplic.CuotaPOZ
    
        IncrementarProgresNew Pb1, 1
        
        NumLin = NumLin + 1
        
        DiferenciaDias = DBLet(RS!fech_act, "F") - DBLet(RS!fech_ant, "F")
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, numlinea, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, difdias) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(NumLin, "N") & "," & DBSet(ActSocio, "N") & ","
        Sql = Sql & DBSet(RS!Hidrante, "T") & "," & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(ConsumoHidrante, "N") & "," & DBSet(vParamAplic.CuotaPOZ, "N") & ","
        Sql = Sql & DBSet(RS!lect_ant, "N") & "," & DBSet(RS!fech_ant, "F") & ","
        Sql = Sql & DBSet(RS!lect_act, "N") & "," & DBSet(RS!fech_act, "F") & ","
        Sql = Sql & DBSet(ConsTra1, "N") & "," & DBSet(ImporteSinFormato(txtcodigo(4).Text), "N") & ","
        Sql = Sql & DBSet(ConsTra2, "N") & "," & DBSet(ImporteSinFormato(txtcodigo(5).Text), "N") & ","
        Sql = Sql & "'Recibo de Consumo',0,"
        Sql = Sql & DBSet(DiferenciaDias, "N") & ")"
        
        conn.Execute Sql
        
        '
        '[Monica]21/10/2011: insertamos las distintas fases(acciones) del socio en la facturacion
        '
        Sql = "insert into rrecibpozos_acc(codtipom,numfactu,fecfactu,numlinea,numfases,acciones,observac) "
        Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & ","
        Sql = Sql & DBSet(NumLin, "N") & ", numfases, acciones, observac from rsocios_pozos where codsocio = " & DBSet(ActSocio, "N")
        
        conn.Execute Sql
            
            
        ' actualizar en los acumulados de hidrantes
        Sql = "update rpozos set acumconsumo = acumconsumo + " & DBSet(ConsumoHidrante, "N")
        Sql = Sql & ", acumcuota = acumcuota + " & DBSet(vParamAplic.CuotaPOZ, "N")
        
'        Sql = Sql & ", lect_ant = lect_act "
'        Sql = Sql & ", fech_ant = fech_act "
'        Sql = Sql & ", consumo = 0 "
        
        
        Sql = Sql & " WHERE hidrante = " & DBSet(RS!Hidrante, "T")
        
        conn.Execute Sql
            
            
'        If b Then b = InsertResumen(tipoMov, CStr(numfactu))
        
'        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
'        AntSocio = ActSocio
        
        RS.MoveNext
    Wend
    
    If HayReg Then
        Sql2 = "select sum(acciones) acciones from rsocios_pozos where codsocio = " & DBSet(AntSocio, "N") 'antes act
        Acciones = DevuelveValor(Sql2)
                                                                        
        Sql2 = "select sum(lect_act - lect_ant) consumo, round(sum(datediff(fech_act, fech_ant)) / count(*),0) dias"
        Sql2 = Sql2 & " from " & cTabla
        If cwhere <> "" Then
            Sql2 = Sql2 & " WHERE " & cwhere & " and "
        Else
            Sql2 = Sql2 & " WHERE "
        End If
        Sql2 = Sql2 & " rpozos.codsocio = " & DBSet(AntSocio, "N") ' antes act
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        If DBLet(Rs2!Consumo, "N") <> 0 And Acciones <> 0 Then
            ConsumoHan = Round2(((DBLet(Rs2!Consumo, "N") / Acciones) * 30) / DBLet(Rs2!Dias, "N"), 0)
        End If
    
        Consumo1 = 0
        Consumo2 = 0
    
        If ConsumoHan < CLng(txtcodigo(3).Text) Then
            If ConsumoHan < CLng(txtcodigo(2).Text) Then
                Consumo1 = DBLet(Rs2!Consumo, "N")
                Consumo2 = 0
            Else
                Consumo1 = CLng(txtcodigo(2).Text) * (Acciones / 30 * DBLet(Rs2!Dias, "N"))
                Consumo2 = DBLet(Rs2!Consumo) - Consumo1
            End If
        End If
        
        Set Rs2 = Nothing
    
    
        '[Monica]28/10/2011: añadido el recalculo de tramos de los contadores de la factura
        Sql = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
        
        Set RsFacturas = New ADODB.Recordset
        RsFacturas.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
        ConsumoTramo1 = Consumo1
        ConsumoTramo2 = Consumo2
        
        While Not RsFacturas.EOF
            If DBLet(RsFacturas!Consumo, "N") < ConsumoTramo1 Then
                vConsumo1 = DBLet(RsFacturas!Consumo, "N")
                vConsumo2 = 0
                ConsumoTramo1 = ConsumoTramo1 - DBLet(RsFacturas!Consumo, "N")
            Else
                vConsumo2 = DBLet(RsFacturas!Consumo, "N") - ConsumoTramo1
                vConsumo1 = DBLet(RsFacturas!Consumo, "N") - vConsumo2
                ConsumoTramo2 = ConsumoTramo2 - vConsumo2
                If ConsumoTramo1 > 0 Then ConsumoTramo1 = ConsumoTramo1 - vConsumo1
            End If
        
            TotalFac = Round2(vConsumo1 * CCur(ImporteSinFormato(txtcodigo(4).Text)), 2) + _
                       Round2(vConsumo2 * CCur(ImporteSinFormato(txtcodigo(5).Text)), 2) + _
                       vParamAplic.CuotaPOZ
        
            Sql = "update rrecibpozos set consumo1 = " & DBSet(vConsumo1, "N") & ", consumo2 = " & DBSet(vConsumo2, "N")
            Sql = Sql & ", baseimpo = " & DBSet(TotalFac, "N") & ", totalfact = " & DBSet(TotalFac, "N")
            Sql = Sql & " where codtipom = 'RCP' and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
            Sql = Sql & " and numlinea = " & DBSet(RsFacturas!numlinea, "N")
            
            conn.Execute Sql
        
            RsFacturas.MoveNext
        Wend
        
        Set RsFacturas = Nothing
        
        
        Sql = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
        
        Set RsFacturas = New ADODB.Recordset
        RsFacturas.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        While Not RsFacturas.EOF
            Sql = "update rpozos set "
            Sql = Sql & " lect_ant = lect_act "
            Sql = Sql & ", fech_ant = fech_act "
            Sql = Sql & ", consumo = 0 "
            Sql = Sql & " WHERE hidrante = " & DBSet(RsFacturas!Hidrante, "T")
            
            conn.Execute Sql
            
            RsFacturas.MoveNext
        Wend
    
        Set RsFacturas = Nothing
            
        
        'fin añadido

        b = InsertResumen(tipoMov, CStr(NumFactu))
        If b And HayReg Then b = vTipoMov.IncrementarContador(tipoMov)
    
    End If
    
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionConsumo = False
    Else
        conn.CommitTrans
        FacturacionConsumo = True
    End If
End Function

Private Function FacturacionConsumoQUATRETONDA(cTabla As String, cwhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String, ConsumoRectif As Long, EsRectificativa As Boolean) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

Dim Neto As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vTipoMov As CTiposMov
Dim NumFactu As Long
Dim ImpoIva As Currency
Dim BaseImpo As Currency
Dim TotalFac As Currency

Dim Precio1 As Currency
Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim CuotaHda As Currency

Dim Consumo As Long
Dim ConsumoHidrante As Long

Dim ImpCuota As Currency
Dim ImpConsumoHda As Currency
Dim ImpConsumo As Currency

Dim NumLin As Long
Dim LecturaAct As Long

    On Error GoTo eFacturacion

    FacturacionConsumoQUATRETONDA = False
    
    tipoMov = "RCP"
    
    conn.BeginTrans
    
    If EsRectificativa Then
        Sql = "update rpozos set consumo = " & DBSet(ConsumoRectif, "N")
        Sql = Sql & ", lect_act = " & DBSet(txtcodigo(51).Text, "N")
        Sql = Sql & ", fech_act = " & DBSet(txtcodigo(54).Text, "F")
        Sql = Sql & " where " & cwhere

        conn.Execute Sql
    End If
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rpozos.codsocio,hidrante,nroorden,codparti,poligono,parcelas,hanegada,lect_ant,lect_act,fech_ant,fech_act,nroacciones,codpozo,consumo,calibre "
    Sql = Sql & " FROM  " & cTabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by rpozos.codsocio, rpozos.hidrante "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    b = True
    
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    If vPorcIva = "" Then
        MsgBox "No se ha encontrado el tipo de Iva " & vParamAplic.CodIvaPOZ & ". Revise.", vbExclamation
        conn.RollbackTrans
        Exit Function
    End If
    
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        AntSocio = CStr(DBLet(RS!CodSocio, "N"))
        ActSocio = CStr(DBLet(RS!CodSocio, "N"))

        BaseImpo = 0
        ImpoIva = 0
        TotalFac = 0

        NumFactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            NumFactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                NumFactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe

        NumLin = 0
    End If
    
    
    While Not RS.EOF And b
        HayReg = True
        
        ActSocio = RS!CodSocio
        
        If ActSocio <> AntSocio Then
            
            AntSocio = ActSocio
            
            If b Then b = InsertResumen(tipoMov, CStr(NumFactu))
           
            If b Then b = vTipoMov.IncrementarContador(tipoMov)
            
            NumFactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                NumFactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    NumFactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            NumLin = 0
        End If
            
        Sql2 = "select precio1, imporcuota, imporcuotahda from rtipopozos where codpozo = " & DBSet(RS!Codpozo, "N")
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs2.EOF Then
            Precio1 = DBLet(Rs2.Fields(0).Value, "N")
            ImpCuota = DBLet(Rs2.Fields(1).Value, "N")
            CuotaHda = DBLet(Rs2.Fields(2).Value, "N")
        End If
            
        Set Rs2 = Nothing
            
        Acciones = DBLet(RS!nroacciones, "N")
            
        ImpConsumo = Round2(DBLet(RS!Consumo, "N") * Precio1, 2)
        ImpConsumoHda = Round2(Acciones * CuotaHda, 2)
            
        '[Monica]22/09/2011: en caso de venir de una rectificativa solo se cobra el consumo
        If EsRectificativa Then
            ImpConsumoHda = 0
            ImpCuota = 0
            Acciones = 0
        End If
            
        BaseImpo = ImpConsumo + ImpCuota + ImpConsumoHda
        ImpoIva = Round2(BaseImpo * vPorcIva / 100, 2)
        TotalFac = BaseImpo + ImpoIva
    
        IncrementarProgresNew Pb1, 1
        
        NumLin = NumLin + 1
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, numlinea, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, "
        '[Monica]28/02/2012: introducimos los nuevos campos
        Sql = Sql & "codparti, calibre, codpozo) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(NumLin, "N") & "," & DBSet(ActSocio, "N") & ","
        Sql = Sql & DBSet(RS!Hidrante, "T") & "," & DBSet(BaseImpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(vPorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(RS!Consumo, "N") & "," & DBSet(ImpCuota, "N") & ","
        Sql = Sql & DBSet(RS!lect_ant, "N") & "," & DBSet(RS!fech_ant, "F") & ","
        Sql = Sql & DBSet(RS!lect_act, "N") & "," & DBSet(RS!fech_act, "F") & ","
        Sql = Sql & DBSet(RS!Consumo, "N") & "," & DBSet(Precio1, "N") & "," ' consumo
        Sql = Sql & DBSet(Acciones, "N") & "," & DBSet(CuotaHda, "N") & ","  ' mantenimiento
        Sql = Sql & DBSet(txtcodigo(48).Text, "T") & ",0,"
        '[Monica]28/02/2012: introducimos los nuevos campos: partida,calibre y codpozo
        Sql = Sql & DBSet(RS!codparti, "N") & "," & DBSet(RS!calibre, "N") & "," & DBSet(RS!Codpozo, "N") & ")"
        
        conn.Execute Sql
            
        ' actualizar en los acumulados de hidrantes
        Sql = "update rpozos set acumconsumo = acumconsumo + " & DBSet(RS!Consumo, "N")
        Sql = Sql & ", acumcuota = acumcuota + " & DBSet(ImpCuota, "N")
        
        Sql = Sql & ", lect_ant = lect_act "
        Sql = Sql & ", fech_ant = fech_act "
        Sql = Sql & ", lect_act = null "
        Sql = Sql & ", fech_act = null "
        Sql = Sql & ", consumo = 0 "
        
        Sql = Sql & " WHERE hidrante = " & DBSet(RS!Hidrante, "T")
        
        conn.Execute Sql
        
        RS.MoveNext
    Wend
    
    If HayReg Then b = InsertResumen(tipoMov, CStr(NumFactu))
    If b And HayReg Then b = vTipoMov.IncrementarContador(tipoMov)
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionConsumoQUATRETONDA = False
    Else
        conn.CommitTrans
        FacturacionConsumoQUATRETONDA = True
    End If
End Function



Private Function TotalFacturasSocios(cTabla As String, cwhere As String) As Long
Dim Sql As String

    TotalFacturasSocios = 0
    
    Sql = "SELECT  count(distinct rpozos.codsocio) "
    Sql = Sql & " FROM  " & cTabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If

    TotalFacturasSocios = TotalRegistros(Sql)

End Function

Private Function TotalFacturasHidrante(cTabla As String, cwhere As String) As Long
Dim Sql As String

    TotalFacturasHidrante = 0
    
    Sql = "SELECT  count(distinct rpozos.hidrante) "
    Sql = Sql & " FROM  " & cTabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If

    TotalFacturasHidrante = TotalRegistros(Sql)

End Function



Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim Sql As String
Dim FecFac As Date
Dim FecUlt As Date

    On Error GoTo EDatosOK

    DatosOk = False
    b = True
    Select Case OpcionListado
        Case 3 ' generacion de recibos de consumo
            If txtcodigo(14).Text = "" Then
                MsgBox "Debe introducir un valor para la Fecha de Recibo.", vbExclamation
                PonerFoco txtcodigo(14)
                b = False
            End If
            If b Then
                If txtcodigo(2).Text = "" Or txtcodigo(3).Text = "" Or txtcodigo(4).Text = "" Or txtcodigo(5).Text = "" Then
                    MsgBox "Debe introducir valores en rangos y precios de los tramos.", vbExclamation
                    PonerFoco txtcodigo(2)
                    b = False
                End If
            End If
                    
        Case 4 ' generacion de recibos de mantenimiento
            If txtcodigo(10).Text = "" Then
                MsgBox "Debe introducir un valor para la Fecha de Recibo.", vbExclamation
                PonerFoco txtcodigo(10)
                b = False
            End If
            If b Then
                If txtcodigo(8).Text = "" Then
                    MsgBox "Debe introducir un valor en Euros/Acción.", vbExclamation
                    PonerFoco txtcodigo(8)
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(9).Text = "" Then
                    MsgBox "Debe introducir un valor en el concepto", vbExclamation
                    PonerFoco txtcodigo(9)
                    b = False
                End If
            End If
            
            'o metemos una bonificacion o un recargo o nada, pero no ambos a la vez
            If b Then
                If ComprobarCero(txtcodigo(53).Text) <> 0 And ComprobarCero(txtcodigo(61).Text) <> 0 Then
                    MsgBox "No se permite introducir a la vez una Bonificacion y un Recargo. Revise.", vbExclamation
                    PonerFoco txtcodigo(53)
                    b = False
                End If
            End If
            
        Case 5 ' generacion de recibos de contadores
            If txtcodigo(22).Text = "" Then
                MsgBox "Debe introducir un valor para la Fecha de Recibo.", vbExclamation
                PonerFoco txtcodigo(22)
                b = False
            End If
            If b Then
                If txtcodigo(21).Text <> "" And txtcodigo(20).Text = "" Then
                    MsgBox "Si introduce un Importe para Mano de Obra, debe introducir un Concepto.", vbExclamation
                    PonerFoco txtcodigo(20)
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(26).Text <> "" And txtcodigo(25).Text = "" Then
                    MsgBox "Si introduce un Importe para el Artículo 1, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtcodigo(25)
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(28).Text <> "" And txtcodigo(27).Text = "" Then
                    MsgBox "Si introduce un Importe para el Artículo 2, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtcodigo(27)
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(30).Text <> "" And txtcodigo(29).Text = "" Then
                    MsgBox "Si introduce un Importe para el Artículo 3, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtcodigo(29)
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(32).Text <> "" And txtcodigo(31).Text = "" Then
                    MsgBox "Si introduce un Importe para el Artículo 4, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtcodigo(31)
                    b = False
                End If
            End If
            
            If b Then
                If txtcodigo(33).Text = "" Then
                    MsgBox "El Recibo debe de ser de un valor distinto de cero. Revise."
                    PonerFoco txtcodigo(20)
                    b = False
                End If
            End If
    
        Case 8 ' etiquetas de contadores
            If txtcodigo(44).Text = 0 Then
                MsgBox "El número de etiquetas debe ser superior a 0. Revise."
                PonerFoco txtcodigo(44)
                b = False
            End If
        
            If b Then
                If Trim(txtcodigo(45).Text) = "" And Trim(txtcodigo(46).Text) = "" And Trim(txtcodigo(47).Text) = "" Then
                    MsgBox "Debe haber algún valor en alguna de las Líneas. Revise."
                    PonerFoco txtcodigo(45)
                    b = False
                End If
            End If
            
        Case 9 ' Rectificacion de Lecturas
            If txtcodigo(52).Text = "" Then
                MsgBox "Debe introducir un Nº de Factura. Revise", vbExclamation
                PonerFoco txtcodigo(52)
                b = False
            End If
            If b And txtcodigo(56).Text = "" Then
                MsgBox "Debe introducir el Socio de la Factura. Revise", vbExclamation
                PonerFoco txtcodigo(56)
                b = False
            End If
            If b And txtcodigo(55).Text = "" Then
                MsgBox "Debe introducir el Hidrante de la Factura. Revise", vbExclamation
                PonerFoco txtcodigo(55)
                b = False
            End If
            If b And txtcodigo(54).Text = "" Then
                MsgBox "Debe introducir la Fecha de la Factura. Revise", vbExclamation
                PonerFoco txtcodigo(54)
                b = False
            End If
            If b And txtcodigo(51).Text = "" Then
                MsgBox "Debe introducir cual es la lectura actual. Revise", vbExclamation
                PonerFoco txtcodigo(51)
                b = False
            End If
            If b Then
                Sql = "select count(*) from rrecibpozos where codtipom = " & DBSet(Mid(Combo1(2).Text, 1, 3), "T")
                Sql = Sql & " and numfactu = " & DBSet(txtcodigo(52).Text, "N")
                Sql = Sql & " and codsocio = " & DBSet(txtcodigo(56).Text, "N")
                Sql = Sql & " and hidrante = " & DBSet(txtcodigo(55).Text, "T")
                If TotalRegistros(Sql) = 0 Then
                    MsgBox "No existe ninguna factura con estos datos para rectificar. Revise.", vbExclamation
                    PonerFoco txtcodigo(52)
                    b = False
                Else
                    ' miramos si es la ultima factura de ese hidrante
                    ' en este caso no debemos hacer la rectificativa porque dejariamos el hidrante con las
                    ' lecturas incorrectas
                    Sql = "select max(fecfactu) from rrecibpozos where codtipom = " & DBSet(Mid(Combo1(2).Text, 1, 3), "T")
                    Sql = Sql & " and hidrante = " & DBSet(txtcodigo(55).Text, "T")
                    FecUlt = DevuelveValor(Sql)
                    
                    Sql = "select fecfactu from rrecibpozos where codtipom = " & DBSet(Mid(Combo1(2).Text, 1, 3), "T")
                    Sql = Sql & " and numfactu= " & DBSet(txtcodigo(52).Text, "N")
                    Sql = Sql & " and hidrante = " & DBSet(txtcodigo(55).Text, "T")
                    FecFac = DevuelveValor(Sql)
                    
                    If FecUlt > FecFac Then
                        MsgBox "Existe un factura de fecha posterior sobre este hidrante, no se permite el proceso. Revise.", vbExclamation
                        PonerFoco txtcodigo(52)
                        b = False
                    End If
                    
                    If b Then
                        If CDate(txtcodigo(54).Text) < FecUlt Then
                            MsgBox "la fecha de la factura rectificativa es inferior a la que rectifica. Revise.", vbExclamation
                            PonerFoco txtcodigo(52)
                            b = False
                        End If
                    End If
                    
                    ' comprobaciones contables
                End If
            End If
            
            
        Case 10 ' Carta de Tallas a socios
            If txtcodigo(69).Text = "" Then
                MsgBox "Debe introducir la fecha de recibo. Revise.", vbExclamation
                PonerFoco txtcodigo(69)
                b = False
            End If
            
        Case 11, 12 ' generacion y a actualizacion de recibos de talla para Escalona
            If txtcodigo(73).Text = "" Then
                MsgBox "Debe introducir la fecha de recibo. Revise.", vbExclamation
                PonerFoco txtcodigo(73)
                b = False
            End If
            
            If OpcionListado = 11 Then
                If CCur(ComprobarCero(txtNombre(0).Text)) = 0 Then
                    MsgBox "Debe introducir un valor en cuotas para facturar. Revise.", vbExclamation
                    PonerFoco txtcodigo(72)
                    b = False
                End If
            End If
            If OpcionListado = 12 Then
                'o metemos una bonificacion o un recargo o nada, pero no ambos a la vez
                If b Then
                    If ComprobarCero(txtcodigo(78).Text) <> 0 And ComprobarCero(txtcodigo(77).Text) <> 0 Then
                        MsgBox "No se permite introducir a la vez una Bonificacion y un Recargo. Revise.", vbExclamation
                        PonerFoco txtcodigo(78)
                        b = False
                    End If
                End If
            End If
    End Select
    DatosOk = b
    
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


'????????????????????????????????????????????????????
'???????????
'??????????? FACTURACION MANTENIMIENTO ???????????????
'???????????
'?????????????????????????????????????????????????????

Private Sub ProcesoFacturacionMantenimiento(nTabla As String, cadSelect As String)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadHasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim NRegs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String
    
    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Mantenimiento"
                    
    NRegs = TotalRegFacturasMto(nTabla, cadSelect)
    If NRegs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.Pb2.visible = True
        Me.Pb2.Max = NRegs
        Me.Pb2.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturación Mantenimiento: " & vbCrLf & vbCrLf
        b = FacturacionMantenimiento(nTabla, cadSelect, txtcodigo(10).Text, Me.Pb2, Mens)
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadFormula = ""
                cadParam = cadParam & "pFecFac= """ & txtcodigo(10).Text & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pTitulo= ""Resumen Facturación Mantenimiento""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtcodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                ConSubInforme = False
                
                LlamarImprimir
            End If
            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
            If Me.Check1(0).Value Then
                cadFormula = ""
                cadSelect = ""
                'Nº Factura
                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(1) & "])"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                'Fecha de Factura
                FecFac = CDate(txtcodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresión de Facturas de Mantenimiento"
                ConSubInforme = True

                LlamarImprimir

                If frmVisReport.EstaImpreso Then
'                            ActualizarRegistrosFac "rrecibpozos", cadSelect
                End If
            End If
            'SALIR DE LA FACTURACION DE RECIBOS DE CONTADORES
            cmdCancel_Click (1)
        Else
            MsgBox "Error en el proceso" & vbCrLf & Mens, vbExclamation
        End If
    End If
End Sub

Private Sub ProcesoFacturacionMantenimientoUTXERA(nTabla As String, cadSelect As String)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadHasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim NRegs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Mantenimiento"
                    
    NRegs = TotalRegFacturasMtoUTXERA(nTabla, cadSelect)
    If NRegs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.Pb2.visible = True
        Me.Pb2.Max = NRegs
        Me.Pb2.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturación Mantenimiento: " & vbCrLf & vbCrLf
        b = FacturacionMantenimientoUTXERA(nTabla, cadSelect, txtcodigo(10).Text, Me.Pb2, Mens)
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadFormula = ""
                cadParam = cadParam & "pFecFac= """ & txtcodigo(10).Text & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pTitulo= ""Resumen Facturación Mantenimiento""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtcodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                ConSubInforme = False
                
                LlamarImprimir
            End If
            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
            If Me.Check1(0).Value Then
                cadFormula = ""
                cadSelect = ""
                'Nº Factura
                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(1) & "])"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                'Fecha de Factura
                FecFac = CDate(txtcodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresión de Facturas de Mantenimiento"
                ConSubInforme = True

                LlamarImprimir

                If frmVisReport.EstaImpreso Then
'                            ActualizarRegistrosFac "rrecibpozos", cadSelect
                End If
            End If
            'SALIR DE LA FACTURACION DE RECIBOS DE CONTADORES
            cmdCancel_Click (1)
        Else
            MsgBox "Error en el proceso" & vbCrLf & Mens, vbExclamation
        End If
    End If
End Sub

Private Sub ProcesoFacturacionMantenimientoESCALONA(nTabla As String, cadSelect As String)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadHasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim NRegs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Mantenimiento"
                    
    NRegs = TotalRegFacturasMtoUTXERA(nTabla, cadSelect)
    If NRegs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.Pb2.visible = True
        Me.Pb2.Max = NRegs
        Me.Pb2.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturación Mantenimiento: " & vbCrLf & vbCrLf
        b = FacturacionMantenimientoESCALONA(nTabla, cadSelect, txtcodigo(10).Text, Me.Pb2, Mens)
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadFormula = ""
                cadParam = cadParam & "pFecFac= """ & txtcodigo(10).Text & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pTitulo= ""Resumen Facturación Mantenimiento""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtcodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                ConSubInforme = False
                
                LlamarImprimir
            End If
            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
            If Me.Check1(0).Value Then
                cadFormula = ""
                cadSelect = ""
                'Nº Factura
                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(1) & "])"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                'Fecha de Factura
                FecFac = CDate(txtcodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresión de Facturas de Mantenimiento"
                ConSubInforme = True

                LlamarImprimir

                If frmVisReport.EstaImpreso Then
'                            ActualizarRegistrosFac "rrecibpozos", cadSelect
                End If
            End If
            'SALIR DE LA FACTURACION DE RECIBOS DE CONTADORES
            cmdCancel_Click (1)
        Else
            MsgBox "Error en el proceso" & vbCrLf & Mens, vbExclamation
        End If
    End If
End Sub



Private Sub ProcesoFacturacionTallaESCALONA(nTabla As String, cadSelect As String)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadHasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim NRegs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String
Dim cadena As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Talla"
                    
    NRegs = TotalRegFacturasTallaUTXERA(nTabla, cadSelect)
    If NRegs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.pb4.visible = True
        Me.pb4.Max = NRegs
        Me.pb4.Value = 0
        Me.Label2(78).visible = True
        Me.Refresh
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        If OpcionListado = 11 Then
            LOG.Insertar 8, vUsu, "Facturacion Talla Recibos Pozos: " & vbCrLf & nTabla & vbCrLf & cadSelect
        Else
            If CCur(ComprobarCero(txtcodigo(78).Text)) <> 0 Then
                cadena = "Bonificacion: " & CCur(ImporteSinFormato(txtcodigo(78).Text)) & "%"
            Else
                cadena = "Recargo: " & CCur(ImporteSinFormato(txtcodigo(77).Text)) & "%"
            End If
        
            LOG.Insertar 8, vUsu, "Actualización Recibos Talla Pozos: " & vbCrLf & cadena & vbCrLf & cadSelect
        End If
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        Mens = "Proceso Facturación Talla: " & vbCrLf & vbCrLf
        If OpcionListado = 11 Then
            b = FacturacionTallaESCALONA(nTabla, cadSelect, txtcodigo(73).Text, Me.pb4, Mens)
        Else
            Me.Label2(78).Caption = "Comprobando recibos ..."
            Me.Refresh
            If Not HayFactContabilizadas(nTabla, cadSelect) Then
                Me.Label2(78).Caption = "Actualizando recibos ..."
                Me.Refresh
                b = ActualizacionTallaESCALONA(nTabla, cadSelect, txtcodigo(73).Text, Me.pb4, Mens)
            Else
                Me.pb4.visible = False
                Me.Label2(78).visible = False
                DoEvents
                Exit Sub
            End If
        End If
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(7).Value Then
                cadFormula = ""
                cadParam = cadParam & "pFecFac= """ & txtcodigo(73).Text & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pTitulo= ""Resumen Facturación Talla""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtcodigo(73).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                ConSubInforme = False
                
                LlamarImprimir
            End If
            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
            If Me.Check1(6).Value Then
                cadFormula = ""
                cadSelect = ""
                'Nº Factura
                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(3) & "])"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                'Fecha de Factura
                FecFac = CDate(txtcodigo(73).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = Replace(nomDocu, "Mto.", "Tal.")
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresión de Facturas de Talla"
                ConSubInforme = True

                LlamarImprimir

                If frmVisReport.EstaImpreso Then
'                            ActualizarRegistrosFac "rrecibpozos", cadSelect
                End If
            End If
            'SALIR DE LA FACTURACION DE RECIBOS DE CONTADORES
            cmdCancel_Click (1)
        Else
            MsgBox "Error en el proceso" & vbCrLf & Mens, vbExclamation
        End If
    End If
End Sub

Private Function HayFactContabilizadas(Tabla As String, cSelect As String) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Contabiliz As Boolean
Dim LEtra As String
Dim EstaEnTesoreria As String
Dim numasien As String

    On Error GoTo eHayFactContabilizadas

    Screen.MousePointer = vbHourglass

    Sql = "SELECT rrecibpozos.* "
    Sql = Sql & " FROM  " & Tabla

    If cSelect <> "" Then
        cSelect = QuitarCaracterACadena(cSelect, "{")
        cSelect = QuitarCaracterACadena(cSelect, "}")
        cSelect = QuitarCaracterACadena(cSelect, "_1")
        Sql = Sql & " WHERE " & cSelect
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Contabiliz = False
    While Not RS.EOF And Not Contabiliz
    
        'Cojo la letra de serie
        LEtra = ObtenerLetraSerie2(DBLet(RS!CodTipom))
        
        'Primero comprobaremos que esta el cobro en contabilidad
        EstaEnTesoreria = ""
        If ComprobarCobroArimoney(EstaEnTesoreria, LEtra, CLng(DBLet(RS!NumFactu)), CDate(DBLet(RS!fecfactu))) Then
            MsgBox "La factura " & LEtra & " " & DBLet(RS!NumFactu) & " de fecha " & DBLet(RS!fecfactu) & vbCrLf & EstaEnTesoreria & vbCrLf & vbCrLf & "Revise.", vbExclamation
            Contabiliz = True
        End If
    
        ' En Escalona no va a estar en registro de iva nunca
        If Not Contabiliz Then
            If LEtra <> "" Then
                numasien = DevuelveDesdeBDNew(cConta, "cabfact", "numasien", "numserie", LEtra, "T", , "codfaccl", CStr(RS!NumFactu), "N", "anofaccl", Year(RS!fecfactu), "N")
                If Val(ComprobarCero(numasien)) <> 0 Then
                    
                Else
                    numasien = ""
                End If
            Else
                numasien = ""
            End If
            If numasien <> "" Then
                LEtra = "La factura esta en la contabilidad, " & DBLet(RS!NumFactu) & " de fecha " & DBLet(RS!fecfactu)
                If numasien <> "" Then LEtra = LEtra & vbCrLf & "Nº asiento: " & numasien
                
                numasien = String(50, "*") & vbCrLf
                numasien = numasien & numasien & vbCrLf & vbCrLf
                LEtra = numasien & LEtra & vbCrLf & vbCrLf & numasien
                MsgBox LEtra, vbInformation
                Contabiliz = True
            End If
        End If
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    HayFactContabilizadas = Contabiliz
    
    Screen.MousePointer = vbDefault
    Exit Function

eHayFactContabilizadas:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Hay Facturas Contabilizadas", Err.Description
End Function

Public Function TotalRegFacturasMto(cTabla As String, cwhere As String) As Long
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rsocios_pozos.codsocio, sum(acciones) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    Sql = Sql & " group by 1 having sum(acciones) <> 0"
    
    TotalRegFacturasMto = TotalRegistrosConsulta(Sql)
    
End Function


Public Function TotalRegFacturasMtoUTXERA(cTabla As String, cwhere As String) As Long
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rpozos.codsocio, sum(hanegada) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    Sql = Sql & " group by 1 having sum(hanegada) <> 0"
    
    TotalRegFacturasMtoUTXERA = TotalRegistrosConsulta(Sql)
    
End Function


Public Function TotalRegFacturasTallaUTXERA(cTabla As String, cwhere As String) As Long
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    '[Monica]19/09/2012: ahora se factura al propietario del campo no al socio
    Sql = "Select rcampos.codpropiet codsocio, sum(round(supcoope * " & DBSet(vParamAplic.Faneca, "N") & ",2)) hanegada  FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    Sql = Sql & " group by 1 having sum(round(supcoope * " & DBSet(vParamAplic.Faneca, "N") & ",2)) <> 0"
    
    TotalRegFacturasTallaUTXERA = TotalRegistrosConsulta(Sql)
    
End Function




Private Function FacturacionMantenimiento(cTabla As String, cwhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

Dim Neto As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vTipoMov As CTiposMov
Dim NumFactu As Long
Dim ImpoIva As Currency
Dim BaseImpo As Currency
Dim TotalFac As Currency


Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long


    On Error GoTo eFacturacion

    FacturacionMantenimiento = False
    
    tipoMov = "RMP"
    
    conn.BeginTrans
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rsocios_pozos.codsocio, sum(acciones) acciones "
    Sql = Sql & " FROM  " & cTabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    
    Sql = Sql & " group by 1 having sum(acciones) <> 0 "
    ' ordenado por socio, variedad, campo, calidad
    Sql = Sql & " order by codsocio "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    b = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF And b
        HayReg = True
        
        NumFactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            NumFactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                NumFactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        
        BaseImpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        Sql2 = "select sum(acciones) acciones from rsocios_pozos where codsocio = " & DBSet(RS!CodSocio, "N")
        Acciones = DevuelveValor(Sql2)
        
        TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtcodigo(8).Text)), 2)
    
        IncrementarProgresNew Pb2, 1
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        Sql = Sql & "concepto, contabilizado) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(RS!CodSocio, "N") & ","
        Sql = Sql & ValorNulo & "," & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(txtcodigo(9).Text, "T") & ",0)"
        
        conn.Execute Sql
            
        If b Then b = InsertResumen(tipoMov, CStr(NumFactu))
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        RS.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionMantenimiento = False
    Else
        conn.CommitTrans
        FacturacionMantenimiento = True
    End If
End Function


Private Function FacturacionMantenimientoUTXERA(cTabla As String, cwhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

Dim Neto As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vTipoMov As CTiposMov
Dim NumFactu As Long
Dim ImpoIva As Currency
Dim BaseImpo As Currency
Dim TotalFac As Currency


Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long
Dim Brazas As Long
Dim CadMen As String

Dim PorcDto As Currency
Dim Descuento As Currency


    On Error GoTo eFacturacion

    FacturacionMantenimientoUTXERA = False
    
    tipoMov = "RMP"
    
    conn.BeginTrans
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rpozos.codsocio, rpozos.hidrante, rpozos.hanegada  "
    Sql = Sql & " FROM  " & cTabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    
    
    Sql = Sql & " group by 1, 2 having hanegada <> 0 "
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by codsocio "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    b = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF And b
        HayReg = True
        
        NumFactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            NumFactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                NumFactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        
        BaseImpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        Acciones = DBLet(RS!hanegada, "N")
        
'        Brazas = (Int(Acciones) * 200) + ((Acciones - Int(Acciones)) * 1000)

'        TotalFac = Round2(Brazas * CCur(ImporteSinFormato(txtcodigo(8).Text)), 2)
        TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtcodigo(8).Text)), 2)
    
        
        '[Monica]14/05/2012: tambien añadimos el poder poner una bonificacion o recargo (como en escalona)
        ' si hay bonificacion la calculamos
        If ComprobarCero(txtcodigo(53).Text) <> "0" Then
            PorcDto = CCur(ImporteSinFormato(txtcodigo(53).Text))
            Descuento = Round2(TotalFac * PorcDto / 100, 2)
            
            TotalFac = TotalFac + Descuento
        End If
    
    
        '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
        BaseImpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - BaseImpo
    
        IncrementarProgresNew Pb2, 1
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        Sql = Sql & "concepto, contabilizado, porcdto, impdto, precio) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(RS!CodSocio, "N") & ","
        Sql = Sql & DBSet(RS!Hidrante, "T") & "," & DBSet(BaseImpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(txtcodigo(9).Text, "T") & ",0,"
        Sql = Sql & DBSet(PorcDto, "N") & ","
        Sql = Sql & DBSet(Descuento, "N") & ","
        Sql = Sql & DBSet(CCur(ImporteSinFormato(txtcodigo(8).Text)), "N") & ")"
        
        conn.Execute Sql
            
        If b Then b = InsertResumen(tipoMov, CStr(NumFactu))
        
        CadMen = ""
        If b Then b = RepartoCoopropietarios(tipoMov, CStr(NumFactu), CStr(FecFac), CadMen, False)
        CadMen = "Reparto Coopropietarios: " & CadMen
        
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        RS.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionMantenimientoUTXERA = False
    Else
        conn.CommitTrans
        FacturacionMantenimientoUTXERA = True
    End If
End Function



Private Function FacturacionMantenimientoESCALONA(cTabla As String, cwhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs8 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

Dim Neto As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vTipoMov As CTiposMov
Dim NumFactu As Long
Dim ImpoIva As Currency
Dim BaseImpo As Currency
Dim TotalFac As Currency


Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long
Dim Brazas As Long
Dim CadMen As String

Dim PorcDto As Currency
Dim Descuento As Currency
Dim CadValues As String


    On Error GoTo eFacturacion

    FacturacionMantenimientoESCALONA = False
    
    tipoMov = "RMP"
    
    conn.BeginTrans
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rpozos.codsocio, sum(rpozos.hanegada) hanegada, count(*) nrohidrante  "
'    Sql = "SELECT rpozos.codsocio, round(sum(rcampos.supcoope) * 12.03, 2) hanegada, count(*) nrohidrante  "
    Sql = Sql & " FROM  " & cTabla ' & ") INNER JOIN rcampos On rpozos.codcampo = rcampos.codcampo "

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    
    
    Sql = Sql & " group by 1 having hanegada <> 0 "
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by codsocio "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    b = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF And b
        HayReg = True
        
        NumFactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            NumFactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                NumFactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        
        BaseImpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        Acciones = DBLet(RS!hanegada, "N")
        
'        Brazas = (Int(Acciones) * 200) + ((Acciones - Int(Acciones)) * 1000)
'        Brazas = Acciones * 200

        TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtcodigo(8).Text)), 2)
        
'        ' si lo que hacemos una factura de un importe no multimplicamos por nada
'        If Check1(6).Value Then TotalFac = Round2(DBLet(Rs!nrohidrante, "N") * CCur(ImporteSinFormato(txtCodigo(8).Text)), 2)
    
        ' si hay bonificacion la calculamos
        If ComprobarCero(txtcodigo(53).Text) <> "0" Then
            PorcDto = CCur(ImporteSinFormato(txtcodigo(53).Text)) * (-1)
            Descuento = Round2(TotalFac * PorcDto / 100, 2)
            
            TotalFac = TotalFac + Descuento
        Else
            If ComprobarCero(txtcodigo(61).Text) <> 0 Then
                PorcDto = CCur(ImporteSinFormato(txtcodigo(61).Text))
                Descuento = Round2(TotalFac * PorcDto / 100, 2)
                
                TotalFac = TotalFac + Descuento
            End If
        End If
    
        '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
        BaseImpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - BaseImpo
    
        IncrementarProgresNew Pb2, 1
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        Sql = Sql & "concepto, contabilizado, porcdto, impdto, precio) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(RS!CodSocio, "N") & ","
        Sql = Sql & ValorNulo & "," & DBSet(BaseImpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(txtcodigo(9).Text, "T") & ",0,"
        Sql = Sql & DBSet(PorcDto, "N") & ","
        Sql = Sql & DBSet(Descuento, "N") & ","
        Sql = Sql & DBSet(CCur(ImporteSinFormato(txtcodigo(8).Text)), "N") & ")"
        
        conn.Execute Sql
            
            
        ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
'        Sql = "SELECT hidrante, round(rcampos.supcoope * 12.03, 2) hanegada "
        Sql = "SELECT hidrante, hanegada "
        Sql = Sql & " FROM  " & cTabla '& ") INNER JOIN rcampos ON rpozos.codcampo = rcampos.codcampo"
'        Sql = Sql & " FROM  (" & cTabla & ") INNER JOIN rcampos ON rpozos.codcampo = rcampos.codcampo"
        If cwhere <> "" Then
            Sql = Sql & " WHERE " & cwhere
            Sql = Sql & " and rpozos.codsocio = " & DBSet(RS!CodSocio, "N")
        Else
            Sql = Sql & " where rpozos.codsocio = " & DBSet(RS!CodSocio, "N")
        End If
            
        Set Rs8 = New ADODB.Recordset
        Rs8.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = "insert into rrecibpozos_hid (codtipom, numfactu, fecfactu, hidrante, hanegada) values  "
        CadValues = ""
        While Not Rs8.EOF
            CadValues = CadValues & "('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & ","
            CadValues = CadValues & DBSet(Rs8!Hidrante, "T") & "," & DBSet(Rs8!hanegada, "N") & "),"
            Rs8.MoveNext
        Wend
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute Sql & CadValues
        End If
        Set Rs8 = Nothing
            
        If b Then b = InsertResumen(tipoMov, CStr(NumFactu))
        
'[Monica]10/05/2012: no hay reparto de coopropietarios pq ese reparto va por hidrante, ya lo veremos
'        CadMen = ""
'        If b Then b = RepartoCoopropietarios(tipoMov, CStr(NumFactu), CStr(FecFac), CadMen, False)
'        CadMen = "Reparto Coopropietarios: " & CadMen
'
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        RS.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionMantenimientoESCALONA = False
    Else
        conn.CommitTrans
        FacturacionMantenimientoESCALONA = True
    End If
End Function


Private Function FacturacionTallaESCALONA(cTabla As String, cwhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs8 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

Dim Neto As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vTipoMov As CTiposMov
Dim NumFactu As Long
Dim ImpoIva As Currency
Dim BaseImpo As Currency
Dim TotalFac As Currency


Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long
Dim Brazas As Long
Dim CadMen As String

Dim PorcDto As Currency
Dim Descuento As Currency
Dim CadValues As String
Dim Precio As Currency

Dim PrecioBrz As Currency


    On Error GoTo eFacturacion

    FacturacionTallaESCALONA = False
    
    tipoMov = "TAL"
    
    conn.BeginTrans
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rcampos.codpropiet codsocio, round(sum(rcampos.supcoope) / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
    Sql = Sql & " FROM  " & cTabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    
    
    Sql = Sql & " group by 1 having hanegada <> 0 "
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by codsocio "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    b = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF And b
        HayReg = True
        
        NumFactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            NumFactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                NumFactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        
        BaseImpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        Acciones = DBLet(RS!hanegada, "N")
        
        Precio = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(66).Text))) + CCur(ImporteSinFormato(ComprobarCero(txtcodigo(72).Text)))
        
        TotalFac = Round2(Acciones * Precio, 2)
        
        PrecioBrz = Round2(Precio / 200, 4)
        
        ' si hay bonificacion la calculamos
        If CCur(ComprobarCero(txtcodigo(78).Text)) <> 0 Then
            PorcDto = CCur(ImporteSinFormato(txtcodigo(78).Text)) * (-1)
            Descuento = Round2(TotalFac * PorcDto / 100, 2)
            
            TotalFac = TotalFac + Descuento
        Else
            If CCur(ComprobarCero(txtcodigo(77).Text)) <> 0 Then
                PorcDto = CCur(ImporteSinFormato(txtcodigo(77).Text))
                Descuento = Round2(TotalFac * PorcDto / 100, 2)
                
                TotalFac = TotalFac + Descuento
            End If
        End If
    
        '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
        BaseImpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - BaseImpo
    
        IncrementarProgresNew Pb1, 1
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        Sql = Sql & "concepto, contabilizado, porcdto, impdto, precio) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(RS!CodSocio, "N") & ","
        Sql = Sql & ValorNulo & "," & DBSet(BaseImpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(txtcodigo(76).Text, "T") & ",0,"
        Sql = Sql & DBSet(PorcDto, "N") & ","
        Sql = Sql & DBSet(Descuento, "N") & ","
        Sql = Sql & DBSet(PrecioBrz, "N") & ")"
        
        conn.Execute Sql
            
            
        ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
        Sql = "SELECT rcampos.codcampo, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
        Sql = Sql & " FROM  " & cTabla
        If cwhere <> "" Then
            Sql = Sql & " WHERE " & cwhere
            Sql = Sql & " and rcampos.codpropiet = " & DBSet(RS!CodSocio, "N")
        Else
            Sql = Sql & " where rcampos.codpropiet = " & DBSet(RS!CodSocio, "N")
        End If
            
        Set Rs8 = New ADODB.Recordset
        Rs8.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = "insert into rrecibpozos_cam (codtipom, numfactu, fecfactu, codcampo, hanegada) values  "
        CadValues = ""
        While Not Rs8.EOF
            CadValues = CadValues & "('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & ","
            CadValues = CadValues & DBSet(Rs8!CodCampo, "N") & "," & DBSet(Rs8!hanegada, "N") & "),"
            Rs8.MoveNext
        Wend
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute Sql & CadValues
        End If
        Set Rs8 = Nothing
            
        If b Then b = InsertResumen(tipoMov, CStr(NumFactu))
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        RS.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionTallaESCALONA = False
    Else
        conn.CommitTrans
        FacturacionTallaESCALONA = True
    End If
End Function


Private Function ActualizacionTallaESCALONA(cTabla As String, cwhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs8 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

Dim Neto As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vTipoMov As CTiposMov
Dim NumFactu As Long
Dim ImpoIva As Currency
Dim BaseImpo As Currency
Dim TotalFac As Currency


Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long
Dim Brazas As Long
Dim CadMen As String

Dim PorcDto As Currency
Dim Descuento As Currency
Dim CadValues As String
Dim Precio As Currency

Dim PrecioBrz As Currency
Dim LetraSerie As String
Dim vCta As String

    On Error GoTo eFacturacion

    ActualizacionTallaESCALONA = False
    
    tipoMov = "TAL"
    
    conn.BeginTrans
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rrecibpozos.* "
    Sql = Sql & " FROM  " & cTabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    
    
    ' ordenado por socio
    Sql = Sql & " order by rrecibpozos.codsocio, rrecibpozos.numfactu "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    b = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF And b
        HayReg = True
        
        BaseImpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        TotalFac = DBLet(RS!TotalFact, "N")
        
        ' si hay bonificacion la calculamos
        If CCur(ComprobarCero(txtcodigo(78).Text)) <> 0 Then
            PorcDto = CCur(ImporteSinFormato(txtcodigo(78).Text)) * (-1)
            Descuento = Round2(TotalFac * PorcDto / 100, 2)
            
            TotalFac = TotalFac + Descuento
        Else
            If CCur(ComprobarCero(txtcodigo(77).Text)) <> 0 Then
                PorcDto = CCur(ImporteSinFormato(txtcodigo(77).Text))
                Descuento = Round2(TotalFac * PorcDto / 100, 2)
                
                TotalFac = TotalFac + Descuento
            End If
        End If
    
        '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
        BaseImpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - BaseImpo
    
        IncrementarProgresNew Pb1, 1
        
        'modificamos la tabla de recibos de pozos
        Sql = "update rrecibpozos set baseimpo = " & DBSet(BaseImpo, "N")
        Sql = Sql & ", tipoiva = " & DBSet(vParamAplic.CodIvaPOZ, "N")
        Sql = Sql & ", porc_iva = " & DBSet(PorcIva, "N")
        Sql = Sql & ", imporiva = " & DBSet(ImpoIva, "N")
        Sql = Sql & ", totalfact = " & DBSet(TotalFac, "N")
        Sql = Sql & ", porcdto = " & DBSet(PorcDto, "N")
        Sql = Sql & ", impdto = " & DBSet(Descuento, "N")
        Sql = Sql & " where codtipom = 'TAL'"
        Sql = Sql & " and numfactu = " & DBSet(RS!NumFactu, "N")
        Sql = Sql & " and fecfactu = " & DBSet(RS!fecfactu, "F")
        Sql = Sql & " and codsocio = " & DBSet(RS!CodSocio, "N")
        
        conn.Execute Sql
            
        ' Si el recibo está contabilizado actualizaremos el arimoney
        LetraSerie = DevuelveValor("select letraser from usuarios.stipom where codtipom = 'TAL'")

        Sql = "update scobro set impvenci = " & DBSet(TotalFac, "N")
        Sql = Sql & " where numserie = " & DBSet(LetraSerie, "T")
        Sql = Sql & " and codfaccl = " & DBSet(RS!NumFactu, "N")
        Sql = Sql & " and fecfaccl = " & DBSet(RS!fecfactu, "F")
        Sql = Sql & " and numorden = 1 "
            
        ConnConta.Execute Sql
        
        '[Monica]19/09/2012: al enlazar por el propietario y campos me salian todos los campos de ese propietario,
        '                    si el nro de factura, tipo ya existe no lo volvemos a insertar en el resumen
        '                    He añadido: and totalregistros...
        If b And TotalRegistros("select count(*) from tmpinformes where codusu = " & vUsu.Codigo & " and nombre1 = 'TAL' and importe1 = " & DBSet(RS!NumFactu, "N")) = 0 Then b = InsertResumen("TAL", CStr(RS!NumFactu))
        
        RS.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        ActualizacionTallaESCALONA = False
    Else
        conn.CommitTrans
        ActualizacionTallaESCALONA = True
    End If
End Function












Public Function FacturasGeneradasPOZOS(Tipo As String) As String
Dim Sql As String
Dim RS1 As ADODB.Recordset
Dim cad As String
    
    On Error GoTo eFacturasGeneradas

    FacturasGeneradasPOZOS = ""

    Sql = "select nombre1, importe1 from tmpinformes where codusu = " & vUsu.Codigo
    Sql = Sql & " and nombre1 = "
    Select Case Tipo
        Case 0 ' recibos de consumo de pozos
            Sql = Sql & "'RCP'"
        Case 1 ' recibos de mantenimiento de pozos
            Sql = Sql & "'RMP'"
        Case 2 ' recibos de contadores de pozos
            Sql = Sql & "'RVP'"
        Case 3
            Sql = Sql & "'TAL'"
    End Select
    
    Set RS1 = New ADODB.Recordset
    RS1.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    cad = ""
    While Not RS1.EOF
        cad = cad & DBLet(RS1.Fields(1).Value, "N") & ","
    
        RS1.MoveNext
    Wend
    Set RS1 = Nothing
    
    'si hay facturas quitamos la ultima coma
    If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
    
    FacturasGeneradasPOZOS = cad
    Exit Function
    
eFacturasGeneradas:
    MuestraError Err.Number, "Cadena de Facturas Generadas Pozos", Err.Description
End Function

'??????????????????????????????????????????????????
'???????????
'??????????? FACTURACION CONTADORES ???????????????
'???????????
'??????????????????????????????????????????????????

Private Sub ProcesoFacturacionContadores(nTabla As String, cadSelect As String)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadHasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim NRegs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Contadores"
                    
    NRegs = TotalRegFacturasMto(nTabla, cadSelect)
    If NRegs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.Pb3.visible = True
        Me.Pb3.Max = NRegs
        Me.Pb3.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturación Contadores: " & vbCrLf & vbCrLf
        b = FacturacionContadores(nTabla, cadSelect, txtcodigo(22).Text, Me.Pb3, Mens)
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION de recibos de contadores
            If Me.Check1(5).Value Then
                cadFormula = ""
                cadParam = cadParam & "pFecFac= """ & txtcodigo(22).Text & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pTitulo= ""Resumen Facturación Contadores""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtcodigo(22).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                ConSubInforme = False
                
                LlamarImprimir
            End If
            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
            If Me.Check1(4).Value Then
                cadFormula = ""
                cadSelect = ""
                'Nº Factura
                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(2) & "])"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                'Fecha de Factura
                FecFac = CDate(txtcodigo(22).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de contadores de pozos
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresión de Facturas de Contadores"
                ConSubInforme = True

                LlamarImprimir

                If frmVisReport.EstaImpreso Then
'                            ActualizarRegistrosFac "rrecibpozos", cadSelect
                End If
            End If
            'SALIR DE LA FACTURACION DE RECIBOS DE CONTADORES
            cmdCancel_Click (1)
        Else
            MsgBox "Error en el proceso" & vbCrLf & Mens, vbExclamation
        End If
    End If
End Sub

Private Function FacturacionContadores(cTabla As String, cwhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

Dim Neto As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vTipoMov As CTiposMov
Dim NumFactu As Long
Dim ImpoIva As Currency
Dim BaseImpo As Currency
Dim TotalFac As Currency


Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long


    On Error GoTo eFacturacion

    FacturacionContadores = False
    
    tipoMov = "RVP"
    
    conn.BeginTrans
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rsocios.codsocio "
    Sql = Sql & " FROM  " & cTabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If
    
    Sql = Sql & " group by 1 "
    ' ordenado por socio, variedad, campo, calidad
    Sql = Sql & " order by codsocio "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    b = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF And b
        HayReg = True
        
        NumFactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            NumFactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                NumFactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        
        BaseImpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        TotalFac = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(33).Text)))
    
        IncrementarProgresNew Pb3, 1
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        Sql = Sql & "concepto, contabilizado, conceptomo, importemo, conceptoar1, importear1, conceptoar2, importear2, conceptoar3, "
        Sql = Sql & "importear3, conceptoar4, importear4) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(RS!CodSocio, "N") & ","
        Sql = Sql & ValorNulo & "," & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & ",0,"
        Sql = Sql & DBSet(txtcodigo(20).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtcodigo(21).Text))), "N", "S") & "," ' mano de obra
        Sql = Sql & DBSet(txtcodigo(25).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtcodigo(26).Text))), "N", "S") & "," ' articulo 1
        Sql = Sql & DBSet(txtcodigo(27).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtcodigo(28).Text))), "N", "S") & "," ' articulo 2
        Sql = Sql & DBSet(txtcodigo(29).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtcodigo(30).Text))), "N", "S") & "," ' articulo 3
        Sql = Sql & DBSet(txtcodigo(31).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtcodigo(32).Text))), "N", "S") & ")" ' articulo 4
        
        conn.Execute Sql
            
        If b Then b = InsertResumen(tipoMov, CStr(NumFactu))
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        RS.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionContadores = False
    Else
        conn.CommitTrans
        FacturacionContadores = True
    End If
End Function


Private Sub CalcularTotales()
Dim total As Currency

    total = 0
    
    If txtcodigo(21).Text <> "" Then total = total + CCur(ImporteSinFormato(txtcodigo(21).Text))
    If txtcodigo(26).Text <> "" Then total = total + CCur(ImporteSinFormato(txtcodigo(26).Text))
    If txtcodigo(28).Text <> "" Then total = total + CCur(ImporteSinFormato(txtcodigo(28).Text))
    If txtcodigo(30).Text <> "" Then total = total + CCur(ImporteSinFormato(txtcodigo(30).Text))
    If txtcodigo(32).Text <> "" Then total = total + CCur(ImporteSinFormato(txtcodigo(32).Text))

    txtcodigo(33).Text = total
    PonerFormatoDecimal txtcodigo(33), 3

End Sub


Private Sub CargaCombo()
Dim RS As ADODB.Recordset
Dim Sql As String
Dim I As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    'tipo de fichero
    Combo1(0).AddItem "RCP-Consumo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "RMP-Mantenimiento"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "RVP-Contadores"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    If vParamAplic.Cooperativa = 10 Then
        Combo1(0).AddItem "TAL-Talla"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    End If
    
    'tipo de fichero
    Combo1(1).AddItem "Todas"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    
    Combo1(1).AddItem "RCP-Consumo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "RMP-Mantenimiento"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    Combo1(1).AddItem "RVP-Contadores"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    If vParamAplic.Cooperativa = 10 Then
        Combo1(1).AddItem "TAL-Talla"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 4
    End If
    
    'tipo de fichero
    Combo1(2).AddItem "RCP-Consumo"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "RMP-Mantenimiento"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "RVP-Contadores"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
    If vParamAplic.Cooperativa = 10 Then
        Combo1(2).AddItem "TAL-Talla"
        Combo1(2).ItemData(Combo1(2).NewIndex) = 3
    End If
    
    
    
End Sub


Private Sub ProcesoFacturacionConsumoUTXERA(nTabla As String, cadSelect As String)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadHasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim NRegs As Long
Dim FecFac As Date

Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Contadores"
                    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        'comprobamos que los tipos de iva existen en la contabilidad de horto
                
        NRegs = TotalFacturasHidrante(nTabla, cadSelect)
        If NRegs <> 0 Then
                Me.Pb1.visible = True
                Me.Pb1.Max = NRegs
                Me.Pb1.Value = 0
                Me.Refresh
                Mens = "Proceso Facturación Consumo: " & vbCrLf & vbCrLf
                b = FacturacionConsumoUTXERA(nTabla, cadSelect, txtcodigo(14).Text, Me.Pb1, Mens)
                If b Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                   
                    'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                    If Me.Check1(2).Value Then
                        cadFormula = ""
                        cadParam = cadParam & "pFecFac= """ & txtcodigo(14).Text & """|"
                        numParam = numParam + 1
                        cadParam = cadParam & "pTitulo= ""Resumen Facturación de Contadores""|"
                        numParam = numParam + 1
                        
                        FecFac = CDate(txtcodigo(14).Text)
                        cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                        ConSubInforme = False
                        
                        LlamarImprimir
                    End If
                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
                    If Me.Check1(3).Value Then
                        cadFormula = ""
                        cadSelect = ""
                        'Nº Factura
                        cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(0) & "])"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        'Fecha de Factura
                        FecFac = CDate(txtcodigo(14).Text)
                        cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        indRPT = 46 'Impresion de recibos de consumo de contadores de pozos
                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        'Nombre fichero .rpt a Imprimir
                        cadTitulo = "Reimpresión de Facturas de Contadores"
                        ConSubInforme = True

                        LlamarImprimir

                        If frmVisReport.EstaImpreso Then
'                            ActualizarRegistrosFac "rrecibpozos", cadSelect
                        End If
                    End If
                    'SALIR DE LA FACTURACION DE RECIBOS DE CONTADORES
                    cmdCancel_Click (1)
                Else
                    MsgBox "Error en el proceso" & vbCrLf & Mens, vbExclamation
                End If
            Else
                MsgBox "No hay contadores a facturar.", vbExclamation
            End If
    End If
End Sub


Private Function FacturacionConsumoUTXERA(cTabla As String, cwhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

Dim Neto As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vTipoMov As CTiposMov
Dim NumFactu As Long
Dim ImpoIva As Currency
Dim BaseImpo As Currency
Dim TotalFac As Currency


Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long

Dim CadMen As String

    On Error GoTo eFacturacion

    FacturacionConsumoUTXERA = False

    tipoMov = "RCP"

    conn.BeginTrans


    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rpozos.codsocio,hidrante,nroorden,codparti,poligono,parcelas,hanegada,lect_ant,lect_act,fech_ant,fech_act,codpozo,consumo "
    Sql = Sql & " FROM  " & cTabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql = Sql & " WHERE " & cwhere
    End If

    ' ordenado por socio, hidrante
    Sql = Sql & " order by rpozos.codsocio, rpozos.hidrante "

    Set vSeccion = New CSeccion

    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If

    b = True

    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    If vPorcIva = "" Then vPorcIva = "0"
    PorcIva = CCur(ImporteSinFormato(vPorcIva))

    Set vTipoMov = New CTiposMov

    HayReg = False

    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Not RS.EOF Then
        AntSocio = CStr(DBLet(RS!CodSocio, "N"))
        ActSocio = CStr(DBLet(RS!CodSocio, "N"))

        BaseImpo = 0
        ImpoIva = 0
        TotalFac = 0

        Sql2 = "select rpozos.codsocio, rpozos.consumo " 'sum(lect_act - lect_ant) consumo "
        Sql2 = Sql2 & " from " & cTabla
        If cwhere <> "" Then
            Sql2 = Sql2 & " WHERE " & cwhere & " and "
        Else
            Sql2 = Sql2 & " WHERE "
        End If
        Sql2 = Sql2 & " rpozos.codsocio = " & DBSet(AntSocio, "N")
        Sql2 = Sql2 & " group by 1 order by 1"

        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

        ConsumoHan = 0
        Consumo1 = 0
        If DBLet(Rs2!Consumo, "N") <> 0 Then
            ConsumoHan = DBLet(Rs2!Consumo, "N")
            Consumo1 = ConsumoHan
        End If

        Set Rs2 = Nothing
    End If


    While Not RS.EOF And b
        HayReg = True

        ActSocio = RS!CodSocio

        If ActSocio <> AntSocio Then

            Sql2 = "select rpozos.codsocio, rpozos.consumo " 'sum(lect_act - lect_ant) consumo "
            Sql2 = Sql2 & " from " & cTabla
            If cwhere <> "" Then
                Sql2 = Sql2 & " WHERE " & cwhere & " and "
            Else
                Sql2 = Sql2 & " WHERE "
            End If
            Sql2 = Sql2 & " rpozos.codsocio = " & DBSet(ActSocio, "N")
            Sql2 = Sql2 & " group by 1 order by 1 "

            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

            Consumo1 = 0
            If DBLet(Rs2!Consumo, "N") <> 0 Then
                ConsumoHan = DBLet(Rs2!Consumo, "N")
                Consumo1 = DBLet(Rs2!Consumo, "N")
            End If

            Set Rs2 = Nothing

            AntSocio = ActSocio

        End If

        '[Monica]24/10/2011: añadida esta condicion para que si no hay consumo se actualicen fechas

        If DBLet(RS!Consumo, "N") <> 0 Then
    
            NumFactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                NumFactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    NumFactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
    
            ConsumoHidrante = DBLet(RS!Consumo, "N") 'DBLet(RS!lect_act, "N") - DBLet(RS!lect_ant, "N")
            Consumo = ConsumoHidrante
    
            ConsTra1 = Consumo
            
            ' consumo de agua y consumo de electricidad
            
            BaseImpo = Round2(ConsTra1 * CCur(ImporteSinFormato(txtcodigo(4).Text)), 2) + _
                       Round2(ConsTra1 * CCur(ImporteSinFormato(txtcodigo(5).Text)), 2)
    
    
            ImpoIva = Round2(BaseImpo * PorcIva / 100, 2)
            TotalFac = BaseImpo + Round2(BaseImpo * PorcIva / 100, 2)
    
            IncrementarProgresNew Pb1, 1
    
            'insertar en la tabla de recibos de pozos
            Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, numlinea, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
            Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, codparti, parcelas) "
            Sql = Sql & " values ('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(ActSocio, "N") & ",1,"
            Sql = Sql & DBSet(RS!Hidrante, "T") & "," & DBSet(BaseImpo, "N") & "," & vParamAplic.CodIvaPOZ & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(ConsumoHidrante, "N") & "," & DBSet(0, "N") & ","
            Sql = Sql & DBSet(RS!lect_ant, "N") & "," & DBSet(RS!fech_ant, "F") & ","
            Sql = Sql & DBSet(RS!lect_act, "N") & "," & DBSet(RS!fech_act, "F") & ","
            Sql = Sql & DBSet(ConsTra1, "N") & "," & DBSet(ImporteSinFormato(txtcodigo(4).Text), "N") & ","
            Sql = Sql & DBSet(ConsTra1, "N") & "," & DBSet(ImporteSinFormato(txtcodigo(5).Text), "N") & ","
            
            '[Monica]22/10/2012: si nos han puesto un concepto guardammos el concepto
            ' antes :     Sql = Sql & "'Recibo de Consumo',0,"
            If txtcodigo(48).Text <> "" Then
                Sql = Sql & DBSet(txtcodigo(48).Text, "T") & ",0,"
            Else
                Sql = Sql & DBSet(vTipoMov.NombreMovimiento, "T") & ",0,"
            End If
            
            '[Monica]22/10/2012: guardamos tambien la partida
            Sql = Sql & DBSet(RS!codparti, "N") & "," & DBSet(RS!parcelas, "T") & ")"
    
            conn.Execute Sql

        
            If b Then b = RepartoCoopropietarios(tipoMov, CStr(NumFactu), CStr(FecFac), CadMen)
            CadMen = "Reparto Coopropietarios: " & CadMen
        
            If b Then b = InsertResumen(tipoMov, CStr(NumFactu))
        
            If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        End If

        If DBLet(RS!fech_act, "F") <> "" Then
            ' actualizar en los acumulados de hidrantes
            Sql = "update rpozos set acumconsumo = acumconsumo + " & DBSet(ConsumoHidrante, "N")
            Sql = Sql & ", lect_ant = lect_act "
            Sql = Sql & ", fech_ant = fech_act "
    '        sql = sql & ", lect_act = null "
            Sql = Sql & ", fech_act = null "
            Sql = Sql & ", consumo = 0 "
            Sql = Sql & " WHERE hidrante = " & DBSet(RS!Hidrante, "T")

            conn.Execute Sql
        End If
        
        AntSocio = ActSocio

        RS.MoveNext
    Wend

    vSeccion.CerrarConta
    Set vSeccion = Nothing

eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionConsumoUTXERA = False
    Else
        conn.CommitTrans
        FacturacionConsumoUTXERA = True
    End If
End Function

Private Function TieneCopropietariosPOZOS(Hidrante As String, Propietario As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from rpozos_cooprop where hidrante = " & DBSet(Hidrante, "T") & " and codsocio <> " & DBSet(Propietario, "N")
    
    TieneCopropietariosPOZOS = TotalRegistros(Sql) > 0

End Function

Private Function RepartoCoopropietarios(tipoMov As String, Factura As String, Fecha As String, cadErr As String, Optional SinIva As Boolean) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Dim Numalbar As Long
Dim vTipoMov As CTiposMov

Dim Albaranes As String

Dim tBaseImpo As Currency
Dim tImporIva As Currency
Dim tTotalFact As Currency

Dim vBaseImpo As Currency
Dim vImporIva As Currency
Dim vTotalFact As Currency

Dim CodTipoMov As String
Dim b As Boolean
Dim devuelve As String
Dim Existe As Boolean
Dim Mens As String
Dim NroPropiedad As String
Dim NumReg As Long
Dim campo As Long
Dim Porcentaje As Single
Dim numFac As Long
Dim vPorcen As String

    On Error GoTo eRepartoCoopropietarios

    RepartoCoopropietarios = False
    
    cadErr = ""
    
    b = True
    
    Sql = "select * from rrecibpozos where codtipom  = " & DBSet(tipoMov, "T")
    Sql = Sql & " and numfactu = " & DBSet(Factura, "N")
    Sql = Sql & " and fecfactu = " & DBSet(Fecha, "F")
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        
        If TieneCopropietariosPOZOS(CStr(RS!Hidrante), CStr(RS!CodSocio)) Then
            CodTipoMov = tipoMov
        
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(CodTipoMov) Then

                tBaseImpo = DBLet(RS!BaseImpo, "N")
                tImporIva = DBLet(RS!ImporIva, "N")
                tTotalFact = DBLet(RS!TotalFact, "N")

                Sql2 = "select * from rpozos_cooprop where hidrante = " & DBSet(RS!Hidrante, "T")
                Sql2 = Sql2 & " and rpozos_cooprop.codsocio <> " & DBSet(RS!CodSocio, "N")
                Sql2 = Sql2 & " order by numlinea "
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs2.EOF And b
                    numFac = vTipoMov.ConseguirContador(CodTipoMov)
                    Do
                        devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "numfactu", CStr(numFac), "N", , "codtipom", tipoMov, "T", "fecfactu", Fecha, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (CodTipoMov)
                            numFac = vTipoMov.ConseguirContador(CodTipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                    
                    vBaseImpo = Round2(DBLet(RS!BaseImpo, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImporIva = Round2(DBLet(RS!ImporIva, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vTotalFact = Round2(DBLet(RS!TotalFact, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    
                    tBaseImpo = tBaseImpo - vBaseImpo
                    tImporIva = tImporIva - vImporIva
                    tTotalFact = tTotalFact - vTotalFact
                    
                    
                    'insertar en la tabla de recibos de pozos
                    Sql4 = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, numlinea, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
                    Sql4 = Sql4 & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado) "
                    Sql4 = Sql4 & " values ('" & tipoMov & "'," & DBSet(numFac, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Rs2!CodSocio, "N") & ",1,"
                    Sql4 = Sql4 & DBSet(RS!Hidrante, "T") & "," & DBSet(vBaseImpo, "N") & ","
                    
                    If SinIva Then
                        Sql4 = Sql4 & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    Else
                        Sql4 = Sql4 & vParamAplic.CodIvaPOZ & "," & DBSet(RS!porc_iva, "N") & "," & DBSet(vImporIva, "N") & ","
                    End If
                    
                    Sql4 = Sql4 & DBSet(vTotalFact, "N") & "," & DBSet(RS!Consumo, "N", "S") & "," & DBSet(0, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!lect_ant, "N") & "," & DBSet(RS!fech_ant, "F") & ","
                    Sql4 = Sql4 & DBSet(RS!lect_act, "N") & "," & DBSet(RS!fech_act, "F") & ","
                    Sql4 = Sql4 & DBSet(RS!Consumo1, "N") & "," & DBSet(RS!Precio1, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!Consumo1, "N") & "," & DBSet(RS!Precio2, "N") & ","
                    If tipoMov = "RCP" Then
                        Sql4 = Sql4 & DBSet(RS!Concepto & " " & Format(DBLet(Rs2!Porcentaje, "N"), "##0.00") & "%", "T") & ",0)"
                    Else
                        Sql4 = Sql4 & DBSet(RS!Concepto, "T") & ",0)"
                    End If
                    
                    conn.Execute Sql4
                    
                    If b Then b = InsertResumen(tipoMov, CStr(numFac))
                    
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                If b Then
                
                    vPorcen = DevuelveValor("select porcentaje from rpozos_cooprop where codsocio = " & DBSet(RS!CodSocio, "N") & " and hidrante = " & DBSet(RS!Hidrante, "T"))
                    vPorcen = Format(vPorcen, "##0.00") & "%"
                    vPorcen = " " & vPorcen
                    
                
                    ' ultimo registro la diferencia ( se updatean las tablas del registro de rrecibpozos origen )
                    Sql4 = "update rrecibpozos set baseimpo = " & DBSet(tBaseImpo, "N") & ","
                    
                    If SinIva Then
                    
                    Else
                        Sql4 = Sql4 & "imporiva = " & DBSet(tImporIva, "N") & ","
                    End If
                    
                    Sql4 = Sql4 & "totalfact = " & DBSet(tTotalFact, "N") & ","
                    
                    If tipoMov = "RCP" Then
                        Sql4 = Sql4 & "concepto = concat(concepto,'" & vPorcen & "')"
                    Else
                        Sql4 = Sql4 & "concepto = concepto "
                    End If
                    Sql4 = Sql4 & " where codtipom = " & DBSet(tipoMov, "T")
                    Sql4 = Sql4 & " and numfactu = " & DBSet(Factura, "N")
                    Sql4 = Sql4 & " and fecfactu = " & DBSet(Fecha, "F")
                    
                    conn.Execute Sql4
                    
                    vTipoMov.IncrementarContador (CodTipoMov)
                ' fin de ultimo registro
                End If
            Else
                b = False
            End If
        
        End If
    
    End If
    
    Set RS = Nothing

eRepartoCoopropietarios:
    If Err.Number <> 0 Or Not b Then
        cadErr = cadErr & Err.Description
    Else
        RepartoCoopropietarios = True
    End If
End Function






' estaba en facturacion de consumo

'    If Not RS.EOF Then
'        AntSocio = CStr(DBLet(RS!Codsocio, "N"))
'        ActSocio = CStr(DBLet(RS!Codsocio, "N"))
'
'        baseimpo = 0
'        ImpoIva = 0
'        TotalFac = 0
'
'        Sql2 = "select sum(acciones) acciones from rsocios_pozos where codsocio = " & DBSet(RS!Codsocio, "N")
'        Acciones = DevuelveValor(Sql2)
'
'        Sql2 = "select codsocio, sum(lect_act - lect_ant) consumo, sum(datediff(fech_act,fech_ant)) dias "
'        Sql2 = Sql2 & " from " & cTabla
'        If cWhere <> "" Then
'            Sql2 = Sql2 & " WHERE " & cWhere & " and "
'        Else
'            Sql2 = Sql2 & " WHERE "
'        End If
'        Sql2 = Sql2 & " codsocio = " & DBSet(AntSocio, "N")
'        Sql2 = Sql2 & " group by 1 order by 1"
'
'        Set Rs2 = New ADODB.Recordset
'        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'
'        ConsumoHan = 0
'        If DBLet(Rs2!Consumo, "N") <> 0 And DBLet(Rs2!dias, "N") <> 0 Then
'            ConsumoHan = Round2(((DBLet(Rs2!Consumo, "N") / Acciones) * 30) / DBLet(Rs2!dias, "N"), 0)
'        End If
'
'        If ConsumoHan < CLng(txtcodigo(3).Text) Then
'            If ConsumoHan < CLng(txtcodigo(2).Text) Then
'                Consumo1 = DBLet(Rs2!Consumo, "N")
'                Consumo2 = 0
'            Else
'                Consumo1 = CLng(txtcodigo(2).Text) * (Acciones / 30 * DBLet(Rs2!dias, "N"))
'                Consumo2 = DBLet(Rs2!Consumo) - Consumo1
'            End If
'        End If
'
'        Set Rs2 = Nothing
'
'    End If



'************************
' ANTIGUA FACTURACION CONSUMO PERO POR HIDRANTE: TENIAMOS UN NRO DE FACTURA POR HIDRANTE
'       ACTUALMENTE LA FACTURACION DE CONSUMO ES POR SOCIO
'************************
'Private Function FacturacionConsumo(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim Rs As ADODB.Recordset
'Dim Rs2 As ADODB.Recordset
'
'Dim AntSocio As String
'Dim ActSocio As String
'
'Dim HayReg As Boolean
'
'Dim NumError As Long
'Dim vImporte As Currency
'Dim vPorcIva As String
'
'Dim PrimFac As String
'Dim UltFac As String
'
'Dim tipoMov As String
'Dim b As Boolean
'Dim vSeccion As CSeccion
'Dim Importe As Currency
'
'Dim devuelve As String
'Dim Existe As Boolean
'
'Dim Neto As Currency
'Dim Recolect As Byte
'Dim vPrecio As Currency
'Dim PorcIva As Currency
'Dim vTipoMov As CTiposMov
'Dim NumFactu As Long
'Dim ImpoIva As Currency
'Dim BaseImpo As Currency
'Dim TotalFac As Currency
'
'
'Dim ConsumoHan As Currency
'Dim Acciones As Currency
'Dim Consumo1 As Long
'Dim Consumo2 As Long
'
'Dim ConsTra1 As Long
'Dim ConsTra2 As Long
'
'Dim Consumo As Long
'Dim ConsumoHidrante As Long
'
'
'    On Error GoTo eFacturacion
'
'    FacturacionConsumo = False
'
'    tipoMov = "RCP"
'
'    conn.BeginTrans
'
'
'    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
'    conn.Execute Sql
'
'    Sql = "SELECT codsocio,hidrante,nroorden,codparti,poligono,parcelas,hanegada,lect_ant,lect_act,fech_ant,fech_act"
'    Sql = Sql & " FROM  " & cTabla
'
'    If cWhere <> "" Then
'        cWhere = QuitarCaracterACadena(cWhere, "{")
'        cWhere = QuitarCaracterACadena(cWhere, "}")
'        cWhere = QuitarCaracterACadena(cWhere, "_1")
'        Sql = Sql & " WHERE " & cWhere
'    End If
'
'    ' ordenado por socio, hidrante
'    Sql = Sql & " order by rpozos.codsocio, rpozos.hidrante "
'
'    Set vSeccion = New CSeccion
'
'    If vSeccion.LeerDatos(vParamAplic.SeccionHorto) Then
'        If Not vSeccion.AbrirConta Then
'            Exit Function
'        End If
'    End If
'
'    b = True
'
'
'    vPorcIva = ""
'    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
'    PorcIva = CCur(ImporteSinFormato(vPorcIva))
'
'    Set vTipoMov = New CTiposMov
'
'    HayReg = False
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If Not Rs.EOF Then
'        AntSocio = CStr(DBLet(Rs!CodSocio, "N"))
'        ActSocio = CStr(DBLet(Rs!CodSocio, "N"))
'
'        BaseImpo = 0
'        ImpoIva = 0
'        TotalFac = 0
'
'        Sql2 = "select sum(acciones) acciones from rsocios_pozos where codsocio = " & DBSet(Rs!CodSocio, "N")
'        Acciones = DevuelveValor(Sql2)
'
'        Sql2 = "select rpozos.codsocio, sum(lect_act - lect_ant) consumo, sum(datediff(fech_act,fech_ant)) dias "
'        Sql2 = Sql2 & " from " & cTabla
'        If cWhere <> "" Then
'            Sql2 = Sql2 & " WHERE " & cWhere & " and "
'        Else
'            Sql2 = Sql2 & " WHERE "
'        End If
'        Sql2 = Sql2 & " rpozos.codsocio = " & DBSet(AntSocio, "N")
'        Sql2 = Sql2 & " group by 1 order by 1"
'
'        Set Rs2 = New ADODB.Recordset
'        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'
'        ConsumoHan = 0
'        If DBLet(Rs2!Consumo, "N") <> 0 And DBLet(Rs2!Dias, "N") <> 0 Then
'            ConsumoHan = Round2(((DBLet(Rs2!Consumo, "N") / Acciones) * 30) / DBLet(Rs2!Dias, "N"), 0)
'        End If
'
'        If ConsumoHan < CLng(txtcodigo(3).Text) Then
'            If ConsumoHan < CLng(txtcodigo(2).Text) Then
'                Consumo1 = DBLet(Rs2!Consumo, "N")
'                Consumo2 = 0
'            Else
'                Consumo1 = CLng(txtcodigo(2).Text) * (Acciones / 30 * DBLet(Rs2!Dias, "N"))
'                Consumo2 = DBLet(Rs2!Consumo) - Consumo1
'            End If
'        End If
'
'        Set Rs2 = Nothing
'
'    End If
'
'
'
'    While Not Rs.EOF And b
'        HayReg = True
'
'        ActSocio = Rs!CodSocio
'
'        If ActSocio <> AntSocio Then
'
'            Sql2 = "select sum(acciones) acciones from rsocios_pozos where codsocio = " & DBSet(ActSocio, "N")
'            Acciones = DevuelveValor(Sql2)
'
'            Sql2 = "select codsocio, sum(lect_act - lect_ant) consumo, sum(datediff(fech_act, fech_ant)) dias "
'            Sql2 = Sql2 & " from " & cTabla
'            If cWhere <> "" Then
'                Sql2 = Sql2 & " WHERE " & cWhere & " and "
'            Else
'                Sql2 = Sql2 & " WHERE "
'            End If
'            Sql2 = Sql2 & " codsocio = " & DBSet(ActSocio, "N")
'            Sql2 = Sql2 & " group by 1 order by 1 "
'
'            Set Rs2 = New ADODB.Recordset
'            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'            If DBLet(Rs2!Consumo, "N") <> 0 Then
'                ConsumoHan = Round2(((DBLet(Rs2!Consumo, "N") / Acciones) * 30) / DBLet(Rs2!Dias, "N"), 0)
'            End If
'
'            Consumo1 = 0
'            Consumo2 = 0
'
'            If ConsumoHan < CLng(txtcodigo(3).Text) Then
'                If ConsumoHan < CLng(txtcodigo(2).Text) Then
'                    Consumo1 = DBLet(Rs2!Consumo, "N")
'                    Consumo2 = 0
'                Else
'                    Consumo1 = CLng(txtcodigo(2).Text) * (Acciones / 30 * DBLet(Rs2!Dias, "N"))
'                    Consumo2 = DBLet(Rs2!Consumo) - Consumo1
'                End If
'            End If
'
'            Set Rs2 = Nothing
'
'            AntSocio = ActSocio
'
'        End If
'
'
'
'        NumFactu = vTipoMov.ConseguirContador(tipoMov)
'        Do
'            NumFactu = vTipoMov.ConseguirContador(tipoMov)
'            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFac, "F")
'            If devuelve <> "" Then
'                'Ya existe el contador incrementarlo
'                Existe = True
'                vTipoMov.IncrementarContador (tipoMov)
'                NumFactu = vTipoMov.ConseguirContador(tipoMov)
'            Else
'                Existe = False
'            End If
'        Loop Until Not Existe
'
'        ConsumoHidrante = DBLet(Rs!lect_act, "N") - DBLet(Rs!lect_ant, "N")
'        Consumo = ConsumoHidrante
'        ConsTra1 = 0
'        ConsTra2 = 0
'
'        If Consumo1 >= Consumo Then
'            ConsTra1 = Consumo
'            Consumo1 = Consumo1 - ConsTra1
'        Else
'            ConsTra1 = Consumo1
'            Consumo = Consumo - ConsTra1
'            If Consumo2 >= Consumo Then
'                ConsTra2 = Consumo
'                Consumo2 = Consumo2 - ConsTra2
'            End If
'        End If
'
'        TotalFac = Round2(ConsTra1 * CCur(ImporteSinFormato(txtcodigo(4).Text)), 2) + _
'                   Round2(ConsTra2 * CCur(ImporteSinFormato(txtcodigo(5).Text)), 2) + _
'                   vParamAplic.CuotaPOZ
'
'        IncrementarProgresNew Pb1, 1
'
'        'insertar en la tabla de recibos de pozos
'        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
'        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado) "
'        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(NumFactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(ActSocio, "N") & ","
'        Sql = Sql & DBSet(Rs!hidrante, "T") & "," & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
'        Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(ConsumoHidrante, "N") & "," & DBSet(vParamAplic.CuotaPOZ, "N") & ","
'        Sql = Sql & DBSet(Rs!lect_ant, "N") & "," & DBSet(Rs!fech_ant, "F") & ","
'        Sql = Sql & DBSet(Rs!lect_act, "N") & "," & DBSet(Rs!fech_act, "F") & ","
'        Sql = Sql & DBSet(ConsTra1, "N") & "," & DBSet(ImporteSinFormato(txtcodigo(4).Text), "N") & ","
'        Sql = Sql & DBSet(ConsTra2, "N") & "," & DBSet(ImporteSinFormato(txtcodigo(5).Text), "N") & ","
'        Sql = Sql & "'Recibo de Consumo',0)"
'
'        conn.Execute Sql
'
'        ' actualizar en los acumulados de hidrantes
'        Sql = "update rpozos set acumconsumo = acumconsumo + " & DBSet(ConsumoHidrante, "N")
'        Sql = Sql & ", acumcuota = acumcuota + " & DBSet(vParamAplic.CuotaPOZ, "N")
'        Sql = Sql & ", lect_ant = lect_act "
'        Sql = Sql & ", fech_ant = fech_act "
'        Sql = Sql & " WHERE hidrante = " & DBSet(Rs!hidrante, "T")
'
'        conn.Execute Sql
'
'
'        If b Then b = InsertResumen(tipoMov, CStr(NumFactu))
'
'        If b Then b = vTipoMov.IncrementarContador(tipoMov)
'
'        AntSocio = ActSocio
'
'        Rs.MoveNext
'    Wend
'
'    vSeccion.CerrarConta
'    Set vSeccion = Nothing
'
'eFacturacion:
'    If Err.Number <> 0 Or Not b Then
'        Mens = Mens & " " & Err.Description
'        conn.RollbackTrans
'        FacturacionConsumo = False
'    Else
'        conn.CommitTrans
'        FacturacionConsumo = True
'    End If
'End Function
'
'


Private Function CalculoConsumoHidrante(Hidrante As String, LectAct As Long, Consumo As Long) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Inicio As Long
Dim Fin As Long
Dim NroDig As Integer
Dim Limite As Long


    On Error GoTo eCalculoConsumoHidrante


    CalculoConsumoHidrante = False
    
    Sql = "select * from rpozos where hidrante = " & DBSet(Hidrante, "T")
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
       Inicio = 0
       Fin = 0
       NroDig = DBLet(RS!Digcontrol, "N")
       Limite = 10 ^ NroDig
       
       Inicio = DBLet(RS!lect_ant, "N")
       Fin = CLng(txtcodigo(51).Text)
    
       If Fin >= Inicio Then
          Consumo = Fin - Inicio
       Else
          If MsgBox("¿ Es un reinicio de contador ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Or (Inicio - Fin >= Limite) Then
              Consumo = (Limite - Inicio) + Fin
          Else
              Consumo = Fin - Inicio
          End If
       End If
    
       If Consumo > Limite - 1 Then
           MsgBox "Error en la lectura. Revise", vbExclamation
           CalculoConsumoHidrante = False
       Else
           If Consumo = 0 Then
                MsgBox "El consumo del contador es 0. Revise", vbExclamation
                CalculoConsumoHidrante = False
           Else
                CalculoConsumoHidrante = True
           End If
       End If
    End If
    Exit Function
   
   
eCalculoConsumoHidrante:
    MuestraError Err.Number, "Cálculo Consumo Hidrante", Err.Description
End Function
             
             
'En vTesoreria pondremos como estaba el recibo
'Es decir. El  msgbox que pondra al final lo guardo en esta variable
Private Function ComprobarCobroArimoney(vTesoreria As String, LEtra As String, Codfaccl As Long, Fecha As Date) As Boolean
Dim vR As ADODB.Recordset
Dim cad As String


On Error GoTo EComprobarCobroArimoney
    
    ComprobarCobroArimoney = False
    Set vR = New ADODB.Recordset
    cad = "Select * from scobro where numserie='" & LEtra & "'"
    cad = cad & " AND codfaccl =" & Codfaccl
    cad = cad & " AND fecfaccl =" & DBSet(Fecha, "F")
    
    '
    vTesoreria = ""
    vR.Open cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    If vR.EOF Then
        vTesoreria = "" '"NO se ha encotrado ningun vencimiento en la tesoreria"
    Else
        While Not vR.EOF
            cad = ""
            If DBLet(vR!codrem, "T") <> "" Then
                cad = "El cobro asociado a la factura esta remesado(" & vR!codrem & ")"
            Else
                If DBLet(vR!recedocu, "N") = 1 Then
                    cad = "Documento recibido"
                Else
                    If DBLet(vR!Estacaja, "N") = 1 Then
                        cad = "Cobrado por caja"
                    Else
                        If DBLet(vR!transfer, "N") = 1 Then
                            cad = "Esta en una transferencia"
                        Else
                           If DBLet(vR!impcobro, "N") > 0 Then cad = "Esta parcialmente cobrado: " & vR!impcobro
                        
                            
                                    'Si hubeira que poner mas coas iria aqui
                        End If 'transfer
                    End If 'estacaja
                End If 'recdedocu
            End If 'remesado
            If cad <> "" Then vTesoreria = vTesoreria & "Vto: " & vR!numorden & "      " & cad & vbCrLf
            vR.MoveNext
        Wend
    End If
    vR.Close
    
    ComprobarCobroArimoney = (vTesoreria <> "")
    
EComprobarCobroArimoney:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vR = Nothing
End Function



Private Sub ConexionConta()
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub

