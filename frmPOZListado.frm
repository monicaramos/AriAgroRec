VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPOZListado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   15075
   Icon            =   "frmPOZListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   15075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameReciboTalla 
      Height          =   5085
      Left            =   0
      TabIndex        =   238
      Top             =   30
      Width           =   6945
      Begin VB.Frame FrameCuota 
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   300
         TabIndex        =   301
         Top             =   2670
         Width           =   6495
         Begin VB.TextBox txtcodigo 
            Height          =   405
            Index           =   76
            Left            =   1590
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   302
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
            Top             =   150
            Width           =   4815
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   1350
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ver precios Zona"
            Top             =   660
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Precios Zona"
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
            Left            =   270
            TabIndex        =   304
            Top             =   660
            Width           =   915
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
            Index           =   67
            Left            =   270
            TabIndex        =   303
            Top             =   210
            Width           =   690
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen"
         Height          =   255
         Index           =   7
         Left            =   4590
         TabIndex        =   259
         Top             =   1950
         Width           =   1965
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Recibo"
         Height          =   195
         Index           =   6
         Left            =   4590
         TabIndex        =   258
         Top             =   2310
         Width           =   1995
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5790
         TabIndex        =   247
         Top             =   4455
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   244
         Top             =   4470
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   75
         Left            =   1995
         MaxLength       =   10
         TabIndex        =   242
         Top             =   1455
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   74
         Left            =   1995
         MaxLength       =   10
         TabIndex        =   241
         Top             =   1110
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   73
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   243
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
         TabIndex        =   240
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
         TabIndex        =   239
         Text            =   "Text5"
         Top             =   1470
         Width           =   3675
      End
      Begin MSComctlLib.ProgressBar pb4 
         Height          =   255
         Left            =   480
         TabIndex        =   260
         Top             =   4080
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame FrameBonif 
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   855
         Left            =   420
         TabIndex        =   253
         Top             =   1680
         Width           =   6255
         Begin VB.CheckBox Check1 
            Caption         =   "S�lo efectos"
            Height          =   195
            Index           =   8
            Left            =   4980
            TabIndex        =   261
            Top             =   420
            Width           =   1695
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   78
            Left            =   1470
            MaxLength       =   10
            TabIndex        =   245
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
            TabIndex        =   246
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bonificaci�n"
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
            Index           =   71
            Left            =   150
            TabIndex        =   257
            Top             =   390
            Width           =   840
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
            Left            =   2850
            TabIndex        =   256
            Top             =   390
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   69
            Left            =   2460
            TabIndex        =   255
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
            TabIndex        =   254
            Top             =   390
            Width           =   120
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Comprobando"
         Height          =   195
         Index           =   78
         Left            =   480
         TabIndex        =   262
         Top             =   4440
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   66
         Left            =   1020
         TabIndex        =   252
         Top             =   1110
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   65
         Left            =   1020
         TabIndex        =   251
         Top             =   1470
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
         Index           =   64
         Left            =   540
         TabIndex        =   250
         Top             =   870
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Generaci�n Recibos de Tallas"
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
         TabIndex        =   249
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
         TabIndex        =   248
         Top             =   1980
         Width           =   1005
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   1680
         Picture         =   "frmPOZListado.frx":000C
         Top             =   1980
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1680
         MouseIcon       =   "frmPOZListado.frx":0097
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1680
         MouseIcon       =   "frmPOZListado.frx":01E9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1470
         Width           =   240
      End
   End
   Begin VB.Frame FrameReimpresion 
      Height          =   5220
      Left            =   0
      TabIndex        =   113
      Top             =   0
      Width           =   6675
      Begin VB.Frame FrameTipoPago 
         Caption         =   "Tipo Pago"
         Enabled         =   0   'False
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
         Height          =   1215
         Left            =   3600
         TabIndex        =   267
         Top             =   1920
         Visible         =   0   'False
         Width           =   2145
         Begin VB.OptionButton Option1 
            Caption         =   "Contado"
            Height          =   225
            Index           =   2
            Left            =   420
            TabIndex        =   270
            Top             =   300
            Width           =   1005
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Efecto"
            Height          =   225
            Index           =   3
            Left            =   420
            TabIndex        =   269
            Top             =   570
            Width           =   1005
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   4
            Left            =   420
            TabIndex        =   268
            Top             =   840
            Width           =   885
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   1350
         Width           =   2070
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   119
         Tag             =   "N� Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1755
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1740
         MaxLength       =   7
         TabIndex        =   118
         Tag             =   "N� Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1365
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   123
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   121
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelReimp 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5340
         TabIndex        =   117
         Top             =   4275
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarReimp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   116
         Top             =   4275
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   127
         Top             =   3765
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   125
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
         TabIndex        =   115
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
         TabIndex        =   114
         Text            =   "Text5"
         Top             =   3390
         Width           =   3675
      End
      Begin VB.Label Label4 
         Height          =   195
         Index           =   52
         Left            =   450
         TabIndex        =   484
         Top             =   4410
         Width           =   3510
      End
      Begin VB.Label Label6 
         Caption         =   "Reimpresi�n de Facturas Socios"
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
         TabIndex        =   134
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   900
         TabIndex        =   133
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   132
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
         TabIndex        =   131
         Top             =   1110
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
         Index           =   16
         Left            =   465
         TabIndex        =   130
         Top             =   2115
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   825
         TabIndex        =   129
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   825
         TabIndex        =   128
         Top             =   2775
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1485
         Picture         =   "frmPOZListado.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   2415
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   1485
         Picture         =   "frmPOZListado.frx":03C6
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   855
         TabIndex        =   126
         Top             =   3405
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   870
         TabIndex        =   124
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
         TabIndex        =   122
         Top             =   3165
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1470
         MouseIcon       =   "frmPOZListado.frx":0451
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3765
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1470
         MouseIcon       =   "frmPOZListado.frx":05A3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3390
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   120
         Top             =   1110
         Width           =   1815
      End
   End
   Begin VB.Frame FrameAsignacionPrecios 
      Height          =   4545
      Left            =   0
      TabIndex        =   320
      Top             =   0
      Width           =   6645
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   120
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   328
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
         Top             =   3390
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   334
         Text            =   "Text5"
         Top             =   2970
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   71
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   327
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
         Top             =   2550
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   70
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   326
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
         Top             =   2130
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   96
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   332
         Text            =   "Text5"
         Top             =   1650
         Width           =   3045
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   96
         Left            =   2220
         MaxLength       =   4
         TabIndex        =   325
         Tag             =   "Zona|N|N|1|9999|rcampos|codzonas|0000||"
         Top             =   1650
         Width           =   555
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   95
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   331
         Text            =   "Text5"
         Top             =   1290
         Width           =   3045
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   95
         Left            =   2220
         MaxLength       =   4
         TabIndex        =   324
         Tag             =   "Zona|N|N|1|9999|rcampos|codzonas|0000||"
         Top             =   1290
         Width           =   555
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   4890
         TabIndex        =   330
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepAsigPrec 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   329
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Riego a Manta"
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
         Left            =   660
         TabIndex        =   483
         Top             =   3420
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�/hanegada"
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
         Index           =   103
         Left            =   3240
         TabIndex        =   482
         Top             =   3420
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�/hanegada"
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
         Index           =   77
         Left            =   3240
         TabIndex        =   339
         Top             =   2160
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�/hanegada"
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
         Left            =   3240
         TabIndex        =   338
         Top             =   2580
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL "
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
         Left            =   660
         TabIndex        =   337
         Top             =   3000
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Talla Ordinaria"
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
         Index           =   62
         Left            =   660
         TabIndex        =   336
         Top             =   2580
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Amortizacion Canal"
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
         Left            =   660
         TabIndex        =   335
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Zonas"
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
         Left            =   660
         TabIndex        =   333
         Top             =   1140
         Width           =   435
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1890
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Zona"
         Top             =   1650
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1890
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Zona"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   91
         Left            =   1230
         TabIndex        =   323
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   90
         Left            =   1230
         TabIndex        =   322
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label14 
         Caption         =   "Asignaci�n de precios"
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
         TabIndex        =   321
         Top             =   450
         Width           =   5250
      End
   End
   Begin VB.Frame FrameReciboConsumoManta 
      Height          =   5115
      Left            =   0
      TabIndex        =   375
      Top             =   0
      Width           =   6945
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   5640
         TabIndex        =   391
         Top             =   4395
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarRecManta 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4560
         TabIndex        =   389
         Top             =   4410
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   115
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   383
         Top             =   1200
         Width           =   960
      End
      Begin VB.Frame Frame17 
         BorderStyle     =   0  'None
         Height          =   1725
         Left            =   210
         TabIndex        =   377
         Top             =   1770
         Width           =   6375
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   114
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   384
            Top             =   150
            Width           =   1005
         End
         Begin VB.TextBox txtcodigo 
            Height          =   435
            Index           =   113
            Left            =   1650
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   387
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
            Top             =   1050
            Width           =   4725
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   112
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   385
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   570
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   9
            Left            =   4320
            TabIndex        =   379
            Top             =   120
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Ticket"
            Height          =   195
            Index           =   10
            Left            =   4320
            TabIndex        =   378
            Top             =   480
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Euros/Acci�n"
            Enabled         =   0   'False
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
            Left            =   330
            TabIndex        =   382
            Top             =   570
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ticket"
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
            Left            =   330
            TabIndex        =   381
            Top             =   -30
            Width           =   900
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   16
            Left            =   1350
            Picture         =   "frmPOZListado.frx":06F5
            Top             =   150
            Width           =   240
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
            Index           =   107
            Left            =   330
            TabIndex        =   380
            Top             =   960
            Width           =   690
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   115
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   376
         Text            =   "Text5"
         Top             =   1200
         Width           =   3375
      End
      Begin MSComctlLib.ProgressBar Pb7 
         Height          =   255
         Left            =   480
         TabIndex        =   386
         Top             =   3840
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
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
         Left            =   540
         TabIndex        =   390
         Top             =   1170
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Generaci�n Tickets Consumo a Manta"
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
         TabIndex        =   388
         Top             =   300
         Width           =   5925
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   22
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":0780
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame FrameRecConsPdtesCobro 
      Height          =   5850
      Left            =   0
      TabIndex        =   457
      Top             =   0
      Width           =   6675
      Begin VB.CheckBox Check3 
         Caption         =   "Para Excel "
         Height          =   255
         Left            =   450
         TabIndex        =   481
         Top             =   5100
         Width           =   2565
      End
      Begin VB.Frame Frame24 
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
         Height          =   615
         Left            =   450
         TabIndex        =   466
         Top             =   4350
         Width           =   3285
         Begin VB.OptionButton Option14 
            Caption         =   "Hidrante"
            Height          =   195
            Left            =   360
            TabIndex        =   470
            Top             =   270
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Socio"
            Height          =   195
            Left            =   1890
            TabIndex        =   468
            Top             =   270
            Width           =   1305
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   127
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   465
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3750
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   126
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   464
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3330
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   124
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   462
         Text            =   "Text5"
         Top             =   1230
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   125
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   460
         Text            =   "Text5"
         Top             =   1605
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   125
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   459
         Top             =   1620
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   124
         Left            =   1770
         MaxLength       =   6
         TabIndex        =   458
         Top             =   1230
         Width           =   830
      End
      Begin VB.CommandButton CmdAcepRecConsPdtes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4230
         TabIndex        =   467
         Top             =   5145
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   5400
         TabIndex        =   469
         Top             =   5130
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   123
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   463
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   122
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   461
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   45
         Left            =   810
         TabIndex        =   480
         Top             =   3750
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   44
         Left            =   810
         TabIndex        =   479
         Top             =   3390
         Width           =   465
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   53
         Left            =   510
         TabIndex        =   478
         Top             =   3120
         Width           =   825
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   32
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":08D2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   31
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":0A24
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
         Index           =   51
         Left            =   510
         TabIndex        =   477
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   50
         Left            =   870
         TabIndex        =   476
         Top             =   1620
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   855
         TabIndex        =   475
         Top             =   1245
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   1470
         Picture         =   "frmPOZListado.frx":0B76
         ToolTipText     =   "Buscar fecha"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   1470
         Picture         =   "frmPOZListado.frx":0C01
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   48
         Left            =   825
         TabIndex        =   474
         Top             =   2775
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   825
         TabIndex        =   473
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
         Index           =   46
         Left            =   510
         TabIndex        =   472
         Top             =   2115
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "Recibos Consumo Pdtes de Cobro"
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
         TabIndex        =   471
         Top             =   315
         Width           =   5160
      End
   End
   Begin VB.Frame FrameRecPdtesCobro 
      Height          =   6960
      Left            =   0
      TabIndex        =   392
      Top             =   -30
      Width           =   6675
      Begin VB.Frame Frame16 
         Caption         =   "Tipo de Listado"
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
         Height          =   615
         Left            =   480
         TabIndex        =   425
         Top             =   4890
         Width           =   5835
         Begin VB.OptionButton Option9 
            Caption         =   "Ambos"
            Height          =   195
            Left            =   4140
            TabIndex        =   428
            Top             =   300
            Width           =   1305
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Sector"
            Height          =   195
            Left            =   330
            TabIndex        =   427
            Top             =   300
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Bra�al"
            Height          =   195
            Left            =   2340
            TabIndex        =   426
            Top             =   300
            Width           =   1305
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   109
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   409
         Top             =   4575
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   108
         Left            =   1770
         MaxLength       =   6
         TabIndex        =   408
         Top             =   4200
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   109
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   422
         Text            =   "Text5"
         Top             =   4575
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   108
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   421
         Text            =   "Text5"
         Top             =   4200
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   107
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   405
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   106
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   404
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   5400
         TabIndex        =   411
         Top             =   6300
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepRecPdtesCob 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4230
         TabIndex        =   410
         Top             =   6315
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   105
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   403
         Top             =   1605
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   104
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   402
         Top             =   1230
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   105
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   401
         Text            =   "Text5"
         Top             =   1605
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   104
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   400
         Text            =   "Text5"
         Top             =   1230
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   103
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   407
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3690
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   102
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   406
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3315
         Width           =   1050
      End
      Begin VB.Frame Frame19 
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
         Height          =   615
         Left            =   480
         TabIndex        =   397
         Top             =   5610
         Width           =   3285
         Begin VB.OptionButton Option6 
            Caption         =   "Socio"
            Height          =   195
            Left            =   1920
            TabIndex        =   399
            Top             =   300
            Width           =   1305
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Sector/Bra�al"
            Height          =   195
            Left            =   330
            TabIndex        =   398
            Top             =   300
            Value           =   -1  'True
            Width           =   1305
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Tipo Pago"
         Enabled         =   0   'False
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
         Height          =   1215
         Left            =   4110
         TabIndex        =   393
         Top             =   2160
         Visible         =   0   'False
         Width           =   2145
         Begin VB.OptionButton Option1 
            Caption         =   "Contado"
            Height          =   225
            Index           =   10
            Left            =   420
            TabIndex        =   396
            Top             =   300
            Width           =   1005
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Efecto"
            Height          =   225
            Index           =   9
            Left            =   420
            TabIndex        =   395
            Top             =   570
            Width           =   1005
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   8
            Left            =   420
            TabIndex        =   394
            Top             =   840
            Width           =   885
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   840
         TabIndex        =   424
         Top             =   4230
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   840
         TabIndex        =   423
         Top             =   4590
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   1470
         MouseIcon       =   "frmPOZListado.frx":0C8C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar bra�al"
         Top             =   4590
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1470
         MouseIcon       =   "frmPOZListado.frx":0DDE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar bra�al"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Recibos Pendientes de Cobro"
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
         TabIndex        =   420
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
         Index           =   32
         Left            =   510
         TabIndex        =   419
         Top             =   2115
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   825
         TabIndex        =   418
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   30
         Left            =   825
         TabIndex        =   417
         Top             =   2775
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1470
         Picture         =   "frmPOZListado.frx":0F30
         ToolTipText     =   "Buscar fecha"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1470
         Picture         =   "frmPOZListado.frx":0FBB
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   29
         Left            =   855
         TabIndex        =   416
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   870
         TabIndex        =   415
         Top             =   1620
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
         Index           =   25
         Left            =   510
         TabIndex        =   414
         Top             =   1005
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":1046
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":1198
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Bra�al"
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
         Index           =   24
         Left            =   480
         TabIndex        =   413
         Top             =   3990
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Sector"
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
         Index           =   23
         Left            =   510
         TabIndex        =   412
         Top             =   3180
         Width           =   1815
      End
   End
   Begin VB.Frame FrameInfMantaFechaRiego 
      Height          =   5910
      Left            =   0
      TabIndex        =   429
      Top             =   0
      Width           =   6675
      Begin VB.Frame Frame21 
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
         Height          =   615
         Left            =   480
         TabIndex        =   454
         Top             =   4440
         Width           =   3555
         Begin VB.OptionButton Option1 
            Caption         =   "Fecha de Riego"
            Height          =   195
            Index           =   14
            Left            =   270
            TabIndex        =   456
            Top             =   300
            Value           =   -1  'True
            Width           =   1665
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Fecha de Pago"
            Height          =   195
            Index           =   15
            Left            =   1950
            TabIndex        =   455
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Tipo Pago"
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
         Height          =   1215
         Left            =   4020
         TabIndex        =   450
         Top             =   2160
         Width           =   2145
         Begin VB.OptionButton Option1 
            Caption         =   "Contado"
            Height          =   225
            Index           =   13
            Left            =   420
            TabIndex        =   453
            Top             =   300
            Width           =   1005
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Efecto"
            Height          =   225
            Index           =   12
            Left            =   420
            TabIndex        =   452
            Top             =   570
            Width           =   1005
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   11
            Left            =   420
            TabIndex        =   451
            Top             =   840
            Width           =   885
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   111
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   437
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3945
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   110
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   436
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3570
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   119
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   435
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   118
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   434
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   5400
         TabIndex        =   439
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepLisTicFecRiego 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4230
         TabIndex        =   438
         Top             =   5175
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   117
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   433
         Top             =   1605
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   116
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   432
         Top             =   1230
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   117
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   431
         Text            =   "Text5"
         Top             =   1605
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   116
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   430
         Text            =   "Text5"
         Top             =   1230
         Width           =   3675
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1515
         Picture         =   "frmPOZListado.frx":12EA
         ToolTipText     =   "Buscar fecha"
         Top             =   3945
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   43
         Left            =   840
         TabIndex        =   449
         Top             =   3930
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   1500
         Picture         =   "frmPOZListado.frx":1375
         ToolTipText     =   "Buscar fecha"
         Top             =   3555
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   36
         Left            =   825
         TabIndex        =   448
         Top             =   3570
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Pago"
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
         Index           =   35
         Left            =   510
         TabIndex        =   447
         Top             =   3270
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "Recibos por Fecha de Riego"
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
         TabIndex        =   446
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Riego"
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
         Index           =   42
         Left            =   510
         TabIndex        =   445
         Top             =   2115
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   41
         Left            =   825
         TabIndex        =   444
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   40
         Left            =   825
         TabIndex        =   443
         Top             =   2775
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   1500
         Picture         =   "frmPOZListado.frx":1400
         ToolTipText     =   "Buscar fecha"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   1500
         Picture         =   "frmPOZListado.frx":148B
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   39
         Left            =   855
         TabIndex        =   442
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   38
         Left            =   870
         TabIndex        =   441
         Top             =   1620
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
         Index           =   37
         Left            =   510
         TabIndex        =   440
         Top             =   1005
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":1516
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":1668
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
   End
   Begin VB.Frame FrameCuota2 
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   2655
      Left            =   8400
      TabIndex        =   275
      Top             =   2370
      Width           =   6525
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   293
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   72
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   292
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   291
         Text            =   "Text5"
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   79
         Left            =   570
         MaxLength       =   4
         TabIndex        =   290
         Tag             =   "Zona|N|N|1|9999|rcampos|codzonas|0000||"
         Top             =   1710
         Width           =   555
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   80
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   289
         Top             =   1710
         Width           =   945
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   81
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   288
         Top             =   1710
         Width           =   945
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   287
         Text            =   "Text5"
         Top             =   1710
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   79
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   286
         Text            =   "Text5"
         Top             =   1710
         Width           =   2055
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   82
         Left            =   570
         MaxLength       =   4
         TabIndex        =   285
         Tag             =   "Zona|N|N|1|9999|rcampos|codzonas|0000||"
         Top             =   2010
         Width           =   555
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   83
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   284
         Top             =   2010
         Width           =   945
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   84
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   283
         Top             =   2010
         Width           =   945
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   282
         Text            =   "Text5"
         Top             =   2010
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   82
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   281
         Text            =   "Text5"
         Top             =   2010
         Width           =   2055
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   85
         Left            =   570
         MaxLength       =   4
         TabIndex        =   280
         Tag             =   "Zona|N|N|1|9999|rcampos|codzonas|0000||"
         Top             =   2310
         Width           =   555
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   86
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   279
         Top             =   2310
         Width           =   945
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   87
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   278
         Top             =   2310
         Width           =   945
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   277
         Text            =   "Text5"
         Top             =   2310
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   85
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   276
         Text            =   "Text5"
         Top             =   2310
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cuota Talla Ordinaria"
         ForeColor       =   &H00972E0B&
         Height          =   390
         Index           =   58
         Left            =   4200
         TabIndex        =   300
         Top             =   600
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cuota Amortizacion Canal"
         ForeColor       =   &H00972E0B&
         Height          =   390
         Index           =   59
         Left            =   2940
         TabIndex        =   299
         Top             =   600
         Width           =   1530
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�/hanegada"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   72
         Left            =   3300
         TabIndex        =   298
         Top             =   990
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�/hanegada"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   73
         Left            =   4440
         TabIndex        =   297
         Top             =   990
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio Hda"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   74
         Left            =   5430
         TabIndex        =   296
         Top             =   930
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Zona General"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   79
         Left            =   2010
         TabIndex        =   295
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Otras Zonas"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   80
         Left            =   240
         TabIndex        =   294
         Top             =   1440
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   270
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Zona"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   270
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Zona"
         Top             =   2010
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   270
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Zona"
         Top             =   2310
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
   Begin VB.Frame FrameComprobacion 
      Height          =   3885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6945
      Begin VB.CommandButton CmdCancel 
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
         Left            =   540
         TabIndex        =   11
         Top             =   870
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Informe de Comprobaci�n de Lecturas"
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
         Left            =   570
         TabIndex        =   7
         Top             =   1800
         Width           =   435
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1545
         Picture         =   "frmPOZListado.frx":17BA
         Top             =   2325
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1560
         Picture         =   "frmPOZListado.frx":1845
         Top             =   1980
         Width           =   240
      End
   End
   Begin VB.Frame FrameReciboContador 
      Height          =   7725
      Left            =   0
      TabIndex        =   78
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
         TabIndex        =   111
         Top             =   6270
         Width           =   1185
      End
      Begin VB.Frame Frame4 
         Caption         =   "Art�culos"
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
         TabIndex        =   108
         Top             =   3900
         Width           =   7815
         Begin VB.TextBox txtcodigo 
            Height          =   315
            Index           =   31
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   96
            Text            =   "frmPOZListado.frx":18D0
            Top             =   1620
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   97
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
            TabIndex        =   94
            Text            =   "frmPOZListado.frx":1919
            Top             =   1260
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   95
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
            TabIndex        =   92
            Text            =   "frmPOZListado.frx":1962
            Top             =   900
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   28
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   93
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
            TabIndex        =   90
            Text            =   "frmPOZListado.frx":19AB
            Top             =   540
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   91
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
            TabIndex        =   110
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
            TabIndex        =   109
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
         TabIndex        =   105
         Top             =   2670
         Width           =   7785
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   6390
            MaxLength       =   10
            TabIndex        =   89
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
            TabIndex        =   88
            Text            =   "frmPOZListado.frx":19F4
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
            TabIndex        =   107
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
            TabIndex        =   106
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   6810
         TabIndex        =   100
         Top             =   7095
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarRecCont 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5700
         TabIndex        =   98
         Top             =   7110
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   86
         Top             =   1440
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   85
         Top             =   1080
         Width           =   960
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3165
         Left            =   210
         TabIndex        =   81
         Top             =   2010
         Width           =   7815
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   87
            Top             =   150
            Width           =   1005
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   5
            Left            =   4290
            TabIndex        =   83
            Top             =   60
            Width           =   1965
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Recibo"
            Height          =   195
            Index           =   4
            Left            =   4290
            TabIndex        =   82
            Top             =   420
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recibo"
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
            Left            =   330
            TabIndex        =   84
            Top             =   -30
            Width           =   960
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   1350
            Picture         =   "frmPOZListado.frx":1A3D
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
         TabIndex        =   80
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
         TabIndex        =   79
         Text            =   "Text5"
         Top             =   1080
         Width           =   3375
      End
      Begin MSComctlLib.ProgressBar Pb3 
         Height          =   255
         Left            =   510
         TabIndex        =   99
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
         TabIndex        =   112
         Top             =   6300
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   30
         Left            =   900
         TabIndex        =   104
         Top             =   1110
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   900
         TabIndex        =   103
         Top             =   1470
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
         Index           =   13
         Left            =   540
         TabIndex        =   102
         Top             =   870
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Generaci�n de Recibos Contadores"
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
         TabIndex        =   101
         Top             =   300
         Width           =   5925
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1590
         MouseIcon       =   "frmPOZListado.frx":1AC8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1590
         MouseIcon       =   "frmPOZListado.frx":1C1A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1080
         Width           =   240
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
            TabIndex        =   216
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
            TabIndex        =   215
            Top             =   1020
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Recargo"
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
            Index           =   46
            Left            =   3030
            TabIndex        =   208
            Top             =   1020
            Width           =   600
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
            Caption         =   "Bonificaci�n"
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
            Index           =   37
            Left            =   330
            TabIndex        =   201
            Top             =   1020
            Width           =   840
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
            Picture         =   "frmPOZListado.frx":1D6C
            Top             =   150
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recibo"
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
            Left            =   330
            TabIndex        =   53
            Top             =   -30
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Euros/Acci�n"
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
            Index           =   6
            Left            =   330
            TabIndex        =   52
            Top             =   570
            Width           =   930
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
      Begin VB.CommandButton CmdCancel 
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
         TabIndex        =   214
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   51
         Left            =   3270
         TabIndex        =   213
         Top             =   3570
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alta"
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
         Left            =   540
         TabIndex        =   212
         Top             =   3330
         Width           =   765
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   3810
         Picture         =   "frmPOZListado.frx":1DF7
         Top             =   3570
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   1560
         Picture         =   "frmPOZListado.frx":1E82
         Top             =   3540
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   49
         Left            =   540
         TabIndex        =   211
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   48
         Left            =   900
         TabIndex        =   210
         Top             =   1980
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   47
         Left            =   3270
         TabIndex        =   209
         Top             =   1980
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   45
         Left            =   3270
         TabIndex        =   207
         Top             =   3030
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   44
         Left            =   900
         TabIndex        =   206
         Top             =   3030
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
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
         Index           =   43
         Left            =   540
         TabIndex        =   205
         Top             =   2790
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   42
         Left            =   3270
         TabIndex        =   204
         Top             =   2490
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   41
         Left            =   900
         TabIndex        =   203
         Top             =   2490
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pol�gono"
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
         Index           =   40
         Left            =   540
         TabIndex        =   202
         Top             =   2250
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":1F0D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":205F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Generaci�n de Recibos Mantenimiento"
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
         Left            =   540
         TabIndex        =   47
         Top             =   870
         Width           =   375
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
         Tag             =   "N� Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1665
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   0
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "N� Parte|N|S|||rpartes|nroparte|0000000|S|"
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
      Begin VB.CommandButton CmdCancel 
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
         Left            =   600
         TabIndex        =   21
         Top             =   1080
         Width           =   615
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
         Caption         =   "L�nea 3"
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
         Index           =   34
         Left            =   570
         TabIndex        =   168
         Top             =   1920
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "L�nea 2"
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
         Index           =   39
         Left            =   570
         TabIndex        =   167
         Top             =   1500
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "Nro.Etiquetas"
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
         Index           =   0
         Left            =   540
         TabIndex        =   163
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "L�nea 1"
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
         Index           =   36
         Left            =   570
         TabIndex        =   161
         Top             =   1080
         Width           =   510
      End
   End
   Begin VB.Frame FrameRectificacion 
      Height          =   4680
      Left            =   0
      TabIndex        =   183
      Top             =   0
      Width           =   6675
      Begin VB.Frame Frame9 
         Caption         =   "Datos para Selecci�n"
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
         TabIndex        =   187
         Top             =   870
         Width           =   6315
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   55
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   191
            Top             =   1350
            Width           =   1065
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   4080
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   196
            Top             =   450
            Width           =   2070
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   52
            Left            =   1620
            MaxLength       =   7
            TabIndex        =   189
            Tag             =   "N� Factura|N|S|||facturas|numfactu|0000000|S|"
            Top             =   450
            Width           =   1035
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   56
            Left            =   1620
            MaxLength       =   6
            TabIndex        =   190
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
            TabIndex        =   188
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
            TabIndex        =   200
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
            TabIndex        =   199
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
            TabIndex        =   198
            Top             =   900
            Width           =   375
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   1320
            MouseIcon       =   "frmPOZListado.frx":21B1
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   900
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Height          =   255
            Index           =   28
            Left            =   2850
            TabIndex        =   197
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
         TabIndex        =   192
         Top             =   3060
         Width           =   1065
      End
      Begin VB.CommandButton CmdAceptarRectif 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   194
         Top             =   3915
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelRectif 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5340
         TabIndex        =   195
         Top             =   3915
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   54
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   193
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
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
         TabIndex        =   186
         Top             =   3060
         Width           =   540
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   1560
         Picture         =   "frmPOZListado.frx":2303
         ToolTipText     =   "Buscar fecha"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
         Index           =   22
         Left            =   510
         TabIndex        =   185
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Rectificaci�n de Lecturas"
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
         TabIndex        =   184
         Top             =   405
         Width           =   5160
      End
   End
   Begin VB.Frame Frame10 
      Height          =   3495
      Left            =   0
      TabIndex        =   263
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5340
         TabIndex        =   266
         Top             =   2580
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   264
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Cambio de Zonas de Campos"
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
         Left            =   780
         TabIndex        =   265
         Top             =   720
         Width           =   5925
      End
   End
   Begin VB.Frame FrameCartaTallas 
      Height          =   8415
      Left            =   -30
      TabIndex        =   217
      Top             =   30
      Width           =   6945
      Begin VB.TextBox txtcodigo 
         Height          =   405
         Index           =   97
         Left            =   2010
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   221
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
         Top             =   2520
         Width           =   4725
      End
      Begin VB.Frame Frame13 
         Height          =   900
         Left            =   360
         TabIndex        =   317
         Top             =   6510
         Width           =   6405
         Begin VB.OptionButton OptMail 
            Caption         =   "Enviar por e-mail e imprimir a los socios sin correo"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   319
            Top             =   360
            Width           =   3885
         End
         Begin VB.OptionButton OptMail 
            Caption         =   "Imprimir Todos"
            Height          =   255
            Index           =   1
            Left            =   4680
            TabIndex        =   318
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "e-Mail"
         Enabled         =   0   'False
         Height          =   780
         Left            =   4980
         TabIndex        =   313
         Top             =   6630
         Visible         =   0   'False
         Width           =   1755
         Begin VB.OptionButton OptMailCom 
            Caption         =   "Compras"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   315
            Top             =   450
            Width           =   1335
         End
         Begin VB.OptionButton OptMailAdm 
            Caption         =   "Administraci�n"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   314
            Top             =   210
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Datos de Impresi�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2955
         Left            =   360
         TabIndex        =   305
         Top             =   3480
         Width           =   6375
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   94
            Left            =   1650
            MaxLength       =   35
            TabIndex        =   228
            Top             =   2520
            Width           =   4365
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   93
            Left            =   1650
            MaxLength       =   35
            TabIndex        =   227
            Top             =   2160
            Width           =   4365
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   92
            Left            =   1650
            MaxLength       =   35
            TabIndex        =   226
            Top             =   1800
            Width           =   4365
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   91
            Left            =   1650
            MaxLength       =   40
            TabIndex        =   225
            Top             =   1440
            Width           =   4365
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   90
            Left            =   1650
            MaxLength       =   40
            TabIndex        =   224
            Top             =   1080
            Width           =   4365
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   89
            Left            =   1650
            MaxLength       =   40
            TabIndex        =   223
            Top             =   720
            Width           =   4365
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   88
            Left            =   1650
            MaxLength       =   40
            TabIndex        =   222
            Top             =   360
            Width           =   4365
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Recargos"
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
            Left            =   210
            TabIndex        =   312
            Top             =   2520
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Per�odo voluntario"
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
            Left            =   210
            TabIndex        =   311
            Top             =   2160
            Width           =   1305
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bonificaci�n"
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
            Index           =   86
            Left            =   210
            TabIndex        =   310
            Top             =   1800
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Prohibici�n"
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
            Left            =   210
            TabIndex        =   309
            Top             =   1440
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Fin Comunic."
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
            Index           =   84
            Left            =   210
            TabIndex        =   308
            Top             =   1080
            Width           =   1395
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio Pago"
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
            Index           =   83
            Left            =   210
            TabIndex        =   307
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Junta"
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
            Index           =   82
            Left            =   210
            TabIndex        =   306
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   68
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   237
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
         TabIndex        =   236
         Text            =   "Text5"
         Top             =   1110
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   69
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   220
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   67
         Left            =   1995
         MaxLength       =   10
         TabIndex        =   218
         Top             =   1110
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   68
         Left            =   1995
         MaxLength       =   10
         TabIndex        =   219
         Top             =   1455
         Width           =   960
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   4290
         TabIndex        =   229
         Top             =   7740
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5400
         TabIndex        =   230
         Top             =   7725
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb5 
         Height          =   255
         Left            =   360
         TabIndex        =   342
         Top             =   7440
         Visible         =   0   'False
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
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
         Index           =   93
         Left            =   570
         TabIndex        =   341
         Top             =   2580
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precios Zona"
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
         Left            =   570
         TabIndex        =   340
         Top             =   3030
         Width           =   915
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1650
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ver precios Zona"
         Top             =   3030
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   420
         TabIndex        =   316
         Top             =   6450
         Width           =   3705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1680
         MouseIcon       =   "frmPOZListado.frx":238E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1470
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1680
         MouseIcon       =   "frmPOZListado.frx":24E0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1680
         Picture         =   "frmPOZListado.frx":2632
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
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
         Left            =   570
         TabIndex        =   235
         Top             =   1980
         Width           =   960
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
         TabIndex        =   234
         Top             =   300
         Width           =   5925
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
         Index           =   57
         Left            =   540
         TabIndex        =   233
         Top             =   870
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   56
         Left            =   1020
         TabIndex        =   232
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   55
         Left            =   1020
         TabIndex        =   231
         Top             =   1110
         Width           =   465
      End
   End
   Begin VB.Frame FrameComprobacionDatos 
      Height          =   5685
      Left            =   30
      TabIndex        =   343
      Top             =   0
      Width           =   6825
      Begin VB.Frame Frame15 
         BorderStyle     =   0  'None
         Caption         =   "Diferencias con Indefa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   4035
         Left            =   210
         TabIndex        =   346
         Top             =   780
         Width           =   6465
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   98
            Left            =   1590
            MaxLength       =   10
            TabIndex        =   354
            Top             =   540
            Width           =   960
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   99
            Left            =   1605
            MaxLength       =   10
            TabIndex        =   355
            Top             =   900
            Width           =   960
         End
         Begin VB.Frame Frame14 
            Caption         =   "Tipo Informe"
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
            Height          =   2625
            Left            =   150
            TabIndex        =   347
            Top             =   1290
            Width           =   6165
            Begin VB.OptionButton Option4 
               Caption         =   "Diferencias entre Datos de Contadores y Campos"
               Height          =   435
               Index           =   0
               Left            =   420
               TabIndex        =   353
               Top             =   1272
               Width           =   3885
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Diferencias entre Escalona e Indefa"
               Height          =   435
               Index           =   1
               Left            =   420
               TabIndex        =   352
               Top             =   255
               Width           =   3885
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Contadores que existen en Indefa y no en Escalona"
               Height          =   435
               Index           =   2
               Left            =   420
               TabIndex        =   351
               Top             =   594
               Width           =   5415
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Contadores con Socio Bloqueado"
               Height          =   435
               Index           =   4
               Left            =   420
               TabIndex        =   350
               Top             =   1611
               Width           =   3885
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Contadores con consumo que est�n en Inelcom y no est�n en Escalona"
               Height          =   435
               Index           =   5
               Left            =   420
               TabIndex        =   349
               Top             =   1950
               Width           =   5655
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Contadores que existen en Escalona y no en Indefa"
               Height          =   435
               Index           =   3
               Left            =   420
               TabIndex        =   348
               Top             =   933
               Width           =   5415
            End
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   94
            Left            =   690
            TabIndex        =   360
            Top             =   570
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   95
            Left            =   690
            TabIndex        =   359
            Top             =   930
            Width           =   420
         End
         Begin VB.Label Label2 
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
            Index           =   96
            Left            =   330
            TabIndex        =   358
            Top             =   330
            Width           =   615
         End
      End
      Begin VB.CommandButton CmdAceptarComp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4410
         TabIndex        =   356
         Top             =   5130
         Width           =   945
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5520
         TabIndex        =   357
         Top             =   5115
         Width           =   945
      End
      Begin MSComctlLib.ProgressBar pb6 
         Height          =   255
         Left            =   300
         TabIndex        =   361
         Top             =   4800
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Height          =   195
         Index           =   97
         Left            =   360
         TabIndex        =   345
         Top             =   4860
         Width           =   3825
      End
      Begin VB.Label Label15 
         Caption         =   "Informe de Comprobaci�n de Datos"
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
         TabIndex        =   344
         Top             =   300
         Width           =   5925
      End
   End
   Begin VB.Frame FrameComprobacionCCC 
      Height          =   3255
      Left            =   0
      TabIndex        =   362
      Top             =   0
      Width           =   6825
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   100
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   373
         Text            =   "Text5"
         Top             =   1290
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   101
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   372
         Text            =   "Text5"
         Top             =   1650
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   101
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   368
         Top             =   1650
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   100
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   367
         Top             =   1290
         Width           =   960
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5520
         TabIndex        =   364
         Top             =   2475
         Width           =   945
      End
      Begin VB.CommandButton CmdAceptarCompCCC 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4410
         TabIndex        =   363
         Top             =   2490
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Procesando"
         Height          =   195
         Index           =   102
         Left            =   570
         TabIndex        =   374
         Top             =   2220
         Visible         =   0   'False
         Width           =   4395
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":26BD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":280F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1320
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
         Index           =   100
         Left            =   570
         TabIndex        =   371
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   99
         Left            =   930
         TabIndex        =   370
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   98
         Left            =   930
         TabIndex        =   369
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label16 
         Caption         =   "Informe de Cuentas Bancarias Err�neas"
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
         TabIndex        =   366
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         Height          =   195
         Index           =   101
         Left            =   360
         TabIndex        =   365
         Top             =   4860
         Width           =   3825
      End
   End
   Begin VB.Frame FrameReciboConsumo 
      Height          =   6285
      Left            =   0
      TabIndex        =   57
      Top             =   30
      Width           =   6945
      Begin VB.Frame Frame99 
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   240
         TabIndex        =   75
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
               TabIndex        =   65
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
               Left            =   330
               TabIndex        =   170
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
            TabIndex        =   171
            Top             =   480
            Width           =   3555
            Begin VB.TextBox txtcodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   1290
               MaxLength       =   10
               TabIndex        =   175
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
               TabIndex        =   174
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
               TabIndex        =   173
               Top             =   480
               Width           =   1005
            End
            Begin VB.TextBox txtcodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   5
               Left            =   2400
               MaxLength       =   10
               TabIndex        =   172
               Top             =   780
               Width           =   1005
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Rango Consumo"
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
               Index           =   4
               Left            =   60
               TabIndex        =   178
               Top             =   180
               Width           =   1170
            End
            Begin VB.Label Label2 
               Caption         =   "Hasta m3"
               Height          =   195
               Index           =   5
               Left            =   1290
               TabIndex        =   177
               Top             =   300
               Width           =   945
            End
            Begin VB.Label Label2 
               Caption         =   "Precio"
               Height          =   195
               Index           =   7
               Left            =   2430
               TabIndex        =   176
               Top             =   300
               Width           =   945
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recibo"
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
            Left            =   330
            TabIndex        =   76
            Top             =   -30
            Width           =   960
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1305
            Picture         =   "frmPOZListado.frx":2961
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
         TabIndex        =   66
         Top             =   5610
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5640
         TabIndex        =   67
         Top             =   5595
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   540
         TabIndex        =   77
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
         Picture         =   "frmPOZListado.frx":29EC
         Top             =   2370
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1530
         Picture         =   "frmPOZListado.frx":2A77
         Top             =   1980
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
         Index           =   22
         Left            =   570
         TabIndex        =   74
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   930
         TabIndex        =   73
         Top             =   2370
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   930
         TabIndex        =   72
         Top             =   2010
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Generaci�n de Recibos de Consumo"
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
         TabIndex        =   71
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
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
         Index           =   16
         Left            =   540
         TabIndex        =   70
         Top             =   870
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   15
         Left            =   900
         TabIndex        =   69
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   900
         TabIndex        =   68
         Top             =   1110
         Width           =   465
      End
   End
   Begin VB.Frame FrameFacturasHidrante 
      Height          =   6030
      Left            =   0
      TabIndex        =   136
      Top             =   0
      Width           =   6675
      Begin VB.Frame FrameTipoPago2 
         Caption         =   "Tipo Pago"
         Enabled         =   0   'False
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
         Height          =   1215
         Left            =   4110
         TabIndex        =   271
         Top             =   2310
         Visible         =   0   'False
         Width           =   2145
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   7
            Left            =   420
            TabIndex        =   274
            Top             =   840
            Width           =   885
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Efecto"
            Height          =   225
            Index           =   6
            Left            =   420
            TabIndex        =   273
            Top             =   570
            Width           =   1005
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Contado"
            Height          =   225
            Index           =   5
            Left            =   420
            TabIndex        =   272
            Top             =   300
            Width           =   1005
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ordenado por"
         Enabled         =   0   'False
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
         Height          =   615
         Left            =   480
         TabIndex        =   180
         Top             =   5130
         Width           =   3285
         Begin VB.OptionButton Option3 
            Caption         =   "Nro.Factura"
            Height          =   195
            Left            =   1800
            TabIndex        =   182
            Top             =   270
            Width           =   1305
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Socio"
            Height          =   195
            Left            =   210
            TabIndex        =   181
            Top             =   270
            Value           =   -1  'True
            Width           =   1305
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Resumen Facturaci�n"
         Height          =   285
         Left            =   510
         TabIndex        =   144
         Top             =   4740
         Width           =   2175
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   50
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   142
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3825
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   141
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3450
         Width           =   1050
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   143
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
         TabIndex        =   147
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
         TabIndex        =   137
         Top             =   1230
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   138
         Top             =   1605
         Width           =   830
      End
      Begin VB.CommandButton cmdAceptarListFact 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4230
         TabIndex        =   145
         Top             =   5355
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelList 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5400
         TabIndex        =   146
         Top             =   5340
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   139
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   140
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1050
      End
      Begin VB.Label Label4 
         Caption         =   "Nro. Factura"
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
         Index           =   6
         Left            =   510
         TabIndex        =   179
         Top             =   3180
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo Factura"
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
         MouseIcon       =   "frmPOZListado.frx":2B02
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":2C54
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
         Picture         =   "frmPOZListado.frx":2DA6
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   1485
         Picture         =   "frmPOZListado.frx":2E31
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
    
    ' 3 .- Generaci�n Recibos de Consumo (Facturacion de consumo)
    ' 4 .- Generaci�n Recibos de Mantenimiento (Factura de Mantenimiento)
    ' 5 .- Generacion Recibos de Contadores ( Factura de Contadores )
    
    ' 6 .- Reimpresion de recibos de pozos
    ' 7 .- Listado de consumo por hidrante
    
    ' 8 .- Etiquetas contadores
    ' 9 .- Facturas rectificativas
    
    ' 10.- Listado de tallas, recibos de talla (solo para Escalona)
    ' 11.- Generacion de recibos de talla
    ' 12.- C�lculo de bonificacion de Recibos de Talla (solo para Escalona)
    
    ' 13.-
    ' 14.- Asignacion de Precios de Talla
    
    ' 15.- Listado de Diferencias con Indefa
    ' 16.- Listado de cuentas bancarias de socios err�neas
    
    ' 17.- Generacion de recibos a manta
    
    ' 18.- Informe de recibos pendientes de cobro por bra�al y por sector
    ' 19.- Informe de recibos de riego a manta por fecha de riego
    ' 20.- Informe de recibos de consumo pendientes de cobro
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmSoc  As frmManSocios 'mantenimiento de socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmZon  As frmManZonas 'mantenimiento de zonas
Attribute frmZon.VB_VarHelpID = -1
 
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'ayuda de hidrantes por socio
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMens2 As frmMensajes 'ayuda de hidrantes por socio
Attribute frmMens2.VB_VarHelpID = -1
Private WithEvents frmMens3 As frmMensajes 'ayuda de hidrantes por socio
Attribute frmMens3.VB_VarHelpID = -1
Private WithEvents frmMens4 As frmMensajes 'ayuda de campos por socio a facturar a manta
Attribute frmMens4.VB_VarHelpID = -1

Private WithEvents frmMen As frmMensajes  'Form Mensajes para mostrar las etiquetas a imprimir
Attribute frmMen.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos
 
Dim IndRptReport As Integer
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String
Dim vSeccion As CSeccion


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean

Dim PrecioTalla1 As Currency
Dim PrecioTalla2 As Currency
Dim ZonaTalla As Long

Dim HayFacturas As Boolean
Dim Continuar As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check2_Click()
    Me.Frame7.Enabled = (Check2.Value = 1)
End Sub

Private Sub CmdAcepAsigPrec_Click()
Dim SQL As String

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

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H ZONA
    cDesde = Trim(txtcodigo(95).Text)
    cHasta = Trim(txtcodigo(96).Text)
    nDesde = txtNombre(95).Text
    nHasta = txtNombre(96).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codzonas}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHZona=""") Then Exit Sub
    End If
    
    nTabla = "rzonas"
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        If cadSelect = "" Then cadSelect = "(1=1)"
        If Not BloqueaRegistro(nTabla, cadSelect) Then
            MsgBox "No se pueden Actualizar precios. Hay registros bloqueados.", vbExclamation
            Screen.MousePointer = vbDefault
        Else
            If ProcesarCambios(nTabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (8)
            End If
        End If
    End If

End Sub

Private Function ProcesarCambios(nTabla As String, cadSelect As String) As Boolean
Dim vSQL As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim Importe As String
Dim SQL As String
Dim Sql1 As String

Dim Albaran As Long
Dim linea As Long

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

    vSQL = "update rzonas set precio1 =  " & DBSet(txtcodigo(70).Text, "N")
    vSQL = vSQL & ", precio2 = " & DBSet(txtcodigo(71).Text, "N")
    vSQL = vSQL & ", preciomanta = " & DBSet(txtcodigo(120).Text, "N")
    If cadSelect <> "" Then vSQL = vSQL & " where " & cadSelect

    conn.Execute vSQL
    
       
    conn.CommitTrans
    ProcesarCambios = True
    Exit Function
    
eProcesarCambios:
    conn.RollbackTrans
    MuestraError Err.Number, "Procesar Cambios", Err.Description
End Function


Private Sub CmdAcepLisTicFecRiego_Click()
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
    
    Tabla = "rpozticketsmanta"
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtcodigo(116).Text)
    cHasta = Trim(txtcodigo(117).Text)
    nDesde = txtNombre(116).Text
    nHasta = txtNombre(117).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha riego
    cDesde = Trim(txtcodigo(118).Text)
    cHasta = Trim(txtcodigo(119).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecriego}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    If Not AnyadirAFormula(cadSelect, "not " & Tabla & ".fecriego is null") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "not isnull({" & Tabla & ".fecriego})") Then Exit Sub
    
    '++
    '[Monica]25/09/2014: a�adimos la fecha de pago y el tipo de pago que es
    'D/H Fecha pago
    cDesde = Trim(txtcodigo(110).Text)
    cHasta = Trim(txtcodigo(111).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecpago}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFechaP= """) Then Exit Sub
    End If
    
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        'tengo que seleccionar que albaranes vamos a listar dependiendo de si son o no contado o todos
        'insertamos en tmpinformes
        
        If Not CargarAlbaranes(cadSelect) Then Exit Sub
        
'        Tabla = Tabla & " INNER JOIN rsocios ON rpozticketsmanta.codsocio = rsocios.codsocio "
'
'        '[Monica]23/11/2012: si es escalona o utxera seleccionamos que tipo de socio
'        If Option1(11).Value Then    ' solo contado
'            If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
'            If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
'        End If
'        If Option1(12).Value Then    ' solo efecto
'            If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba<>""8888888888""") Then Exit Sub
'            If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}<>""8888888888""") Then Exit Sub
'        End If
    End If
    '++
    
    
    If HayRegistros("tmpinformes", "codusu = " & vUsu.Codigo) Then
        indRPT = 105
        ConSubInforme = False
        cadTitulo = "Recibos por Fecha de Riego"
        
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
          
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
          
        If Option1(13).Value Then CadParam = CadParam & "pTipo=""Contado""|"
        If Option1(12).Value Then CadParam = CadParam & "pTipo=""Banco""|"
        If Option1(11).Value Then CadParam = CadParam & "pTipo=""Todos""|"
        
        numParam = numParam + 1
          
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
        
        If Option1(15).Value Then
            cadNombreRPT = Replace(cadNombreRPT, ".rpt", "FPag.rpt")
        End If
        
        LlamarImprimir
    End If

End Sub


Private Function CargarAlbaranes(cWhere As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarAlbaranes

    Screen.MousePointer = vbHourglass


    CargarAlbaranes = False

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = "Select " & vUsu.Codigo & ",numalbar, fecalbar FROM rpozticketsmanta inner join rsocios on rpozticketsmanta.codsocio = rsocios.codsocio "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    If Option1(11).Value Then
        ' no hacemos nada
    
    Else
        If cWhere <> "" Then
            SQL = SQL & " and "
        Else
            SQL = SQL & " where "
        End If
        
        If Option1(13).Value Then ' contado
            SQL = SQL & " ((numalbar, fecalbar) in (select numalbar, fecalbar from rrecibpozos where codtipom = 'RMT' and escontado = 1)  or        "
            SQL = SQL & " (not (numalbar, fecalbar) in (select numalbar, fecalbar from rrecibpozos where codtipom = 'RMT') and rsocios.cuentaba='8888888888'))"
        Else ' banco
            SQL = SQL & " ((numalbar, fecalbar) in (select numalbar, fecalbar from rrecibpozos where codtipom = 'RMT' and escontado = 0)  or       "
            SQL = SQL & " (not (numalbar, fecalbar) in (select numalbar, fecalbar from rrecibpozos where codtipom = 'RMT') and rsocios.cuentaba<>'8888888888'))"
        End If
    End If
    
    Sql2 = "insert into tmpinformes (codusu, importe1, fecha1) "
    conn.Execute Sql2 & SQL
    
    Screen.MousePointer = vbDefault
    CargarAlbaranes = True
    Exit Function
    
eCargarAlbaranes:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Cargar Albaranes", Err.Description
End Function

Private Sub CmdAcepRecConsPdtes_Click()
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
Dim SQL As String
Dim Sql3 As String
Dim Sql2 As String
Dim ctabla1 As String
Dim Cad As String

Dim SqlZonas As String
    
Dim cadSelect1 As String
Dim cadFormula1 As String
Dim Cadena As String
    
Dim CadSelect0 As String
Dim SqlZonas0 As String
Dim Sql1 As String
Dim NConta As Integer
    
    InicializarVbles
    
    NConta = DevuelveValor("select empresa_conta from rseccion where codsecci = " & vParamAplic.SeccionPOZOS)
    
    ' recibos de consumo
    ctabla1 = "conta" & NConta & ".scobro cc, " & vEmpresa.BDAriagro & ".rrecibpozos rr, " & vEmpresa.BDAriagro & ".rsocios ss, usuarios.stipom tt "
    
'[Monica]07/01/2015: cambiamos la condicion
'    Sql1 = "where ((cc.codforpa = 1 and (cc.codrem is null or cc.codrem = 0)) or (cc.codforpa = 0))"
    Sql1 = "where ((cc.impvenci + coalesce(cc.gastos,0) - coalesce(cc.impcobro,0) <> 0)) "
    Sql1 = Sql1 & " and cc.impvenci > 0"
    Sql1 = Sql1 & " and rr.codtipom = 'RCP'"
    Sql1 = Sql1 & " and rr.codtipom = tt.codtipom "
    Sql1 = Sql1 & " and cc.numserie = tt.letraser "
    Sql1 = Sql1 & " and cc.codfaccl = rr.numfactu"
    Sql1 = Sql1 & " and cc.fecfaccl = rr.fecfactu"
    Sql1 = Sql1 & " and mid(cc.codmacta,5,6) = ss.codsocio"
    
    '
    If Not DatosOk Then Exit Sub
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H Socio
    If txtcodigo(124).Text <> "" Then
        Sql1 = Sql1 & " and rr.codsocio >= " & DBSet(txtcodigo(124).Text, "N")
    End If
    If txtcodigo(125).Text <> "" Then
        Sql1 = Sql1 & " and rr.codsocio <= " & DBSet(txtcodigo(125).Text, "N")
    End If
    If txtcodigo(124).Text <> "" Or txtcodigo(125).Text <> "" Then
        Cad = ""
        If txtcodigo(124).Text <> "" Then Cad = Cad & " DESDE: " & txtcodigo(124).Text & " " & txtNombre(124).Text
        If txtcodigo(125).Text <> "" Then Cad = Cad & "  HASTA: " & txtcodigo(125).Text & " " & txtNombre(125).Text
        CadParam = CadParam & "pDHSocio=""" & Cad & """|"
        numParam = numParam + 1
    End If
    
    'D/H fecha
    If txtcodigo(122).Text <> "" Then
        Sql1 = Sql1 & " and rr.fecfactu >= " & DBSet(txtcodigo(122).Text, "F")
    End If
    If txtcodigo(123).Text <> "" Then
        Sql1 = Sql1 & " and rr.fecfactu <= " & DBSet(txtcodigo(123).Text, "F")
    End If
    If txtcodigo(122).Text <> "" Or txtcodigo(123).Text <> "" Then
        Cad = ""
        If txtcodigo(122).Text <> "" Then Cad = Cad & " DESDE: " & txtcodigo(122).Text
        If txtcodigo(123).Text <> "" Then Cad = Cad & "  HASTA: " & txtcodigo(123).Text
        CadParam = CadParam & "pDHFecha=""" & Cad & """|"
        numParam = numParam + 1
    End If

    ' hidrante
    If txtcodigo(126).Text <> "" Then Sql1 = Sql1 & " and rr.hidrante >= " & DBSet(txtcodigo(126).Text, "N")
    If txtcodigo(127).Text <> "" Then Sql1 = Sql1 & " and rr.hidrante <= " & DBSet(txtcodigo(127).Text, "N")
    If txtcodigo(102).Text <> "" Or txtcodigo(103).Text <> "" Then
        Cad = ""
        If txtcodigo(126).Text <> "" Then Cad = Cad & " DESDE: " & txtcodigo(126).Text
        If txtcodigo(127).Text <> "" Then Cad = Cad & "  HASTA: " & txtcodigo(127).Text
        CadParam = CadParam & "pDHHidrante=""" & Cad & """|"
        numParam = numParam + 1
    End If
    
    If CargarTemporalRecibosConsumoPdtes(ctabla1, Sql1, NConta) Then
        If HayRegParaInforme("tmpinformes", "tmpinformes.codusu = " & vUsu.Codigo) Then
        
            cadTitulo = "Recibos Consumo Pendientes de Cobro"

        
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            cadNombreRPT = "EscPOZRecConsPdtesCob.rpt"
    
            CadParam = CadParam & "pUsu=" & vUsu.Codigo & "|"
            If Me.Option13.Value Then
                CadParam = CadParam & "pTipo=1|" 'pTipo = 1 por socio
                                                 '      = 0 por hidrante
            Else
                CadParam = CadParam & "pTipo=0|"
            End If
            numParam = numParam + 2
            
            CadParam = CadParam & "pExcel=" & Check3.Value & "|"
            numParam = numParam + 1
            
            
            ConSubInforme = True
            LlamarImprimir
        
        End If
    End If
    

End Sub

Private Sub CmdAcepRecPdtesCob_Click()
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
Dim SQL As String
Dim Sql3 As String
Dim Sql2 As String
Dim ctabla1 As String
Dim Cad As String

Dim SqlZonas As String
    
Dim cadSelect1 As String
Dim cadFormula1 As String
Dim Cadena As String
    
Dim CadSelect0 As String
Dim SqlZonas0 As String
Dim Sql1 As String
Dim NConta As Integer
    
    InicializarVbles
    
    NConta = DevuelveValor("select empresa_conta from rseccion where codsecci = " & vParamAplic.SeccionPOZOS)
    
    ' recibos de talla
    cTabla = "conta" & NConta & ".scobro cc, " & vEmpresa.BDAriagro & ".rrecibpozos rr, " & vEmpresa.BDAriagro & ".rrecibpozos_cam ll, " & vEmpresa.BDAriagro & ".rzonas zz, " & vEmpresa.BDAriagro & ".rcampos cam, " & vEmpresa.BDAriagro & ".rsocios ss, usuarios.stipom tt"
    
'[Monica]07/01/2015: cambiamos la condicion
'    SQL = "where ((cc.codforpa = 1 and (cc.codrem is null or cc.codrem = 0)) or (cc.codforpa = 0))"
    SQL = "where ((cc.impvenci + coalesce(cc.gastos,0) - coalesce(cc.impcobro,0) <> 0)) "
    SQL = SQL & " and cc.impvenci > 0 "
    SQL = SQL & " and rr.codtipom = ll.codtipom "
    SQL = SQL & " and rr.numfactu = ll.numfactu "
    SQL = SQL & " and rr.fecfactu = ll.fecfactu "
    SQL = SQL & " and mid(cc.codmacta,5,6) = rr.codsocio"
    SQL = SQL & " and rr.codtipom = tt.codtipom "
    SQL = SQL & " and cc.numserie = tt.letraser "
    SQL = SQL & " and cc.codfaccl = rr.numfactu"
    SQL = SQL & " and cc.fecfaccl = rr.fecfactu"
    SQL = SQL & " and ll.codcampo = cam.codcampo"
    SQL = SQL & " and rr.codtipom = 'TAL'"
    SQL = SQL & " and  mid(cc.codmacta,5,6) = ss.codsocio"
    SQL = SQL & " and cam.codzonas = zz.codzonas"
    
    
'[Monica]07/01/2015: cambiamos la condicion
'    Sql2 = "where ((cc.codforpa = 1 and (cc.codrem is null or cc.codrem = 0)) or (cc.codforpa = 0))"
    Sql2 = "where ((cc.impvenci + coalesce(cc.gastos,0) - coalesce(cc.impcobro,0) <> 0)) "
    Sql2 = Sql2 & " and cc.impvenci > 0 "
    Sql2 = Sql2 & " and rr.codtipom = ll.codtipom "
    Sql2 = Sql2 & " and rr.numfactu = ll.numfactu "
    Sql2 = Sql2 & " and rr.fecfactu = ll.fecfactu "
    Sql2 = Sql2 & " and mid(cc.codmacta,5,6) = rr.codsocio"
    Sql2 = Sql2 & " and rr.codtipom = tt.codtipom "
    Sql2 = Sql2 & " and cc.numserie = tt.letraser "
    Sql2 = Sql2 & " and cc.codfaccl = rr.numfactu"
    Sql2 = Sql2 & " and cc.fecfaccl = rr.fecfactu"
    Sql2 = Sql2 & " and ll.codcampo = cam.codcampo"
    Sql2 = Sql2 & " and rr.codtipom = 'RMT'"
    Sql2 = Sql2 & " and  mid(cc.codmacta,5,6) = ss.codsocio"
    Sql2 = Sql2 & " and cam.codzonas = zz.codzonas"
    
    
    ' recibos de consumo
    ctabla1 = "conta" & NConta & ".scobro cc, " & vEmpresa.BDAriagro & ".rrecibpozos rr, " & vEmpresa.BDAriagro & ".rsocios ss, usuarios.stipom tt"
    
'[Monica]07/01/2015: cambiamos la condicion
'    Sql1 = "where ((cc.codforpa = 1 and (cc.codrem is null or cc.codrem = 0)) or (cc.codforpa = 0))"
    Sql1 = "where ((cc.impvenci + coalesce(cc.gastos,0) - coalesce(cc.impcobro,0) <> 0)) "
    Sql1 = Sql1 & " and cc.impvenci > 0"
    Sql1 = Sql1 & " and rr.codtipom = 'RCP'"
    Sql1 = Sql1 & " and rr.codtipom = tt.codtipom "
    Sql1 = Sql1 & " and cc.numserie = tt.letraser "
    Sql1 = Sql1 & " and cc.codfaccl = rr.numfactu"
    Sql1 = Sql1 & " and cc.fecfaccl = rr.fecfactu"
    Sql1 = Sql1 & " and mid(cc.codmacta,5,6) = ss.codsocio"
    
    
    '
    
    If Not DatosOk Then Exit Sub
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H Socio
    If txtcodigo(104).Text <> "" Then
        SQL = SQL & " and rr.codsocio >= " & DBSet(txtcodigo(104).Text, "N")
        Sql1 = Sql1 & " and rr.codsocio >= " & DBSet(txtcodigo(104).Text, "N")
        Sql2 = Sql2 & " and rr.codsocio >= " & DBSet(txtcodigo(104).Text, "N")
    End If
    If txtcodigo(105).Text <> "" Then
        SQL = SQL & " and rr.codsocio <= " & DBSet(txtcodigo(105).Text, "N")
        Sql1 = Sql1 & " and rr.codsocio <= " & DBSet(txtcodigo(105).Text, "N")
        Sql2 = Sql2 & " and rr.codsocio <= " & DBSet(txtcodigo(105).Text, "N")
    End If
    If txtcodigo(104).Text <> "" Or txtcodigo(105).Text <> "" Then
        Cad = ""
        If txtcodigo(104).Text <> "" Then Cad = Cad & " DESDE: " & txtcodigo(104).Text & " " & txtNombre(104).Text
        If txtcodigo(105).Text <> "" Then Cad = Cad & "  HASTA: " & txtcodigo(105).Text & " " & txtNombre(105).Text
        CadParam = CadParam & "pDHSocio=""" & Cad & """|"
        numParam = numParam + 1
    End If
    
    'D/H fecha
    If txtcodigo(106).Text <> "" Then
        SQL = SQL & " and rr.fecfactu >= " & DBSet(txtcodigo(106).Text, "F")
        Sql1 = Sql1 & " and rr.fecfactu >= " & DBSet(txtcodigo(106).Text, "F")
        Sql2 = Sql2 & " and rr.fecfactu >= " & DBSet(txtcodigo(106).Text, "F")
    End If
    If txtcodigo(107).Text <> "" Then
        SQL = SQL & " and rr.fecfactu <= " & DBSet(txtcodigo(107).Text, "F")
        Sql1 = Sql1 & " and rr.fecfactu <= " & DBSet(txtcodigo(107).Text, "F")
        Sql2 = Sql2 & " and rr.fecfactu <= " & DBSet(txtcodigo(107).Text, "F")
    End If
    If txtcodigo(106).Text <> "" Or txtcodigo(107).Text <> "" Then
        Cad = ""
        If txtcodigo(107).Text <> "" Then Cad = Cad & " DESDE: " & txtcodigo(106).Text
        If txtcodigo(108).Text <> "" Then Cad = Cad & "  HASTA: " & txtcodigo(107).Text
        CadParam = CadParam & "pDHFecha=""" & Cad & """|"
        numParam = numParam + 1
    End If

    ' bra�al
    If txtcodigo(108).Text <> "" Then
        SQL = SQL & " and cam.codzonas >= " & DBSet(txtcodigo(108).Text, "N")
        Sql2 = Sql2 & " and cam.codzonas >= " & DBSet(txtcodigo(108).Text, "N")
    End If
    If txtcodigo(109).Text <> "" Then
        SQL = SQL & " and cam.codzonas <= " & DBSet(txtcodigo(109).Text, "N")
        Sql2 = Sql2 & " and cam.codzonas <= " & DBSet(txtcodigo(109).Text, "N")
    End If
    If txtcodigo(108).Text <> "" Or txtcodigo(109).Text <> "" Then
        Cad = ""
        If txtcodigo(108).Text <> "" Then Cad = Cad & " DESDE: " & txtcodigo(108).Text & " " & txtNombre(108).Text
        If txtcodigo(109).Text <> "" Then Cad = Cad & "  HASTA: " & txtcodigo(109).Text & " " & txtNombre(109).Text
        CadParam = CadParam & "pDHZona=""" & Cad & """|"
        numParam = numParam + 1
    End If

    ' sector
    If txtcodigo(102).Text <> "" Then Sql1 = Sql1 & " and mid(rr.hidrante,1,2) >= " & DBSet(txtcodigo(102).Text, "N")
    If txtcodigo(103).Text <> "" Then Sql1 = Sql1 & " and mid(rr.hidrante,1,2) <= " & DBSet(txtcodigo(103).Text, "N")
    If txtcodigo(102).Text <> "" Or txtcodigo(103).Text <> "" Then
        Cad = ""
        If txtcodigo(102).Text <> "" Then Cad = Cad & " DESDE: " & txtcodigo(102).Text
        If txtcodigo(103).Text <> "" Then Cad = Cad & "  HASTA: " & txtcodigo(103).Text
        CadParam = CadParam & "pDHSector=""" & Cad & """|"
        numParam = numParam + 1
    End If
    
    If Option7.Value Then CadParam = CadParam & "pTipo=2|" 'sector
    If Option8.Value Then CadParam = CadParam & "pTipo=1|" 'bra�al
    If Option9.Value Then CadParam = CadParam & "pTipo=0|" 'ambos
    numParam = numParam + 1
    
    If CargarTemporalRecibosPdtes(cTabla, SQL, Sql2, ctabla1, Sql1) Then
        If HayRegParaInforme("tmpinformes", "tmpinformes.codusu = " & vUsu.Codigo) Then
        
            cadTitulo = "Recibos Pendientes de Cobro"

        
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            If Option7.Value Then cadFormula = cadFormula & " and {tmpinformes.campo1} = 2" 'sector
            If Option8.Value Then cadFormula = cadFormula & " and {tmpinformes.campo1} = 1" 'bra�al
            
            cadNombreRPT = "EscPOZrecPdtesCob.rpt"
    
            If Me.Option6.Value Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "Soc.rpt")
        
            CadParam = CadParam & "pUsu=" & vUsu.Codigo & "|"
            numParam = numParam + 1
            
            ConSubInforme = True
            LlamarImprimir
        
        End If
    End If
    

End Sub

Private Sub CmdAceptarComp_Click()
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
Dim SQL As String
Dim Sql3 As String

Dim SqlZonas As String
    
Dim cadSelect1 As String
Dim cadFormula1 As String
Dim Cadena As String
    
Dim CadSelect0 As String
Dim SqlZonas0 As String
    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '======== FORMULA  ====================================
    'D/H Hidrante
    cDesde = Trim(txtcodigo(98).Text)
    cHasta = Trim(txtcodigo(99).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.hidrante}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHHidrante=""") Then Exit Sub
    End If
    
    indRPT = 95
    ConSubInforme = False
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    
    cadNombreRPT = nomDocu
    
    If Option4(0).Value Then cadTitulo = "Diferencias entre Datos de Contadores y Campos"
    If Option4(1).Value Then cadTitulo = "Diferencias entre Escalona e Indefa"
    If Option4(2).Value Then cadTitulo = "Contadores que existen en Indefa y no en Escalona"
    If Option4(3).Value Then cadTitulo = "Contadores que existen en Escalona y no en Indefa"
    If Option4(4).Value Then cadTitulo = "Contadores con Socio Bloqueado"
    If Option4(5).Value Then cadTitulo = "Contadores con consumo en Inelcom y no en Escalona"
    
    If CargarTemporalDiferencias(Tabla, cadSelect) Then
        If HayRegParaInforme("tmpinformes", "tmpinformes.codusu = " & vUsu.Codigo) Then
        
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
    
            If Me.Option4(1).Value Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "1.rpt")
            If Me.Option4(2).Value Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "2.rpt")
            If Me.Option4(3).Value Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "3.rpt")
            If Me.Option4(4).Value Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "4.rpt")
            If Me.Option4(5).Value Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "5.rpt")
            
        
            CadParam = CadParam & "pUsu=" & vUsu.Codigo & "|"
            numParam = numParam + 1
            
            ConSubInforme = True
            LlamarImprimir
        End If
    End If


End Sub

Private Function CargarTemporalDiferencias(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Cad2 As String
Dim Cad3 As String
Dim CadValues As String
Dim cadInsert As String
Dim Contador As String
Dim Nregs As Integer
Dim Fecha As Date

    On Error GoTo eCargarTemporal
    
    CargarTemporalDiferencias = False
    
    Screen.MousePointer = vbHourglass
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    If Me.Option4(0).Value Then
                                                'tipo,   h.contador,h.socio,c.socio, h.campo, c.campo,h.poligono,c.polig,  h.parce, c.parcela  h.hda,  c.hda
        SQL = "insert into tmpinformes (codusu, codigo1, nombre1, importe1, importe2,importe3,importe4, nombre2, importe5, nombre3, importeb1, precio1, precio2) "
    
        SQL = SQL & "SELECT " & vUsu.Codigo & ",0 tipo, rpozos.hidrante,rpozos.codsocio,rcampos.codsocio, rpozos.codcampo, rcampos.codcampo, rpozos.poligono, rcampos.poligono, rpozos.parcelas, rcampos.parcela, rpozos.hanegada, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4)"
        SQL = SQL & " FROM rpozos,rcampos "
        SQL = SQL & "  WHERE rpozos.poligono=rcampos.poligono AND rpozos.parcelas=rcampos.parcela AND rpozos.codcampo<>rcampos.codcampo "
        SQL = SQL & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtcodigo(98).Text <> "" Then SQL = SQL & " and rpozos.hidrante >= " & DBSet(txtcodigo(98).Text, "T")
        If txtcodigo(99).Text <> "" Then SQL = SQL & " and rpozos.hidrante <= " & DBSet(txtcodigo(99).Text, "T")
        SQL = SQL & " union "
        SQL = SQL & "SELECT " & vUsu.Codigo & ",1 tipo, rpozos.hidrante,rpozos.codsocio,rcampos.codsocio, rpozos.codcampo, rcampos.codcampo, rpozos.poligono, rcampos.poligono, rpozos.parcelas, rcampos.parcela, rpozos.hanegada, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4)"
        SQL = SQL & " FROM rpozos,rcampos "
        SQL = SQL & " WHERE  rpozos.codcampo=rcampos.codcampo AND (rpozos.poligono<>rcampos.poligono)  "
        SQL = SQL & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtcodigo(98).Text <> "" Then SQL = SQL & " and rpozos.hidrante >= " & DBSet(txtcodigo(98).Text, "T")
        If txtcodigo(99).Text <> "" Then SQL = SQL & " and rpozos.hidrante <= " & DBSet(txtcodigo(99).Text, "T")
        SQL = SQL & " union "
        SQL = SQL & "SELECT " & vUsu.Codigo & ",2 tipo, rpozos.hidrante,rpozos.codsocio,rcampos.codsocio, rpozos.codcampo, rcampos.codcampo, rpozos.poligono, rcampos.poligono, rpozos.parcelas, rcampos.parcela, rpozos.hanegada, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4)"
        SQL = SQL & " FROM rpozos,rcampos "
        SQL = SQL & " WHERE  rpozos.codcampo=rcampos.codcampo AND rpozos.codsocio <> rcampos.codsocio "
        SQL = SQL & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtcodigo(98).Text <> "" Then SQL = SQL & " and rpozos.hidrante >= " & DBSet(txtcodigo(98).Text, "T")
        If txtcodigo(99).Text <> "" Then SQL = SQL & " and rpozos.hidrante <= " & DBSet(txtcodigo(99).Text, "T")
        SQL = SQL & " union "
        SQL = SQL & "SELECT " & vUsu.Codigo & ",3 tipo, rpozos.hidrante,rpozos.codsocio,rcampos.codsocio, rpozos.codcampo, rcampos.codcampo, rpozos.poligono, rcampos.poligono, rpozos.parcelas, rcampos.parcela, rpozos.hanegada, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4)"
        SQL = SQL & " FROM rpozos,rcampos "
        SQL = SQL & " WHERE  rpozos.codcampo=rcampos.codcampo and "
        SQL = SQL & " truncate(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4),0) <> truncate(rpozos.hanegada,0) "
        SQL = SQL & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtcodigo(98).Text <> "" Then SQL = SQL & " and rpozos.hidrante >= " & DBSet(txtcodigo(98).Text, "T")
        If txtcodigo(99).Text <> "" Then SQL = SQL & " and rpozos.hidrante <= " & DBSet(txtcodigo(99).Text, "T")
        SQL = SQL & " union "
        SQL = SQL & "SELECT " & vUsu.Codigo & ",4 tipo, rpozos.hidrante,rpozos.codsocio,rcampos.codsocio, rpozos.codcampo, rcampos.codcampo, rpozos.poligono, rcampos.poligono, rpozos.parcelas, rcampos.parcela, rpozos.hanegada, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4)"
        SQL = SQL & " FROM rpozos,rcampos "
        SQL = SQL & " WHERE  rpozos.codcampo=rcampos.codcampo AND (rpozos.parcelas<>rcampos.parcela)  "
        SQL = SQL & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtcodigo(98).Text <> "" Then SQL = SQL & " and rpozos.hidrante >= " & DBSet(txtcodigo(98).Text, "T")
        If txtcodigo(99).Text <> "" Then SQL = SQL & " and rpozos.hidrante <= " & DBSet(txtcodigo(99).Text, "T")
        
        conn.Execute SQL
    
    End If
    
    ' listado de discrepancias indefa
    If Me.Option4(1).Value Then
        If AbrirConexionIndefa() = False Then
            MsgBox "No se ha podido acceder a los datos de Indefa. ", vbExclamation
            Exit Function
        End If
        
                                            '   h.contador,h.poligono,h.parcelas, h.hdas     h.socio_revisado toma
        cadInsert = "insert into tmpinformes (codusu,  nombre1, nombre2, nombre3,   precio1, importe1,      importe2) values "

        SQL = "SELECT " & vUsu.Codigo & ", rpozos.hidrante,rpozos.poligono,rpozos.parcelas,rpozos.hanegada,rpozos.codsocio, rpozos.nroorden "
        SQL = SQL & " FROM rpozos "
        SQL = SQL & "  WHERE length(hidrante) = 6 and cast(hidrante as unsigned) "
        SQL = SQL & " and (rpozos.fechabaja is null or rpozos.fechabaja = '')"
        If txtcodigo(98).Text <> "" Then SQL = SQL & " and rpozos.hidrante >= " & DBSet(txtcodigo(98).Text, "T")
        If txtcodigo(99).Text <> "" Then SQL = SQL & " and rpozos.hidrante <= " & DBSet(txtcodigo(99).Text, "T")

        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        Nregs = TotalRegistrosConsulta(SQL)
        If Nregs <> 0 Then
            pb6.visible = True
            Label2(97).visible = True
            CargarProgres pb6, Nregs
            DoEvents
        End If

        CadValues = ""
        While Not Rs.EOF
            IncrementarProgres pb6, 1
            
            Contador = DBLet(Rs!Hidrante, "T")
            
            Label2(97).Caption = "Procesando contador: " & Contador
            DoEvents
                
            Sql2 = "select poligono, parcelas, hanegadas, socio_revisado, toma "
            Sql2 = Sql2 & " from rae_visitas_hidtomas where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
            Sql2 = Sql2 & " and hidrante = " & DBSet(Int(Mid(Contador, 3, 2)), "T")
            '[Monica]18/07/2013:
                                    '[Monica]27/01/2014: lo cambio a numerico
            Sql2 = Sql2 & " and salida_tch = " & DBSet(Int(Mid(Contador, 5, 2)), "N")


            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs2.EOF Then
                Cad3 = "(" & vUsu.Codigo & "," & DBSet(Rs!Hidrante, "T") & ","
                Cad2 = ""
                If Trim(DBLet(Rs!poligono, "T")) <> Trim(DBLet(Rs2!poligono, "T")) Then
                    If DBLet(Rs2!poligono, "T") = "" Then
                        Cad2 = Cad2 & "'',"
                    Else
                        Cad2 = Cad2 & DBSet(Rs2!poligono, "T") & ","
                    End If
                Else
                    Cad2 = Cad2 & ValorNulo & ","
                End If
                If Mid(Trim(DBLet(Rs!parcelas, "T")), 1, 25) <> Mid(Trim(DBLet(Rs2!parcelas, "T")), 1, 25) Then
                    If DBLet(Rs2!parcelas, "T") = "" Then
                        Cad2 = Cad2 & "'',"
                    Else
                        Cad2 = Cad2 & DBSet(Rs2!parcelas, "T") & ","
                    End If
                Else
                    Cad2 = Cad2 & ValorNulo & ","
                End If
                If Int(ComprobarCero(DBLet(Rs!hanegada, "N"))) <> Int(Round2(ComprobarCero(DBLet(Rs2!Hanegadas, "N")), 4)) Then
                    If DBLet(Rs2!Hanegadas, "N") = 0 Then
                        Cad2 = Cad2 & "0,"
                    Else
                        Cad2 = Cad2 & DBSet(Rs2!Hanegadas, "N") & ","
                    End If
                Else
                    Cad2 = Cad2 & ValorNulo & ","
                End If
                If CLng(DBLet(Rs!Codsocio, "N")) <> CLng(ComprobarCero(DBLet(Rs2!socio_revisado, "N"))) And CLng(ComprobarCero(DBLet(Rs2!socio_revisado, "N"))) <> 0 Then
                    Cad2 = Cad2 & DBSet(Rs2!socio_revisado, "N") & ","
                Else
                    Cad2 = Cad2 & ValorNulo & ","
                End If
                
                '[Monica]30/10/2013: a�adimos la parte de la toma
                If (ComprobarCero(DBLet(Rs!nroorden, "N")) Mod 100) <> ComprobarCero(DBLet(Rs2!toma, "N")) Then
                    If DBLet(Rs2!toma, "N") = 0 Then
                        Cad2 = Cad2 & "0,"
                    Else
                        Cad2 = Cad2 & DBSet(Rs2!toma, "N") & ","
                    End If
                Else
                    Cad2 = Cad2 & ValorNulo & ","
                End If


                If Cad2 <> "Null,Null,Null,Null,Null," Then
                    CadValues = CadValues & Cad3 & Mid(Cad2, 1, Len(Cad2) - 1) & "),"
                End If

            End If
            Set Rs2 = Nothing
            Rs.MoveNext
        Wend
        Set Rs = Nothing

        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute cadInsert & CadValues
        End If
        CerrarConexionIndefa
        pb6.visible = False
        Label2(97).visible = False
    End If
    
    'listado de contadores que existen en indefa y no en Escalona
    If Me.Option4(2).Value Then
        If AbrirConexionIndefa() = False Then
            MsgBox "No se ha podido acceder a los datos de Indefa. ", vbExclamation
            Exit Function
        End If
                                            '         h.contador,poligono,parcelas,hanegadas, indicamos si tiene fecha de baja
        cadInsert = "insert into tmpinformes (codusu,  nombre1, importe1, nombre2, precio1, importe2) values "
        
        CadValues = ""
        '[Monica]18/07/2013
        Sql2 = "select sector, hidrante, salida_tch, poligono, parcelas, hanegadas "
        Sql2 = Sql2 & " from rae_visitas_hidtomas "
        Sql2 = Sql2 & " where sector < '8' "
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
            
        Nregs = TotalRegistrosIndefa("select count(*) from rae_visitas_hidtomas")
        If Nregs <> 0 Then
            pb6.visible = True
            Label2(97).visible = True
            CargarProgres pb6, Nregs
            DoEvents
        End If
            
        While Not Rs2.EOF
            '[Monica]18/07/2013
            Contador = Format(Rs2!sector, "00") & Format(Rs2!Hidrante, "00") & Format(Rs2!salida_tch, "00")
            
            IncrementarProgres pb6, 1
            Label2(97).Caption = "Procesando contador: " & Contador
            DoEvents
            
            SQL = "select count(*) from rpozos where hidrante = " & DBSet(Contador, "T")
            If TotalRegistros(SQL) = 0 Then
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Contador, "T") & ","
                CadValues = CadValues & DBSet(Rs2!poligono, "N") & "," & DBSet(Rs2!parcelas, "T") & ","
                CadValues = CadValues & DBSet(Rs2!Hanegadas, "N") & ",0),"
            Else
                ' estan en escalona pero tienen fecha de baja pongo una marca para identificarlos
                SQL = "select count(*) from rpozos where hidrante = " & DBSet(Contador, "T") & " and not fechabaja is null"
                If TotalRegistros(SQL) = 1 Then
                    CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Contador, "T") & ","
                    CadValues = CadValues & DBSet(Rs2!poligono, "N") & "," & DBSet(Rs2!parcelas, "T") & ","
                    CadValues = CadValues & DBSet(Rs2!Hanegadas, "N") & ",1),"
                End If
                
            End If
            
            Rs2.MoveNext
        Wend
        
        Set Rs2 = Nothing
        
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute cadInsert & CadValues
        End If
        CerrarConexionIndefa
        pb6.visible = False
        Label2(97).visible = False
    End If
        
        
    'listado de contadores que existen en Escalona y no en Indefa
    If Me.Option4(3).Value Then
        If AbrirConexionIndefa() = False Then
            MsgBox "No se ha podido acceder a los datos de Indefa. ", vbExclamation
            Exit Function
        End If
                                            '   h.contador
        cadInsert = "insert into tmpinformes (codusu,  nombre1) values "
        
        CadValues = ""
        
        SQL = "select hidrante from rpozos where length(hidrante) = 6 and cast(hidrante as unsigned) "
        SQL = SQL & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtcodigo(98).Text <> "" Then SQL = SQL & " and rpozos.hidrante >= " & DBSet(txtcodigo(98).Text, "T")
        If txtcodigo(99).Text <> "" Then SQL = SQL & " and rpozos.hidrante <= " & DBSet(txtcodigo(99).Text, "T")
        SQL = SQL & " order by hidrante "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Nregs = TotalRegistrosConsulta(SQL)
        If Nregs <> 0 Then
            pb6.visible = True
            Label2(97).visible = True
            CargarProgres pb6, Nregs
            DoEvents
        End If
        
        While Not Rs.EOF
            IncrementarProgres pb6, 1
            Label2(97).Caption = "Procesando contador: " & Rs!Hidrante
            DoEvents
         
            Contador = Rs!Hidrante
         
            Sql2 = "select gid "
            Sql2 = Sql2 & " from rae_visitas_hidtomas "
            Sql2 = Sql2 & " where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
            Sql2 = Sql2 & " and hidrante = " & DBSet(Int(Mid(Contador, 3, 2)), "T")
            '[Monica]18/07/2013
                                    '[Monica]27/01/2014: lo cambio a numerico
            Sql2 = Sql2 & " and salida_tch = " & DBSet(Int(Mid(Contador, 5, 2)), "N")
        
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Rs2.EOF Then
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Contador, "T") & "),"
            End If
            Set Rs2 = Nothing
            
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
        
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute cadInsert & CadValues
        End If
        CerrarConexionIndefa
        pb6.visible = False
        Label2(97).visible = False
    End If
        
        
    'listado de contadores con socio bloqueado
    If Me.Option4(4).Value Then
                                            '   h.contador
        SQL = "insert into tmpinformes (codusu,  nombre1) "
        SQL = SQL & " select " & vUsu.Codigo & ",hidrante"
        SQL = SQL & " from rpozos "
        SQL = SQL & " where codsocio in (select codsocio from rsocios where codsitua > 1)"
        SQL = SQL & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "

        If txtcodigo(98).Text <> "" Then SQL = SQL & " and rpozos.hidrante >= " & DBSet(txtcodigo(98).Text, "T")
        If txtcodigo(99).Text <> "" Then SQL = SQL & " and rpozos.hidrante <= " & DBSet(txtcodigo(99).Text, "T")
        
        conn.Execute SQL
    End If
        
        
    'listado de contadores con consumo y no existencia en escalona
    If Me.Option4(5).Value Then
    
        Fecha = DevuelveValor("select max(fecproceso) from rpozos_lectura")
                                            
                                            '    h.contador consumo
        SQL = "insert into tmpinformes (codusu,  nombre1, importe1) "
        SQL = SQL & "select " & vUsu.Codigo & ", contador, "
        If vParamAplic.TipoLecturaPoz Then
            SQL = SQL & "lectura_bd "
        Else
            SQL = SQL & "lectura_equipo "
        End If
        SQL = SQL & " from rpozos_lectura "
        SQL = SQL & " where "
        If vParamAplic.TipoLecturaPoz Then
            SQL = SQL & " lectura_bd <> 0"
        Else
            SQL = SQL & " lectura_equipo <> 0 "
        End If
        
        SQL = SQL & " and (fecproceso is null or fecproceso = " & DBSet(Fecha, "F") & ")"
        SQL = SQL & " and not right(concat('00',contador),6) in (select hidrante from rpozos where (1=1) "
        If txtcodigo(98).Text <> "" Then SQL = SQL & " and rpozos.hidrante >= " & DBSet(txtcodigo(98).Text, "T")
        If txtcodigo(99).Text <> "" Then SQL = SQL & " and rpozos.hidrante <= " & DBSet(txtcodigo(99).Text, "T")
        SQL = SQL & ")"
  
        conn.Execute SQL
        
        
    End If
        
    CargarTemporalDiferencias = True
    Screen.MousePointer = vbDefault

    Exit Function

eCargarTemporal:
    CerrarConexionIndefa
    pb6.visible = False
    Label2(97).visible = False
    Screen.MousePointer = vbDefault
    MuestraError Err.Description, "Cargar Temporal Diferencias", Err.Description
End Function



Private Sub CmdAceptarCompCCC_Click()
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
Dim SQL As String
Dim Sql3 As String

Dim SqlZonas As String
    
Dim cadSelect1 As String
Dim cadFormula1 As String
Dim Cadena As String
    
Dim CadSelect0 As String
Dim SqlZonas0 As String
    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtcodigo(100).Text)
    cHasta = Trim(txtcodigo(101).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rsocios.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If

    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null") Then Exit Sub

    If CargarTemporalCCCErroneas(Tabla, cadSelect) Then
        If HayRegParaInforme("tmpinformes", "tmpinformes.codusu = " & vUsu.Codigo) Then
        
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo

            indRPT = 98
            ConSubInforme = True
            cadTitulo = "Cuentas Bancarias de Socios err�neas"
        
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
          
            'Nombre fichero .rpt a Imprimir
            cadNombreRPT = nomDocu
              
            'Nombre fichero .rpt a Imprimir
            LlamarImprimir
        End If
    End If

End Sub

Private Sub CmdAceptarRecManta_Click()
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
Dim TotalRegs As Long

    InicializarVbles

    If Not DatosOk Then Exit Sub


    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtcodigo(115).Text)
'    cHasta = "" 'Trim(txtcodigo(116).Text)
'    nDesde = ""
'    nHasta = ""
'    If Not (cDesde = "" And cHasta = "") Then
'        'Cadena para seleccion Desde y Hasta
'        '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
'        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
'            Codigo = "{rcampos.codsocio}"
'        Else
'            Codigo = "{rsocios_pozos.codsocio}"
'        End If
'        TipCod = "N"
'        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
'    End If

    If Not AnyadirAFormula(cadSelect, "{rcampos.codsocio} = " & DBSet(txtcodigo(115).Text, "N")) Then Exit Sub


    vSQL = ""
    If txtcodigo(115).Text <> "" Then vSQL = vSQL & " and rcampos.codsocio = " & DBSet(txtcodigo(115).Text, "N")


'09/09/2010 : solo socios que no tengan fecha de baja
'    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
    
    '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        '[Monica]20/04/2015: a�adimos la union a rzonas
        Tabla = "(rcampos INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio) INNER JOIN rzonas ON rcampos.codzonas = rzonas.codzonas "
        
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
        If HayRegParaInforme(Tabla, cadSelect) Then
            If ProcesoCarga(Tabla, cadSelect) Then
                
                frmPOZMantaAux.Show vbModal
                
                TotalRegs = DevuelveValor("select sum(nroimpresion) from rpozauxmanta")
                If TotalRegs <> 0 Then
                    If MsgBox("� Desea continuar con el proceso ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        ProcesoFacturacionConsumoMantaESCALONANew Tabla, cadSelect
                    End If
                End If
                cmdCancel_Click (11)
                
'                If vSQL <> "" And txtcodigo(115).Text = txtcodigo(116).Text Then
'                    Set frmMens4 = New frmMensajes
'
'                    frmMens4.OpcionMensaje = 57
'                    frmMens4.cadwhere = vSQL
'                    frmMens4.Show vbModal
'
'                    Set frmMens4 = Nothing
'                End If
            
            End If
        End If
    End If

'    Select Case vParamAplic.Cooperativa
'        Case 8, 10 ' ESCALONA Y UTXERA
'            ProcesoFacturacionConsumoMantaESCALONA Tabla, cadSelect
'
'        Case Else
'    End Select

End Sub


Private Function ProcesoCarga(cTabla As String, cWhere As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoCarga
    
    ProcesoCarga = False
    
    SQL = "delete from rpozauxmanta "
    conn.Execute SQL

    SQL = "select rcampos.codsocio, rcampos.codcampo, rcampos.codvarie, rcampos.codparti, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, "
    SQL = SQL & " round(supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) hanegadas, "
    '[Monica]20/04/2015: el preciomanta viene de rzonas
'    Sql = Sql & " round(round(supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) * " & DBSet(txtCodigo(112).Text, "N") & ",2) importe, 0  from " & cTabla
    SQL = SQL & " round(round(supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) * preciomanta,2) importe, 0  from " & cTabla
    If cWhere <> "" Then SQL = SQL & " where " & cWhere

    Sql2 = "insert into rpozauxmanta (codsocio, codcampo, codvarie, codparti, codzonas, poligono, parcela, subparce, hanegadas, importe, nroimpresion) "
    Sql2 = Sql2 & SQL
    conn.Execute Sql2
    
    ProcesoCarga = True
    Exit Function
    
eProcesoCarga:
    MuestraError Err.Number, "Proceso de Carga", Err.Description
End Function


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
    
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Tabla = "rpozos"

    cadSelect = " rpozos.hidrante = " & DBSet(txtcodigo(55).Text, "T")     ' Hidrante
    
    
    '[Monica]23/09/2011: de momento solo rectifico las facturas de quatretonda
    Select Case vParamAplic.Cooperativa
        '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
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
Dim SQL As String
Dim Sql2 As String
Dim I As Long

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    'si es listado de CARTAS/eMAIL a socios comprobar que se ha seleccionado
    'una carta para imprimir
    'ETIQUETAS
    CadParam = "|"

    'Nombre fichero .rpt a Imprimir
    nomRPT = "TurPOZEtiqContador.rpt"
    cadTitulo = "Etiquetas de Contadores" '"Etiquetas de Contadores"
    

    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 50 'Impresion de Etiquetas de contadores
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    
    conSubRPT = False
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H Seccion
    '--------------------------------------------
        
    'Parametro Linea 1
    If txtcodigo(45).Text <> "" Then
        CadParam = CadParam & "pLinea1="" " & txtcodigo(45).Text & """|"
    Else
        CadParam = CadParam & "pLinea1=""""|"
    End If
    numParam = numParam + 1
    
    'Parametro Linea 2
    If txtcodigo(46).Text <> "" Then
        CadParam = CadParam & "pLinea2="" " & txtcodigo(46).Text & """|"
    Else
        CadParam = CadParam & "pLinea2=""""|"
    End If
    numParam = numParam + 1
    
    'Parametro Linea 3
    If txtcodigo(47).Text <> "" Then
        CadParam = CadParam & "pLinea3="" " & txtcodigo(47).Text & """|"
    Else
        CadParam = CadParam & "pLinea3=""""|"
    End If
    numParam = numParam + 1
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = ""
    For I = 1 To CLng(txtcodigo(44).Text)
        SQL = SQL & "(" & vUsu.Codigo & "," & I & "),"
    Next I
    
    Sql2 = "insert into tmpinformes (codusu,codigo1) values "
    Sql2 = Sql2 & Mid(SQL, 1, Len(SQL) - 1) ' quitamos la ultima coma
    
    conn.Execute Sql2
    
    cadFormula = ""
    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu}=" & vUsu.Codigo) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{tmpinformes.codusu}=" & vUsu.Codigo) Then Exit Sub
    
    Tabla = "tmpinformes"
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    LlamarImprimir
    

End Sub

Private Sub CmdAcepTicFecRiego_Click()

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
Dim SQL As String
Dim Sql3 As String

Dim SqlZonas As String
    
Dim cadSelect1 As String
Dim cadFormula1 As String
Dim Cadena As String
    
Dim CadSelect0 As String
Dim SqlZonas0 As String
    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
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
                CadParam = CadParam & "pOrden={rpozos.hidrante}|"
                CadParam = CadParam & "pDescOrden=""Ordenado por Hidrante""|"
                CadParam = CadParam & "pOrden1={rpozos.hidrante}|"
                
            End If
            If Me.Option1(1).Value Then
                CadParam = CadParam & "pOrden={rpozos.nroorden}|"
                CadParam = CadParam & "pDescOrden=""Ordenado por Nro.Orden""|"
                CadParam = CadParam & "pOrden1={rpozos.hidrante}|"
            End If
            numParam = numParam + 3
            
            If Not AnyadirAFormula(cadFormula, "isnull({rpozos.fechabaja})") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{rpozos.fechabaja} is null") Then Exit Sub
            
            
            
            indRPT = 44 'listado de toma de lecturas de pozos
            ConSubInforme = False
            cadTitulo = "Listado de Toma de Lecturas"
        
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
            
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
            cadTitulo = "Comprobaci�n de Lecturas"
        
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = nomDocu
            
            If vParamAplic.Cooperativa = 7 Then
                If CargarTemporal(Tabla, cadSelect) Then
                    If HayRegParaInforme("tmpinformes", "tmpinformes.codusu = " & vUsu.Codigo) Then
                        CadParam = CadParam & "pUsu=" & vUsu.Codigo & "|"
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
            
            '[Monica]10/06/2013: Cambiamos, las cartas de talla no tienen que estar generadas para crearlas
            

            '======== FORMULA  ====================================
            'D/H Socio
            cDesde = Trim(txtcodigo(67).Text)
            cHasta = Trim(txtcodigo(68).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rsocios.codsocio}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
            End If
            
            vSQL = ""
            If txtcodigo(67).Text <> "" Then vSQL = vSQL & " and rsocios.codsocio >= " & DBSet(txtcodigo(67).Text, "N")
            If txtcodigo(68).Text <> "" Then vSQL = vSQL & " and rsocios.codsocio <= " & DBSet(txtcodigo(68).Text, "N")
        
        
            '[Monica]19/09/2012: se factura al propietario de los campos | 13/03/2014:se factura al socio antes al propietario
            Tabla = "rcampos INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
            Tabla = "(" & Tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
            
        
            If Not AnyadirAFormula(cadSelect, "{rsituacion.bloqueo} = 0") Then Exit Sub
            
'[Monica]25/03/2013: el campo no puede tener fecha de baja
            If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
'[Monica]25/03/2013: la situacion del campo debe de ser 1
            If Not AnyadirAFormula(cadSelect, "{rcampos.codsitua} = 1") Then Exit Sub
            
            '[Monica]31/05/2013: Comprobamos si de las zonas que vamos a facturar hay alguna con la suma de precios a 0 para ver si quieren o no continuar
            CadSelect0 = cadSelect
            CadSelect0 = Replace(Replace(CadSelect0, "{", ""), "}", "")
            SqlZonas0 = "rcampos.codzonas in (select codzonas from rzonas where if(precio1 is null,0,precio1) + if(precio2 is null,0,precio2) =  0)"
            If Not AnyadirAFormula(CadSelect0, SqlZonas0) Then Exit Sub
            Cadena = "select count(*) from " & Tabla & " where " & CadSelect0
            If TotalRegistros(Cadena) <> 0 Then
                
                Set frmMens2 = New frmMensajes
                
                frmMens2.OpcionMensaje = 49
                frmMens2.Cadena = "select distinct rcampos.codzonas, rzonas.nomzonas from (" & Tabla & ") inner join rzonas on rcampos.codzonas = rzonas.codzonas where " & CadSelect0
                frmMens2.Show vbModal
                If frmMens2.vCampos = "0" Then
                    Set frmMens2 = Nothing
                    cmdCancel_Click (0)
                    Exit Sub
                Else
                    Set frmMens2 = Nothing
                End If
            End If
            
            
            '[Monica]10/04/2013: solo cojo los campos que sean de zonas cuya suma de precios sea distinta de 0
            SqlZonas = "rcampos.codzonas in (select codzonas from rzonas where if(precio1 is null,0,precio1) + if(precio2 is null,0,precio2) <> 0)"
            If Not AnyadirAFormula(cadSelect, SqlZonas) Then Exit Sub
            
            
            If Not FacturacionTallaPreviaESCALONA(Tabla, cadSelect, txtcodigo(69).Text, Me.Pb5, "Practuracion Talla") Then Exit Sub
            
            '[Monica]11/04/2013: a�adimos los textos de la carta parametrizados
            CadParam = CadParam & "pFJunta=""" & txtcodigo(88).Text & """|"
            CadParam = CadParam & "pFInicio=""" & txtcodigo(89).Text & """|"
            CadParam = CadParam & "pFinCom=""" & txtcodigo(90).Text & """|"
            CadParam = CadParam & "pFProhib=""" & txtcodigo(91).Text & """|"
            CadParam = CadParam & "pBonif=""" & txtcodigo(92).Text & """|"
            CadParam = CadParam & "pPerVol=""" & txtcodigo(93).Text & """|"
            CadParam = CadParam & "pRecarg=""" & txtcodigo(94).Text & """|"
            numParam = numParam + 7
            
            
            ' si es un correo electronico miramos solo los que tienen mail
            If OptMail(0).Value Then
                
                cadSelect = QuitarCaracterACadena(cadSelect, "{")
                cadSelect = QuitarCaracterACadena(cadSelect, "}")
                cadSelect = QuitarCaracterACadena(cadSelect, "_1")
                
                cadSelect1 = cadSelect
                cadFormula1 = cadFormula
                
                If Not AnyadirAFormula(cadSelect, "not rsocios.maisocio is null and rsocios.maisocio<>''") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "not isnull({rsocios.maisocio}) and {rsocios.maisocio}<>''") Then Exit Sub
            
                indRPT = 86
                cadTitulo = "Carta de tallas a Socios"
            
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu
                ConSubInforme = True
            
                SQL = "select count(*) from " & Tabla & " where " & cadSelect
            
                If TotalRegistros(SQL) <> 0 Then
                    'Enviarlo por e-mail
                    IndRptReport = indRPT
                    EnviarEMailMulti cadSelect, Titulo, nomDocu, Tabla ' "rSocioCarta.rpt", Tabla  'email para socios
                Else
                    If MsgBox("No hay socios a enviar carta por email. � Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
                End If
            
                If Not AnyadirAFormula(cadSelect1, "rsocios.maisocio is null or rsocios.maisocio=''") Then Exit Sub
                If Not AnyadirAFormula(cadFormula1, "isnull({rsocios.maisocio}) or {rsocios.maisocio}=''") Then Exit Sub
            
                SQL = "select count(*) from " & Tabla & " where " & cadSelect1
                
                If TotalRegistros(SQL) <> 0 Then
                    cadFormula = cadFormula1
                    LlamarImprimir
                Else
                    MsgBox "No hay Socios para imprimir cartas.", vbExclamation
                End If
            
            Else
            
                If HayRegParaInforme(Tabla, cadSelect) Then
                    indRPT = 86
                    ConSubInforme = False
                    cadTitulo = "Carta de tallas a Socios"
                
                    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                    
                    cadNombreRPT = nomDocu
                        
                    ConSubInforme = True
                    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                    LlamarImprimir
                    
                End If
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
            
            
                '[Monica]19/09/2012: se factura al propietario de los campos | 13/03/2014: se factura al socio antes al propietario
                Tabla = "rcampos INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
                Tabla = "(" & Tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
                
            
                If Not AnyadirAFormula(cadSelect, "{rsituacion.bloqueo} = 0") Then Exit Sub
                
'[Monica]25/03/2013: el campo no puede tener fecha de baja
                If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
'[Monica]25/03/2013: la situacion del campo debe de ser 1
                If Not AnyadirAFormula(cadSelect, "{rcampos.codsitua} = 1") Then Exit Sub
                
                '[Monica]31/05/2013: Comprobamos si de las zonas que vamos a facturar hay alguna con la suma de precios a 0 para ver si quieren o no continuar
                
                CadSelect0 = cadSelect
                CadSelect0 = Replace(Replace(CadSelect0, "{", ""), "}", "")
                SqlZonas0 = "rcampos.codzonas in (select codzonas from rzonas where if(precio1 is null,0,precio1) + if(precio2 is null,0,precio2) =  0)"
                If Not AnyadirAFormula(CadSelect0, SqlZonas0) Then Exit Sub
                Cadena = "select count(*) from " & Tabla & " where " & CadSelect0
                If TotalRegistros(Cadena) <> 0 Then
                    
                    Set frmMens2 = New frmMensajes
                    
                    frmMens2.OpcionMensaje = 49
                    frmMens2.Cadena = "select distinct rcampos.codzonas, rzonas.nomzonas from (" & Tabla & ") inner join rzonas on rcampos.codzonas = rzonas.codzonas where " & CadSelect0
                    frmMens2.Show vbModal
                    If frmMens2.vCampos = "0" Then
                        Set frmMens2 = Nothing
                        cmdCancel_Click (0)
                        Exit Sub
                    Else
                        Set frmMens2 = Nothing
                    End If
                End If
                
                
                '[Monica]10/04/2013: solo cojo los campos que sean de zonas cuya suma de precios sea distinta de 0
                SqlZonas = "rcampos.codzonas in (select codzonas from rzonas where if(precio1 is null,0,precio1) + if(precio2 is null,0,precio2) <> 0)"
                If Not AnyadirAFormula(cadSelect, SqlZonas) Then Exit Sub
                
                
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
                
                '[Monica]19/09/2012: se actualiza la factura del propietario de los campos | 13/03/2014: el calculo de talla es para el socio no para el propietario
                Tabla = "rcampos INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
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

Private Function CargarTemporal(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql2 As String

    On Error GoTo eCargarTemporal
    
    CargarTemporal = False
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select " & vUsu.Codigo & ", codpozo, sum(consumo), sum(nroacciones) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2 "
    SQL = SQL & " order by 1, 2"
    
    Sql2 = "insert into tmpinformes (codusu, codigo1, importe1, importe2) "
    Sql2 = Sql2 & SQL
    conn.Execute Sql2
    
    CargarTemporal = True
    Exit Function

eCargarTemporal:
    MuestraError Err.Description, "Cargar Temporal", Err.Description
End Function


Private Function CargarTemporalRecibosPdtes(cTabla As String, cWhere As String, cwhere2 As String, ctabla1 As String, cwhere1 As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql2 As String
Dim SqlInsert As String


    On Error GoTo eCargarTemporal
    
    CargarTemporalRecibosPdtes = False
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SqlInsert = "insert into tmpinformes (codusu, campo1, codigo1, nombre1,importe1, nombre2, importe2, importe3, importe4, nombre3, importeb2, fecha1, importe5) "
    
    SQL = "select " & vUsu.Codigo & ",1, mid(cc.codmacta,5,6) codsocio,ss.nomsocio, cam.codzonas, zz.nomzonas, cam.codcampo, cam.poligono, cam.parcela, rr.codtipom, rr.numfactu, rr.fecfactu, sum(round((coalesce(ll.precio1,0) + coalesce(ll.precio2,0)) * ll.hanegada,2)) importe"
    SQL = SQL & " from " & cTabla
    SQL = SQL & " " & cWhere
    SQL = SQL & " group by 1,2,3,4,5,6,7,8,9,10,11,12 "
    SQL = SQL & " union "
    SQL = SQL & " select " & vUsu.Codigo & ",1, mid(cc.codmacta,5,6) codsocio,ss.nomsocio, cam.codzonas, zz.nomzonas, cam.codcampo, cam.poligono, cam.parcela, rr.codtipom, rr.numfactu, rr.fecfactu, sum(round((coalesce(ll.precio1,0) + coalesce(ll.precio2,0)) * ll.hanegada,2)) importe"
    SQL = SQL & " from " & cTabla
    SQL = SQL & " " & cwhere2
    SQL = SQL & " group by 1,2,3,4,5,6,7,8,9,10,11,12 "
    
    
    conn.Execute SqlInsert & SQL
    
    
    SqlInsert = "insert into tmpinformes (codusu, campo1, codigo1, nombre1,importeb1, nombre3, importeb2, fecha1, importe5) "
    
    SQL = "select " & vUsu.Codigo & ",2, mid(cc.codmacta,5,6) codsocio, ss.nomsocio, mid(rr.hidrante,1,2) seccion, rr.codtipom, rr.numfactu, rr.fecfactu, sum(rr.totalfact)"
    SQL = SQL & " from " & ctabla1
    SQL = SQL & " " & cwhere1
    SQL = SQL & " group by 1,2,3,4,5,6,7,8 "
    
    conn.Execute SqlInsert & SQL

    CargarTemporalRecibosPdtes = True
    Exit Function

eCargarTemporal:
    MuestraError Err.Description, "Cargar Temporal", Err.Description
End Function


Private Function CargarTemporalRecibosConsumoPdtes(ctabla1 As String, cwhere1 As String, NConta As Integer) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql2 As String
Dim SqlInsert As String


    On Error GoTo eCargarTemporal
    
    CargarTemporalRecibosConsumoPdtes = False
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SqlInsert = "insert into tmpinformes (codusu, campo1, codigo1, nombre1,importeb1, nombre3, importeb2, fecha1, importe5) "
    
    SQL = "select " & vUsu.Codigo & ",2, mid(cc.codmacta,5,6) codsocio, ss.nomsocio, rr.hidrante, rr.codtipom, rr.numfactu, rr.fecfactu, sum(rr.totalfact)"
    SQL = SQL & " from " & ctabla1
    SQL = SQL & " " & cwhere1
    SQL = SQL & " group by 1,2,3,4,5,6,7,8 "
    
    conn.Execute SqlInsert & SQL

    '[Monica]13/01/2015: cargamos el nro de reclamaciones que han hecho
    SQL = "update tmpinformes tt, usuarios.stipom ss "
    SQL = SQL & " set tt.importe1 = (select count(*) from conta" & NConta & ".shcocob aa where tt.importeb2 = aa.codfaccl "
    SQL = SQL & " and tt.fecha1 = aa.fecfaccl and ss.letraser = aa.numserie) "
    SQL = SQL & " where tt.codusu = " & vUsu.Codigo
    SQL = SQL & " and tt.nombre3 = ss.codtipom "
    
    conn.Execute SQL



    CargarTemporalRecibosConsumoPdtes = True
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
    
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Tabla = "rrecibpozos"
    End If
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
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
        Case 5
            Tipos = "{rrecibpozos.codtipom} = 'RMT'"
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
    
    '[Monica]26/08/2011: a�adido el nro de factura
    'D/H Nro Factura
    cDesde = Trim(txtcodigo(49).Text)
    cHasta = Trim(txtcodigo(50).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFactura= """) Then Exit Sub
    End If
    
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Tabla = Tabla & " INNER JOIN rsocios ON rrecibpozos.codsocio = rsocios.codsocio "
    
        '[Monica]23/11/2012: si es escalona o utxera seleccionamos que tipo de socio
        If Option1(5).Value Then    ' solo contado
            If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
        End If
        If Option1(6).Value Then    ' solo efecto
            If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba<>""8888888888""") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}<>""8888888888""") Then Exit Sub
        End If
    End If
    
    If HayRegistros(Tabla, cadSelect) Then
        indRPT = 48
        ConSubInforme = False
        cadTitulo = "Facturas por Hidrante"
        
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
          
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
          
        '[Monica]26/08/2011: nuevo report que equivale al resumen de facturas de la facturacion
        If Check2.Value = 1 Then
            If Option2.Value Then
                cadTitulo = "Resumen Facturaci�n por Socio"
                cadNombreRPT = Replace(cadNombreRPT, "RecibHidrante", "ResumFactSocio") ' agrupado por socio
            Else
                cadTitulo = "Resumen Facturaci�n"
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
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
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
        '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
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
            '[Monica]07/03/2014: nuevo campo de si se cobra la cuota
            cadSelect = cadSelect & " and {rpozos.cobrarcuota} = 1 "
        
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


    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtcodigo(23).Text)
    cHasta = Trim(txtcodigo(24).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            Codigo = "{rsocios.codsocio}"
        Else
            Codigo = "{rsocios_pozos.codsocio}"
        End If
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If


'10/06/2011 : facturamos unicamente los hidrantes que no tienen fecha de baja
''09/09/2010 : solo socios que no tengan fecha de baja
'    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null") Then Exit Sub

    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        If AnyadirAFormula(cadSelect, "(rsocios.fechabaja is null or rsocios.fechabaja = '')") = False Then Exit Sub
    
    Else

        If Not AnyadirAFormula(cadSelect, "{rpozos.fechabaja} is null") Then Exit Sub
    End If
    
    '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Tabla = "rsocios inner join rsituacion On rsituacion.bloqueo = 0"
        
        If HayRegParaInforme(Tabla, cadSelect) Then
        
            If txtcodigo(23).Text <> txtcodigo(24).Text Or txtcodigo(23).Text = "" Or txtcodigo(24).Text = "" Then
                Set frmMen = New frmMensajes
                frmMen.cadwhere = cadSelect
                frmMen.OpcionMensaje = 9 'Socios
                frmMen.Show vbModal
                Set frmMen = Nothing
                If cadSelect = "" Then Exit Sub
            End If
        End If
        
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


    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtcodigo(6).Text)
    cHasta = Trim(txtcodigo(7).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
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
    
    '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
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
            frmMens.cadwhere = vSQL
            frmMens.Show vbModal
        
            Set frmMens = Nothing
        End If
    End If

    Select Case vParamAplic.Cooperativa
        '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
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
    
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Tabla = "rrecibpozos"
    End If
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
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
    
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Tabla = Tabla & " INNER JOIN rsocios ON rrecibpozos.codsocio = rsocios.codsocio "
    
    
        '[Monica]23/11/2012: si es escalona o utxera seleccionamos que tipo de socio
        If Option1(2).Value Then    ' solo contado
            If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
        End If
        If Option1(3).Value Then    ' solo efecto
            If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba<>""8888888888""") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}<>""8888888888""") Then Exit Sub
        End If
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
                cadTitulo = "Reimpresi�n de Recibos Consumo"
            Case "RMP"
                indRPT = 47
                ConSubInforme = True
                cadTitulo = "Reimpresi�n de Recibos Mantenimiento"
            Case "RVP"
                indRPT = 47
                ConSubInforme = False
                cadTitulo = "Reimpresi�n de Recibos Contadores"
            Case "TAL"
                indRPT = 47
                ConSubInforme = True
                cadTitulo = "Reimpresi�n de Recibos Talla"
            Case "RMT"
                indRPT = 47
                ConSubInforme = True
                cadTitulo = "Reimpresi�n de Recibos Consumo Manta"
            
        End Select
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
          
        If CodTipom = "TAL" Then nomDocu = Replace(nomDocu, "Mto.", "Tal.")
        If CodTipom = "RVP" Then nomDocu = Replace(nomDocu, "Mto.", "Cont.")
        If CodTipom = "RMT" Then nomDocu = Replace(nomDocu, "Mto.", "Manta.")
  
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
          
        'Nombre fichero .rpt a Imprimir
        
        
        If vParamAplic.Cooperativa = 10 Then
            If CargarTemporalFrasPozos(Tabla, cadSelect) Then
            
                Set frmMens = New frmMensajes
                
                frmMens.OpcionMensaje = 61
                frmMens.Show vbModal
                
                Set frmMens = Nothing
        
                If HayRegistros(Tabla, cadSelect & " and rrecibpozos.imprimir=" & DBSet(vUsu.PC, "T")) Then
                    cadFormula = cadFormula & " and {rrecibpozos.imprimir} = """ & vUsu.PC & """"
                
                    LlamarImprimir
                End If
            
            End If
        Else
            LlamarImprimir
        End If
        
        If frmVisReport.EstaImpreso Then
'            ActualizarRegistros "rfactsoc", cadSelect
        End If
    End If

End Sub

Private Function CargarTemporalFrasPozos(cTabla As String, cSelect As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim cadInsert As String
Dim CadValues As String
Dim numserie As String

    On Error GoTo eCargarTemporalFrasPozos

    CargarTemporalFrasPozos = False

    SQL = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL

    ' desmarcamos todas las facturas que vamos a imprimir
    SQL = "update rrecibpozos, rsocios set imprimir = null "
    SQL = SQL & " where rrecibpozos.codsocio = rsocios.codsocio "
    If cSelect <> "" Then SQL = SQL & " and " & cSelect
    
    conn.Execute SQL
    

    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If

    SQL = "select rrecibpozos.codtipom,rrecibpozos.numfactu,rrecibpozos.fecfactu,rrecibpozos.codsocio,rsocios.nomsocio, rrecibpozos.totalfact "
    SQL = SQL & " from rrecibpozos inner join rsocios on rrecibpozos.codsocio = rsocios.codsocio "
    If cSelect <> "" Then SQL = SQL & " where " & cSelect
    
    cadInsert = "insert into tmpinformes (codusu,nombre1,importe1,fecha1,codigo1,nombre2,importe2,campo1) VALUES "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    numserie = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(Mid(Me.Combo1(0).Text, 1, 3), "T"))
    
    Label4(52).Caption = ""
    
    While Not Rs.EOF
    
        Label4(52).Caption = "Comprobando factura: " & Format(DBLet(Rs!numfactu), "0000000")
        DoEvents
    
        CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs!CodTipom, "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & ","
        CadValues = CadValues & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!TotalFact, "N") & ","
    
        SQL = "select sum(coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) from scobro where numserie = " & DBSet(numserie, "T")
        SQL = SQL & " and codfaccl = " & DBSet(Rs!numfactu, "N")
        SQL = SQL & " and fecfaccl = " & DBSet(Rs!fecfactu, "F")

        Set Rs2 = New ADODB.Recordset
        Rs2.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            If DBLet(Rs2.Fields(0).Value, "N") = 0 Then
                CadValues = CadValues & "1),"
            Else
                CadValues = CadValues & "0),"
            End If
        Else
            CadValues = CadValues & "0),"
        End If
    
        Set Rs2 = Nothing
    
        Rs.MoveNext
    Wend
    
    Label4(52).Caption = ""
    
    Set Rs = Nothing
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute cadInsert & CadValues
    End If
    
    CargarTemporalFrasPozos = True
    Exit Function
    
eCargarTemporalFrasPozos:
    MuestraError Err.Number, "Cargar Temporal Facturas de Pozos", Err.Description
End Function


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub CmdCancelRectif_Click()
    Unload Me
End Sub

Private Sub cmdCancelReimp_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim codzonas As Integer

    On Error GoTo eErrores

    conn.BeginTrans


    SQL = "select * from rcampos order by codcampo"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        codzonas = -1
    
        If DBLet(Rs!codzonas) = 0 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 1 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 2 Then
        codzonas = 54
        ElseIf DBLet(Rs!codzonas) = 3 Then
        codzonas = 10
        ElseIf DBLet(Rs!codzonas) = 4 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 5 Then
        codzonas = 7
        ElseIf DBLet(Rs!codzonas) = 6 Then
        codzonas = 7
        ElseIf DBLet(Rs!codzonas) = 7 Then
         codzonas = 7
        ElseIf DBLet(Rs!codzonas) = 8 Then
        codzonas = 51
        ElseIf DBLet(Rs!codzonas) = 9 Then
        codzonas = 9
        ElseIf DBLet(Rs!codzonas) = 10 Then
        codzonas = 10
        ElseIf DBLet(Rs!codzonas) = 11 Then
        codzonas = 10
        ElseIf DBLet(Rs!codzonas) = 12 Then
        codzonas = 21
        ElseIf DBLet(Rs!codzonas) = 13 Then
        codzonas = 7
        ElseIf DBLet(Rs!codzonas) = 14 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 15 Then
        codzonas = 15
        ElseIf DBLet(Rs!codzonas) = 16 Then
        codzonas = 15
        ElseIf DBLet(Rs!codzonas) = 17 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 18 Then
        codzonas = 19
        ElseIf DBLet(Rs!codzonas) = 19 Then
        codzonas = 19
        ElseIf DBLet(Rs!codzonas) = 20 Then
        codzonas = 19
        ElseIf DBLet(Rs!codzonas) = 21 Then
        codzonas = 21
        ElseIf DBLet(Rs!codzonas) = 22 Then
        codzonas = 22
        ElseIf DBLet(Rs!codzonas) = 23 Then
        codzonas = 22
        ElseIf DBLet(Rs!codzonas) = 24 Then
        codzonas = 25
        ElseIf DBLet(Rs!codzonas) = 25 Then
        codzonas = 25
        ElseIf DBLet(Rs!codzonas) = 26 Then
        codzonas = 25
        ElseIf DBLet(Rs!codzonas) = 27 Then
        codzonas = 28
        ElseIf DBLet(Rs!codzonas) = 28 Then
        codzonas = 28
        ElseIf DBLet(Rs!codzonas) = 29 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 30 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 31 Then
        codzonas = 36
        ElseIf DBLet(Rs!codzonas) = 32 Then
        codzonas = 34
        ElseIf DBLet(Rs!codzonas) = 33 Then
        codzonas = 34
        ElseIf DBLet(Rs!codzonas) = 34 Then
        codzonas = 34
        ElseIf DBLet(Rs!codzonas) = 35 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 36 Then
        codzonas = 36
        ElseIf DBLet(Rs!codzonas) = 37 Then
        codzonas = 37
        ElseIf DBLet(Rs!codzonas) = 38 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 39 Then
        codzonas = 39
        ElseIf DBLet(Rs!codzonas) = 40 Then
        codzonas = 41
        ElseIf DBLet(Rs!codzonas) = 41 Then
        codzonas = 41
        ElseIf DBLet(Rs!codzonas) = 42 Then
        codzonas = 41
        ElseIf DBLet(Rs!codzonas) = 43 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 44 Then
        codzonas = 44
        ElseIf DBLet(Rs!codzonas) = 45 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 46 Then
        codzonas = 46
        ElseIf DBLet(Rs!codzonas) = 47 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 48 Then
        codzonas = 49
        ElseIf DBLet(Rs!codzonas) = 49 Then
        codzonas = 49
        ElseIf DBLet(Rs!codzonas) = 50 Then
        codzonas = 49
        ElseIf DBLet(Rs!codzonas) = 51 Then
        codzonas = 51
        ElseIf DBLet(Rs!codzonas) = 52 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 53 Then
        codzonas = 53
        ElseIf DBLet(Rs!codzonas) = 54 Then
        codzonas = 54
        ElseIf DBLet(Rs!codzonas) = 55 Then
        codzonas = 54
        ElseIf DBLet(Rs!codzonas) = 56 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 57 Then
        codzonas = 57
        ElseIf DBLet(Rs!codzonas) = 58 Then
        codzonas = 58
        ElseIf DBLet(Rs!codzonas) = 59 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 60 Then
        codzonas = 60
        ElseIf DBLet(Rs!codzonas) = 61 Then
        codzonas = 60
        ElseIf DBLet(Rs!codzonas) = 62 Then
        codzonas = 60
        ElseIf DBLet(Rs!codzonas) = 63 Then
        codzonas = 60
        ElseIf DBLet(Rs!codzonas) = 64 Then
        codzonas = 64
        ElseIf DBLet(Rs!codzonas) = 65 Then
        codzonas = 65
        ElseIf DBLet(Rs!codzonas) = 66 Then
        codzonas = 66
        ElseIf DBLet(Rs!codzonas) = 67 Then
        codzonas = 67
        ElseIf DBLet(Rs!codzonas) = 68 Then
        codzonas = 68
        ElseIf DBLet(Rs!codzonas) = 69 Then
        codzonas = 69
        ElseIf DBLet(Rs!codzonas) = 70 Then
        codzonas = 70
        ElseIf DBLet(Rs!codzonas) = 71 Then
        codzonas = 99
        ElseIf DBLet(Rs!codzonas) = 72 Then
        codzonas = 99
        End If
        
        If codzonas <> -1 Then
            Sql2 = "update rcampos set codzonas = " & DBSet(codzonas, "N") & "  where codcampo = " & DBSet(Rs!CodCampo, "N")
        
            conn.Execute Sql2
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    conn.CommitTrans
    MsgBox "Proceso realizado correctamente.", vbExclamation
    cmdCancel_Click (7)
    Exit Sub
    
eErrores:
    MuestraError Err.Number, "Error"
    conn.RollbackTrans

End Sub




Private Sub Form_Activate()
Dim Nregs As Long

    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 1 ' Listado de Toma de Lectura
                PonerFoco txtcodigo(0)
            Case 2  ' Listado de comprobacion de lecturas
                '[Monica]01/07/2013: introducimos los valores por defecto
                txtcodigo(16).Text = Format(Now, "dd/mm/yyyy")
                txtcodigo(17).Text = Format(Now, "dd/mm/yyyy")
            
                PonerFoco txtcodigo(18)
            
            Case 3 ' generacion de facturas de consumo
                '[Monica]01/07/2013: introducimos los valores por defecto
                txtcodigo(13).Text = Format(Now, "dd/mm/yyyy")
                txtcodigo(15).Text = Format(Now, "dd/mm/yyyy")
               
                PonerFoco txtcodigo(11)
            Case 4 ' generacion de facturas de mantenimiento
                PonerFoco txtcodigo(6)
            Case 5 ' generacion de facturas de contadores
                PonerFoco txtcodigo(23)
            Case 6 ' reimpresion de recibos
                '[Monica]01/07/2013: introducimos los valores por defecto
                txtcodigo(36).Text = Format(Now, "dd/mm/yyyy")
                txtcodigo(37).Text = Format(Now, "dd/mm/yyyy")
               
                
                PonerFoco txtcodigo(38)
                
                Option1(4).Value = True
                
            Case 7 ' informe de facturas por hidrante
                PonerFoco txtcodigo(40)
            
                Option1(7).Value = True
            
            Case 8 ' etiquetas contadores
                PonerFoco txtcodigo(45)
                
                txtcodigo(45).Text = "AGUA CON CUPO XXXM3/HG/MES"
                txtcodigo(46).Text = "DIA:"
                txtcodigo(47).Text = "LECTURA:"
                
                Nregs = DevuelveValor("select count(*) from rpozos")
                txtcodigo(44).Text = Format(Nregs, "###,###,##0")
                
            Case 9 ' rectificacion de facturas
                txtcodigo(54).Text = Format(Now, "dd/mm/yyyy")
                PonerFoco txtcodigo(52)
                
            Case 10 ' informe de tallas (recibos de mantenimiento de Escalona)
                PonerFoco txtcodigo(67)
                
                '[Monica]01/07/2013: introducimos los valores por defecto
                txtcodigo(69).Text = Format(Now, "dd/mm/yyyy")
                
                txtcodigo(88).Text = "29 de gener"
                txtcodigo(89).Text = "1 de mar�"
                txtcodigo(90).Text = "25 de febrer"
                txtcodigo(91).Text = "de l'1 SETEMBRE"
                txtcodigo(92).Text = "Mar� 2%"
                txtcodigo(93).Text = "Abril-Maig"
                txtcodigo(94).Text = "Juny fins Desembre 20%"
                                
                Label2(92).Caption = "Precios " & vParamAplic.NomZonaPOZ
                imgBuscar(19).ToolTipText = "Ver precios " & vParamAplic.NomZonaPOZ
                
                
            Case 11, 12 'recibos y bonificacion de talla
                PonerFoco txtcodigo(74)
                If OpcionListado = 11 Then txtcodigo(76).TabIndex = 236
                If OpcionListado = 12 Then ConexionConta

                txtcodigo(73).Text = Format(Now, "dd/mm/yyyy")
                
            Case 14
                PonerFoco txtcodigo(95)
                
            Case 15 ' listado de comprobacion de pozos
                Me.Option4(1).Value = True
                PonerFoco txtcodigo(98)
                
            Case 16 ' listado de comprobacion de cuentas bancarias de socios
                PonerFoco txtcodigo(100)
                
        
            Case 17 ' generacion de recibos a manta
                txtcodigo(113).Text = "TICKET RIEGO A MANTA"
                PonerFoco txtcodigo(115)
        
            Case 18 ' informe de recibos pendientes de cobro
                PonerFoco txtcodigo(104)
            
            Case 19 ' informe de recibos por fecha de riego
                PonerFoco txtcodigo(116)
                
                Option1(13).Value = True
            
            Case 20 ' informe de recibos de consumo pendientes de cobro
                PonerFoco txtcodigo(124)
            
                Option14.Value = True
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim I As Integer
Dim SQL As String
Dim Rs As ADODB.Recordset



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
    
    
    For H = 0 To 28
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 31 To 32
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
    Me.Frame10.visible = False
    Me.FrameAsignacionPrecios.visible = False
    Me.FrameComprobacionDatos.visible = False
    Me.FrameComprobacionCCC.visible = False
    Me.FrameReciboConsumoManta.visible = False
    Me.FrameRecPdtesCobro.visible = False
    Me.FrameInfMantaFechaRiego.visible = False
    Me.FrameRecConsPdtesCobro.visible = False
    
    
    '[Monica]07/06/2013: Zona / Bra�al
    Me.Label2(81).Caption = "Precios " & vParamAplic.NomZonaPOZ
    
    
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
            Label7.Caption = "Informe de Comprobaci�n de Lecturas"
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
            
            
            '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                txtcodigo(2).Text = Format(DevuelveValor("select hastametcub1 from rtipopozos where codpozo = 1"), "0000000")
                txtcodigo(3).Text = Format(DevuelveValor("select hastametcub2 from rtipopozos where codpozo = 1"), "0000000")
                txtcodigo(4).Text = Format(DevuelveValor("select precio1 from rtipopozos where codpozo = 1"), "###,##0.0000")
                txtcodigo(5).Text = Format(DevuelveValor("select precio2 from rtipopozos where codpozo = 1"), "###,##0.0000")
                
                '[Monica]29/01/2014: por la insercion en tesoreria
                txtcodigo(48).MaxLength = 15
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
                '[Monica]29/01/2014: limitacion del concepto al arimoney
                txtcodigo(9).MaxLength = 20 '40
                txtcodigo(8).Text = Format(DevuelveValor("select imporcuotahda from rtipopozos where codpozo = 1"), "###,##0.0000")
                Label2(6).Caption = "Euros/Hanegada"
                Check1(0).Value = 1
                Check1(1).Value = 1
            Else
                If vParamAplic.Cooperativa = 8 Then
                    txtcodigo(9).MaxLength = 20
                End If
            End If
            
        Case 5 ' Generacion de recibos de contadores
            FrameReciboContadorVisible True, H, W
            indFrame = 0
            Tabla = "rsocios_pozos"
            txtcodigo(22).Text = Format(Now, "dd/mm/yyyy")
            Me.Pb3.visible = False
            
            '[Monica]27/06/2013: solo dejamos meter 40 caracteres para utxera y escalona pq se tiene que imprimir en la scobro
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                '[Monica]29/01/2014: solo dejamos meter 3 conceptos
                txtcodigo(20).MaxLength = 30 '40
                txtcodigo(25).MaxLength = 30 '40
                txtcodigo(27).MaxLength = 29 '40
                txtcodigo(29).MaxLength = 40
                txtcodigo(31).MaxLength = 40
            End If
        
        Case 6 ' Reimpresion de recibos de pozos
            FrameReimpresionVisible True, H, W
            Tabla = "rrecibpozos"
            CargaCombo
            Combo1(0).ListIndex = 0
            
            '[Monica]11/03/2013: solo en el caso de escalona y utxera pedimos el tipo de pago
            Me.FrameTipoPago.Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            Me.FrameTipoPago.visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            
        Case 7 ' Informe de recibos por hidrante
            FrameFacturasHidranteVisible True, H, W
            Tabla = "rrecibpozos"
            CargaCombo
            Combo1(1).ListIndex = 0
        
            FrameTipoPago2.visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            FrameTipoPago2.Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            
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
        
            '[Monica]29/01/2014: longitud maxima del concepto
            txtcodigo(76).MaxLength = 15
        
        
            For I = 79 To 87
                txtcodigo(I).Text = ""
            Next I
            
            txtcodigo(72).Text = ""
            txtcodigo(66).Text = ""
            txtNombre(0).Text = ""
            txtNombre(2).Text = ""
            txtNombre(4).Text = ""
            txtNombre(8).Text = ""
        
            I = 0
            SQL = "select rpretallapoz.codzonas, rzonas.nomzonas, rpretallapoz.precio1, rpretallapoz.precio2 "
            SQL = SQL & " from rpretallapoz left join rzonas on rpretallapoz.codzonas = rzonas.codzonas "
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                If DBLet(Rs!codzonas) = 0 Then
                    txtcodigo(72).Text = DBLet(Rs!Precio1, "N")
                    txtcodigo(66).Text = DBLet(Rs!Precio2, "N")
                    PonerFormatoDecimal txtcodigo(72), 7
                    PonerFormatoDecimal txtcodigo(66), 7
                Else
                    txtcodigo(79 + I).Text = DBLet(Rs!codzonas, "N")
                    PonerFormatoEntero txtcodigo(79 + I)
                    txtcodigo(80 + I).Text = DBLet(Rs!Precio1, "N")
                    txtcodigo(81 + I).Text = DBLet(Rs!Precio2, "N")
                    txtNombre(79 + I).Text = DBLet(Rs!nomzonas)
                    PonerFormatoDecimal txtcodigo(80 + I), 7
                    PonerFormatoDecimal txtcodigo(81 + I), 7
                
                    I = I + 3
                End If
                Rs.MoveNext
            Wend
            RealizarCalculos
        
        Case 12 ' Calculo de bonificacion de recibos de talla
            FrameReciboTallaVisible True, H, W
            indFrame = 0
            Tabla = "rrecibpozos"
            Me.pb4.visible = False
            Check1(6).Value = 1
            Check1(7).Value = 1
        
            txtcodigo(78).TabIndex = 236
            txtcodigo(77).TabIndex = 237
            
        
        Case 13
            Me.Frame10.visible = True
            Me.Frame10.Height = 3469
            Me.Frame10.Width = 7335
            W = Me.Frame10.Width
            H = Me.Frame10.Height
            
        Case 14 ' asignacion de precios de talla
            FrameAsignacionPreciosVisible True, H, W
        
            indFrame = 0
            Tabla = "rzonas"
            
        Case 15 ' informe de diferencias
            FrameComprobacionDatosVisible True, H, W
            indFrame = 0
            Tabla = "rpozos"
            Me.pb6.visible = False
        
        Case 16 ' informe de comprobacion de cuentas bancarias de socios
            FrameComprobacionCCCVisible True, H, W
            indFrame = 0
            Tabla = "rsocios"
            
        Case 17
            FrameReciboConsumoMantaVisible True, H, W
            indFrame = 0
            Tabla = "rcampos"
            txtcodigo(114).Text = Format(Now, "dd/mm/yyyy")
            Me.Pb7.visible = False
            
            'Si es Escalona el concepto tiene que caber en textcsb33(40 posiciones)
            If vParamAplic.Cooperativa = 10 Then
                '[Monica]29/01/2014: limitacion del concepto al arimoney
                txtcodigo(113).MaxLength = 20 '40
                txtcodigo(112).Text = Format(DevuelveValor("select imporcuotahda from rtipopozos where codpozo = 1"), "###,##0.0000")
                Label2(109).Caption = "Euros/Hanegada"
                Check1(9).Value = 1
                Check1(10).Value = 1
            Else
                If vParamAplic.Cooperativa = 8 Then
                    txtcodigo(113).MaxLength = 20
                End If
            End If
            
        Case 18 ' informe de recibos pendientes de cobro
            FrameRecPdtesCobroVisible True, H, W
            Tabla = "rrecibpozos"
            
        Case 19 ' informe de recibos por fecha de riego
            FrameInfMantaFechaRiegoVisible True, H, W
            Tabla = "rpozticketsmanta"
    
        Case 20 ' informe de recibos de consumo pendientes de consumo
            FrameRecConsPdtesCobroVisible True, H, W
            Tabla = "rpozticketsmanta"
    
    
    
    End Select
    'Esto se consigue poniendo el cancel en el opcion k corresponda
    Me.CmdCancel(0).Cancel = True
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

Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        cadSelect = cadSelect & " and rsocios.codsocio IN (" & CadenaSeleccion & ")"
    Else 'no seleccionamos ningun cliente
        cadFormula = ""
        cadSelect = ""
    End If
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        SQL = " {rpozos.hidrante} in (" & CadenaSeleccion & ")"
        Sql2 = " {rpozos.hidrante} in [" & CadenaSeleccion & "]"
    Else
        SQL = " {rpozos.hidrante} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, SQL) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMens3_datoseleccionado(CadenaSeleccion As String)
    Continuar = (CadenaSeleccion = "1")
End Sub

Private Sub frmMens4_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        SQL = " {rcampos.codcampo} in (" & CadenaSeleccion & ")"
        Sql2 = " {rcampos.codcampo} in [" & CadenaSeleccion & "]"
    Else
        SQL = " {rcampos.codcampo} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, SQL) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub



Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmZon_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "S�lo podemos poner un porcentaje de bonificaci�n o un porcentaje de" & vbCrLf & _
                      "recargo pero no ambos a la vez. " & vbCrLf & vbCrLf
                                            
        Case 1
           ' "____________________________________________________________"
            vCadena = "S�lo podemos poner un porcentaje de bonificaci�n o un porcentaje de" & vbCrLf & _
                      "recargo pero no ambos a la vez. " & vbCrLf & vbCrLf
    
        Case 2
           ' "____________________________________________________________"
            vCadena = "Concepto que se imprime en el recibo en caso de que tenga valor." & vbCrLf & _
                      "" & vbCrLf & vbCrLf
                                            
    
    
    
    End Select
    MsgBox vCadena, vbInformation, "Descripci�n de Ayuda"

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
        Case 8, 19
            AbrirFrmZonas 8
        Case 14
            AbrirFrmZonas (95)
        Case 15
            AbrirFrmZonas (79)
        Case 16
            AbrirFrmZonas (82)
        Case 17
            AbrirFrmZonas (85)
        Case 18
            AbrirFrmZonas (96)
        Case 20, 21 ' socios
            AbrirFrmSocios (Index + 80)
            
        Case 22, 23 ' socios de generacion de facturas de consumo a manta
            AbrirFrmSocios (Index + 93)
            
        Case 25, 26 ' bra�al
            AbrirFrmZonas (Index + 83)
            
        Case 27, 28 ' socios de recibos de manta por fecha de riego
            AbrirFrmSocios (Index + 89)
            
        Case 31, 32 ' socios de recibos de consumo pendientes de cobro
            AbrirFrmSocios (Index + 93)
            
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
    
    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

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
        Case 19
            indice = 118
        Case 20
            indice = 119
        '[Monica]25/09/2014: a�adimos desde/hasta fecha de pago
        Case 21, 22
            indice = Index + 89
        Case 23, 24
            indice = Index + 99
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


'[Monica]25/09/2014: seleccionamos que tipo agrupacion hay en el report
Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 14
            Label19.Caption = "Recibos por Fecha de Riego"
        Case 15
            Label19.Caption = "Recibos por Fecha de Pago"
    End Select
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
                
            Case 79: KEYBusqueda KeyAscii, 15 ' zona1
            Case 82: KEYBusqueda KeyAscii, 16 ' zona2
            Case 85: KEYBusqueda KeyAscii, 17 ' zona3
            
            Case 95: KEYBusqueda KeyAscii, 14 ' zona desde
            Case 96: KEYBusqueda KeyAscii, 18 ' zona hasta
        
            ' informe de cuentas bancarias erroneas
            Case 100: KEYBusqueda KeyAscii, 20 ' socio desde
            Case 101: KEYBusqueda KeyAscii, 21 ' socio hasta
        
            ' generacion de recibos de consumo a manta
            Case 115: KEYBusqueda KeyAscii, 22 ' socio desde
            Case 116: KEYBusqueda KeyAscii, 23 ' socio hasta
        
            ' informe de recibos pendientes de cobro
            Case 106: KEYFecha KeyAscii, 17 'fecha
            Case 107: KEYFecha KeyAscii, 18 'fecha
            Case 104: KEYFecha KeyAscii, 23 'socio
            Case 105: KEYFecha KeyAscii, 24 'socio
            Case 108: KEYBusqueda KeyAscii, 25 ' zona desde
            Case 109: KEYBusqueda KeyAscii, 26 ' zona hasta
            
            ' informe de recibos por fecha de riego
            Case 118: KEYFecha KeyAscii, 19 'fecha
            Case 119: KEYFecha KeyAscii, 20 'fecha
            Case 116: KEYFecha KeyAscii, 27 'socio
            Case 117: KEYFecha KeyAscii, 28 'socio
            '[Monica]25/09/2014: a�adimos la fecha de pago
            Case 110: KEYFecha KeyAscii, 21 'fecha de pago
            Case 111: KEYFecha KeyAscii, 22 'fecha de pago
        
            '[Monica]30/12/2014: informe de recibos de consumo pendientes de cobro
            Case 124: KEYBusqueda KeyAscii, 31 ' socio desde
            Case 125: KEYBusqueda KeyAscii, 32 ' socio hasta
            Case 122: KEYFecha KeyAscii, 24 'fecha de factura
            Case 123: KEYFecha KeyAscii, 25 'fecha de factura
        
        
        
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
Dim Precio As Currency

    'Quitar espacios en blanco por los lados
    '[Monica]29/07/2013: excepto en el caso de cooperativa = 8 or 9 los conceptos de recibo si me ponen un blanco lo dejamos
    If (Index = 9 Or Index = 48 Or Index = 76) And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) Then
        If txtcodigo(Index).Text <> " " Then txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    Else
        txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    End If
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1, 18, 19, 126, 127 ' Nro.hidrantes
    
        Case 10, 13, 14, 15, 16, 17, 22, 36, 37, 42, 43, 54, 64, 65, 69, 73, 114, 106, 107, 118, 119, 110, 111, 122, 123 'FECHAS
            If txtcodigo(Index).Text <> "" Then
                If PonerFormatoFecha(txtcodigo(Index)) Then
                End If
            End If
            
        Case 2, 3 ' rangos de consumo
            PonerFormatoEntero txtcodigo(Index)
            
        Case 4, 5 'precios para los rangos de consumo
            PonerFormatoDecimal txtcodigo(Index), 7

        Case 6, 7, 23, 24, 34, 35, 40, 41, 56, 67, 68, 74, 75, 100, 101, 115, 116, 104, 105, 116, 117, 124, 125 'socios
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
        
            '[Monica]01/07/2013: si me dan el socio desde introducir el mismo socio hasta
            If Index = 23 And txtcodigo(Index).Text <> "" Then
                txtcodigo(24).Text = txtcodigo(23).Text
                txtNombre(24).Text = txtNombre(23).Text
            End If
       
        
        Case 8 ' euros/accion
            '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                PonerFormatoDecimal txtcodigo(Index), 7
            Else
                PonerFormatoDecimal txtcodigo(Index), 3
            End If
            
        Case 112 ' euros/accion
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                PonerFormatoDecimal txtcodigo(Index), 7
            Else
                PonerFormatoDecimal txtcodigo(Index), 3
            End If

        Case 70, 71 ' cuota amortizacion y de talla ordinaria
'            PonerFormatoDecimal txtcodigo(Index), 3
'            Precio = Round2((CCur(ImporteSinFormato(ComprobarCero(txtcodigo(70).Text))) + CCur(ImporteSinFormato(ComprobarCero(txtcodigo(71).Text)))) / 200, 4)
'            txtNombre(1).Text = Format(Precio, "##,##0.0000")

            If PonerFormatoDecimal(txtcodigo(70), 7) Then
                If PonerFormatoDecimal(txtcodigo(71), 7) Then
                    txtNombre(1).Text = CCur(ComprobarCero(txtcodigo(70).Text)) + CCur(ComprobarCero(txtcodigo(71).Text))
                    If CCur(txtNombre(1).Text) = 0 Then txtNombre(1).Text = ""
                    PonerFormatoDecimal txtNombre(1), 7
                Else
                    txtcodigo(71).Text = "0"
                End If
            Else
                txtcodigo(70).Text = "0"
            End If

        Case 120 ' precio de riego a manta
            PonerFormatoDecimal txtcodigo(Index), 7

        Case 72, 66 ' cuota amortizacion y de talla ordinaria
            PonerFormatoDecimal txtcodigo(Index), 7
            
            RealizarCalculos
            
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
            
        Case 79, 82, 85, 95, 96, 108, 109 ' zonas
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000")
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rzonas", "nomzonas", "codzonas", "N")
            
        Case 80, 81, 83, 84, 86, 87 'precios para los rangos de consumo
            PonerFormatoDecimal txtcodigo(Index), 7
            
            RealizarCalculos
    End Select
End Sub

Private Sub RealizarCalculos()
'hacemos las sumas de lo que hemos descargado
    
    ' Para las zonas en general
    If CCur(ComprobarCero(txtcodigo(72).Text)) + CCur(ComprobarCero(txtcodigo(66).Text)) <> 0 Then
        txtNombre(0).Text = CCur(ComprobarCero(txtcodigo(72).Text)) + CCur(ComprobarCero(txtcodigo(66).Text))
        PonerFormatoDecimal txtNombre(0), 7
    Else
        txtNombre(0).Text = ""
    End If
    
    If CCur(ComprobarCero(txtcodigo(80).Text)) + CCur(ComprobarCero(txtcodigo(81).Text)) <> 0 Then
        txtNombre(2).Text = CCur(ComprobarCero(txtcodigo(80).Text)) + CCur(ComprobarCero(txtcodigo(81).Text))
        PonerFormatoDecimal txtNombre(2), 7
    Else
        txtNombre(2).Text = ""
    End If
    
    If CCur(ComprobarCero(txtcodigo(83).Text)) + CCur(ComprobarCero(txtcodigo(84).Text)) <> 0 Then
        txtNombre(4).Text = CCur(ComprobarCero(txtcodigo(83).Text)) + CCur(ComprobarCero(txtcodigo(84).Text))
        PonerFormatoDecimal txtNombre(4), 7
    Else
        txtNombre(4).Text = ""
    End If
    
    If CCur(ComprobarCero(txtcodigo(86).Text)) + CCur(ComprobarCero(txtcodigo(87).Text)) <> 0 Then
        txtNombre(8).Text = CCur(ComprobarCero(txtcodigo(86).Text)) + CCur(ComprobarCero(txtcodigo(87).Text))
        PonerFormatoDecimal txtNombre(8), 7
    Else
        txtNombre(8).Text = ""
    End If

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

Private Sub AbrirFrmZonas(indice As Integer)
    indCodigo = indice

    Set frmZon = New frmManZonas
    If indice = 8 Then
        frmZon.DeConsulta = False
        frmZon.DatosADevolverBusqueda = ""
        '[Monica]07/06/2013: zonas/bra�al
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            frmZon.Caption = "Bra�als"
        End If
        frmZon.DeInformes = True
    Else
        frmZon.DeConsulta = True
        frmZon.DatosADevolverBusqueda = "0|1|"
    End If
    frmZon.Show vbModal
    Set frmZon = Nothing

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

Private Sub FrameComprobacionDatosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameComprobacionDatos.visible = visible
    If visible = True Then
        Me.FrameComprobacionDatos.Top = -90
        Me.FrameComprobacionDatos.Left = 0
        Me.FrameComprobacionDatos.Height = 5655
        Me.FrameComprobacionDatos.Width = 6945
        W = Me.FrameComprobacionDatos.Width
        H = Me.FrameComprobacionDatos.Height
    End If
End Sub


Private Sub FrameComprobacionCCCVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameComprobacionCCC.visible = visible
    If visible = True Then
        Me.FrameComprobacionCCC.Top = -90
        Me.FrameComprobacionCCC.Left = 0
        Me.FrameComprobacionCCC.Height = 3255
        Me.FrameComprobacionCCC.Width = 6945
        W = Me.FrameComprobacionCCC.Width
        H = Me.FrameComprobacionCCC.Height
    End If
End Sub





Private Sub FrameCartaTallasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameCartaTallas.visible = visible
    If visible = True Then
        Me.FrameCartaTallas.Top = -90
        Me.FrameCartaTallas.Left = 0
        Me.FrameCartaTallas.Height = 8545 '7545
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
        Me.FrameReciboTalla.Height = 5085 '5925
        Me.FrameReciboTalla.Width = 6945
        W = Me.FrameReciboTalla.Width
        H = Me.FrameReciboTalla.Height
        
        If OpcionListado = 11 Then ' generacion de recibos de cuotas
            Me.FrameCuota.visible = True
            Me.FrameCuota.Enabled = True
            Me.FrameBonif.visible = False
            Me.FrameBonif.Enabled = False
        Else
            Label12.Caption = "C�lculo Bonificaci�n Recibos Talla"
            Me.FrameCuota.visible = False
            Me.FrameCuota.Enabled = False
            Me.FrameBonif.visible = True
            Me.FrameBonif.Enabled = True
            
            Me.FrameBonif.Top = 2480
        End If
    End If
End Sub


Private Sub FrameAsignacionPreciosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameAsignacionPrecios.visible = visible
    If visible = True Then
        Me.FrameAsignacionPrecios.Top = -90
        Me.FrameAsignacionPrecios.Left = 0
        Me.FrameAsignacionPrecios.Height = 4545 '5925
        Me.FrameAsignacionPrecios.Width = 6645
        W = Me.FrameAsignacionPrecios.Width
        H = Me.FrameAsignacionPrecios.Height
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

Private Sub FrameReciboConsumoMantaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameReciboConsumoManta.visible = visible
    If visible = True Then
        Me.FrameReciboConsumoManta.Top = -90
        Me.FrameReciboConsumoManta.Left = 0
        Me.FrameReciboConsumoManta.Height = 5325
        Me.FrameReciboConsumoManta.Width = 6945
        W = Me.FrameReciboConsumoManta.Width
        H = Me.FrameReciboConsumoManta.Height
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


Private Sub FrameRecPdtesCobroVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameRecPdtesCobro.visible = visible
    If visible = True Then
        Me.FrameRecPdtesCobro.Top = -90
        Me.FrameRecPdtesCobro.Left = 0
        Me.FrameRecPdtesCobro.Height = 6960 '4230
        Me.FrameRecPdtesCobro.Width = 6675
        W = Me.FrameRecPdtesCobro.Width
        H = Me.FrameRecPdtesCobro.Height
    End If
End Sub

Private Sub FrameRecConsPdtesCobroVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameRecConsPdtesCobro.visible = visible
    If visible = True Then
        Me.FrameRecConsPdtesCobro.Top = -90
        Me.FrameRecConsPdtesCobro.Left = 0
        Me.FrameRecConsPdtesCobro.Height = 5850 '4230
        Me.FrameRecConsPdtesCobro.Width = 6675
        W = Me.FrameRecConsPdtesCobro.Width
        H = Me.FrameRecConsPdtesCobro.Height
    End If
End Sub




Private Sub FrameInfMantaFechaRiegoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameInfMantaFechaRiego.visible = visible
    If visible = True Then
        Me.FrameInfMantaFechaRiego.Top = -90
        Me.FrameInfMantaFechaRiego.Left = 0
        Me.FrameInfMantaFechaRiego.Height = 5910
        Me.FrameInfMantaFechaRiego.Width = 6675
        W = Me.FrameInfMantaFechaRiego.Width
        H = Me.FrameInfMantaFechaRiego.Height
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
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .ConSubInforme = True ' ConSubInforme
        '[Monica]11/09/2015: pasamos la contabilidad que es pq tenemos que imprimir que gastos de cobros tiene.
        If vParamAplic.Cooperativa = 10 Then
            vParamAplic.NumeroConta = DevuelveValor("Select empresa_conta from rseccion where codsecci = " & vParamAplic.Seccionhorto)
        End If
        
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
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long

Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Contadores"
                    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        'comprobamos que los tipos de iva existen en la contabilidad de horto
                
        Nregs = TotalFacturasSocios(nTabla, cadSelect)
        If Nregs <> 0 Then
                Me.Pb1.visible = True
                Me.Pb1.Max = Nregs
                Me.Pb1.Value = 0
                Me.Refresh
                Mens = "Proceso Facturaci�n Consumo: " & vbCrLf & vbCrLf
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
                        CadParam = CadParam & "pFecFac= """ & txtcodigo(14).Text & """|"
                        numParam = numParam + 1
                        CadParam = CadParam & "pTitulo= ""Resumen Facturaci�n de Contadores""|"
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
                            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
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
                            
                            CadParam = CadParam & "pPorcIva=" & vPorcIva & "|"
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
                        'N� Factura
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
                        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        'Nombre fichero .rpt a Imprimir
                        cadTitulo = "Reimpresi�n de Facturas de Contadores"
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


Private Function FacturacionConsumo(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
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
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT rpozos.codsocio,hidrante,nroorden,codparti,poligono,parcelas,hanegada,lect_ant,lect_act,fech_ant,fech_act"
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    ' ordenado por socio, hidrante
    SQL = SQL & " order by rpozos.codsocio, rpozos.hidrante "
    
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))

        numfactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                numfactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe

        NumLin = 0
    End If
    
    
    While Not Rs.EOF And b
        HayReg = True
        
        ActSocio = Rs!Codsocio
        
        If ActSocio <> AntSocio Then
        
            Sql2 = "select sum(acciones) acciones from rsocios_pozos where codsocio = " & DBSet(AntSocio, "N") 'antes act
            Acciones = DevuelveValor(Sql2)
                                                                            
            Sql2 = "select sum(lect_act - lect_ant) consumo, round(sum(datediff(fech_act, fech_ant)) / count(*),0) dias"
            Sql2 = Sql2 & " from " & cTabla
            If cWhere <> "" Then
                Sql2 = Sql2 & " WHERE " & cWhere & " and "
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
            
            '[Monica]28/10/2011: a�adido el recalculo de tramos de los contadores de la factura
            SQL = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
            
            Set RsFacturas = New ADODB.Recordset
            RsFacturas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
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
            
                SQL = "update rrecibpozos set consumo1 = " & DBSet(vConsumo1, "N") & ", consumo2 = " & DBSet(vConsumo2, "N")
                SQL = SQL & ", baseimpo = " & DBSet(TotalFac, "N") & ", totalfact = " & DBSet(TotalFac, "N")
                SQL = SQL & " where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
                SQL = SQL & " and numlinea = " & DBSet(RsFacturas!numlinea, "N")
                
                conn.Execute SQL
            
                RsFacturas.MoveNext
            Wend
            
            Set RsFacturas = Nothing
            
            
            SQL = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
            
            Set RsFacturas = New ADODB.Recordset
            RsFacturas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            While Not RsFacturas.EOF
                SQL = "update rpozos set "
                SQL = SQL & " lect_ant = lect_act "
                SQL = SQL & ", fech_ant = fech_act "
                SQL = SQL & ", consumo = 0 "
                SQL = SQL & " WHERE hidrante = " & DBSet(RsFacturas!Hidrante, "T")
                
                conn.Execute SQL
                
                RsFacturas.MoveNext
            Wend
        
            Set RsFacturas = Nothing
                
            'fin a�adido
            
            AntSocio = ActSocio
            
            If b Then b = InsertResumen(tipoMov, CStr(numfactu))
           
            If b Then b = vTipoMov.IncrementarContador(tipoMov)
            
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            NumLin = 0
        End If
            
        ConsumoHidrante = DBLet(Rs!lect_act, "N") - DBLet(Rs!lect_ant, "N")
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
        
        DiferenciaDias = DBLet(Rs!fech_act, "F") - DBLet(Rs!fech_ant, "F")
        
        'insertar en la tabla de recibos de pozos
        SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, numlinea, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, difdias) "
        SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(NumLin, "N") & "," & DBSet(ActSocio, "N") & ","
        SQL = SQL & DBSet(Rs!Hidrante, "T") & "," & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & DBSet(TotalFac, "N") & "," & DBSet(ConsumoHidrante, "N") & "," & DBSet(vParamAplic.CuotaPOZ, "N") & ","
        SQL = SQL & DBSet(Rs!lect_ant, "N") & "," & DBSet(Rs!fech_ant, "F") & ","
        SQL = SQL & DBSet(Rs!lect_act, "N") & "," & DBSet(Rs!fech_act, "F") & ","
        SQL = SQL & DBSet(ConsTra1, "N") & "," & DBSet(ImporteSinFormato(txtcodigo(4).Text), "N") & ","
        SQL = SQL & DBSet(ConsTra2, "N") & "," & DBSet(ImporteSinFormato(txtcodigo(5).Text), "N") & ","
        SQL = SQL & "'Recibo de Consumo',0,"
        SQL = SQL & DBSet(DiferenciaDias, "N") & ")"
        
        conn.Execute SQL
        
        '
        '[Monica]21/10/2011: insertamos las distintas fases(acciones) del socio en la facturacion
        '
        SQL = "insert into rrecibpozos_acc(codtipom,numfactu,fecfactu,numlinea,numfases,acciones,observac) "
        SQL = SQL & " select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
        SQL = SQL & DBSet(NumLin, "N") & ", numfases, acciones, observac from rsocios_pozos where codsocio = " & DBSet(ActSocio, "N")
        
        conn.Execute SQL
            
            
        ' actualizar en los acumulados de hidrantes
        SQL = "update rpozos set acumconsumo = acumconsumo + " & DBSet(ConsumoHidrante, "N")
        SQL = SQL & ", acumcuota = acumcuota + " & DBSet(vParamAplic.CuotaPOZ, "N")
        
'        Sql = Sql & ", lect_ant = lect_act "
'        Sql = Sql & ", fech_ant = fech_act "
'        Sql = Sql & ", consumo = 0 "
        
        
        SQL = SQL & " WHERE hidrante = " & DBSet(Rs!Hidrante, "T")
        
        conn.Execute SQL
            
            
'        If b Then b = InsertResumen(tipoMov, CStr(numfactu))
        
'        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
'        AntSocio = ActSocio
        
        Rs.MoveNext
    Wend
    
    If HayReg Then
        Sql2 = "select sum(acciones) acciones from rsocios_pozos where codsocio = " & DBSet(AntSocio, "N") 'antes act
        Acciones = DevuelveValor(Sql2)
                                                                        
        Sql2 = "select sum(lect_act - lect_ant) consumo, round(sum(datediff(fech_act, fech_ant)) / count(*),0) dias"
        Sql2 = Sql2 & " from " & cTabla
        If cWhere <> "" Then
            Sql2 = Sql2 & " WHERE " & cWhere & " and "
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
    
    
        '[Monica]28/10/2011: a�adido el recalculo de tramos de los contadores de la factura
        SQL = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
        
        Set RsFacturas = New ADODB.Recordset
        RsFacturas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
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
        
            SQL = "update rrecibpozos set consumo1 = " & DBSet(vConsumo1, "N") & ", consumo2 = " & DBSet(vConsumo2, "N")
            SQL = SQL & ", baseimpo = " & DBSet(TotalFac, "N") & ", totalfact = " & DBSet(TotalFac, "N")
            SQL = SQL & " where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
            SQL = SQL & " and numlinea = " & DBSet(RsFacturas!numlinea, "N")
            
            conn.Execute SQL
        
            RsFacturas.MoveNext
        Wend
        
        Set RsFacturas = Nothing
        
        
        SQL = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
        
        Set RsFacturas = New ADODB.Recordset
        RsFacturas.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        While Not RsFacturas.EOF
            SQL = "update rpozos set "
            SQL = SQL & " lect_ant = lect_act "
            SQL = SQL & ", fech_ant = fech_act "
            SQL = SQL & ", consumo = 0 "
            SQL = SQL & " WHERE hidrante = " & DBSet(RsFacturas!Hidrante, "T")
            
            conn.Execute SQL
            
            RsFacturas.MoveNext
        Wend
    
        Set RsFacturas = Nothing
            
        
        'fin a�adido

        b = InsertResumen(tipoMov, CStr(numfactu))
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

Private Function FacturacionConsumoQUATRETONDA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String, ConsumoRectif As Long, EsRectificativa As Boolean) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
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
        SQL = "update rpozos set consumo = " & DBSet(ConsumoRectif, "N")
        SQL = SQL & ", lect_act = " & DBSet(txtcodigo(51).Text, "N")
        SQL = SQL & ", fech_act = " & DBSet(txtcodigo(54).Text, "F")
        SQL = SQL & " where " & cWhere

        conn.Execute SQL
    End If
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT rpozos.codsocio,hidrante,nroorden,codparti,poligono,parcelas,hanegada,lect_ant,lect_act,fech_ant,fech_act,nroacciones,codpozo,consumo,calibre "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    ' ordenado por socio, hidrante
    SQL = SQL & " order by rpozos.codsocio, rpozos.hidrante "
    
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))

        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0

        numfactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                numfactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe

        NumLin = 0
    End If
    
    
    While Not Rs.EOF And b
        HayReg = True
        
        ActSocio = Rs!Codsocio
        
        If ActSocio <> AntSocio Then
            
            AntSocio = ActSocio
            
            If b Then b = InsertResumen(tipoMov, CStr(numfactu))
           
            If b Then b = vTipoMov.IncrementarContador(tipoMov)
            
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            NumLin = 0
        End If
            
        Sql2 = "select precio1, imporcuota, imporcuotahda from rtipopozos where codpozo = " & DBSet(Rs!Codpozo, "N")
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs2.EOF Then
            Precio1 = DBLet(Rs2.Fields(0).Value, "N")
            ImpCuota = DBLet(Rs2.Fields(1).Value, "N")
            CuotaHda = DBLet(Rs2.Fields(2).Value, "N")
        End If
            
        Set Rs2 = Nothing
            
        Acciones = DBLet(Rs!nroacciones, "N")
            
        ImpConsumo = Round2(DBLet(Rs!Consumo, "N") * Precio1, 2)
        ImpConsumoHda = Round2(Acciones * CuotaHda, 2)
            
        '[Monica]22/09/2011: en caso de venir de una rectificativa solo se cobra el consumo
        If EsRectificativa Then
            ImpConsumoHda = 0
            ImpCuota = 0
            Acciones = 0
        End If
            
        baseimpo = ImpConsumo + ImpCuota + ImpConsumoHda
        ImpoIva = Round2(baseimpo * vPorcIva / 100, 2)
        TotalFac = baseimpo + ImpoIva
    
        IncrementarProgresNew Pb1, 1
        
        NumLin = NumLin + 1
        
        'insertar en la tabla de recibos de pozos
        SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, numlinea, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, "
        '[Monica]28/02/2012: introducimos los nuevos campos
        SQL = SQL & "codparti, calibre, codpozo) "
        SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(NumLin, "N") & "," & DBSet(ActSocio, "N") & ","
        SQL = SQL & DBSet(Rs!Hidrante, "T") & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(vPorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        SQL = SQL & DBSet(TotalFac, "N") & "," & DBSet(Rs!Consumo, "N") & "," & DBSet(ImpCuota, "N") & ","
        SQL = SQL & DBSet(Rs!lect_ant, "N") & "," & DBSet(Rs!fech_ant, "F") & ","
        SQL = SQL & DBSet(Rs!lect_act, "N") & "," & DBSet(Rs!fech_act, "F") & ","
        SQL = SQL & DBSet(Rs!Consumo, "N") & "," & DBSet(Precio1, "N") & "," ' consumo
        SQL = SQL & DBSet(Acciones, "N") & "," & DBSet(CuotaHda, "N") & ","  ' mantenimiento
        SQL = SQL & DBSet(txtcodigo(48).Text, "T") & ",0,"
        '[Monica]28/02/2012: introducimos los nuevos campos: partida,calibre y codpozo
        SQL = SQL & DBSet(Rs!codparti, "N") & "," & DBSet(Rs!calibre, "N") & "," & DBSet(Rs!Codpozo, "N") & ")"
        
        conn.Execute SQL
            
        ' actualizar en los acumulados de hidrantes
        SQL = "update rpozos set acumconsumo = acumconsumo + " & DBSet(Rs!Consumo, "N")
        SQL = SQL & ", acumcuota = acumcuota + " & DBSet(ImpCuota, "N")
        
        SQL = SQL & ", lect_ant = lect_act "
        SQL = SQL & ", fech_ant = fech_act "
        SQL = SQL & ", lect_act = null "
        SQL = SQL & ", fech_act = null "
        SQL = SQL & ", consumo = 0 "
        
        SQL = SQL & " WHERE hidrante = " & DBSet(Rs!Hidrante, "T")
        
        conn.Execute SQL
        
        Rs.MoveNext
    Wend
    
    If HayReg Then b = InsertResumen(tipoMov, CStr(numfactu))
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



Private Function TotalFacturasSocios(cTabla As String, cWhere As String) As Long
Dim SQL As String

    TotalFacturasSocios = 0
    
    SQL = "SELECT  count(distinct rpozos.codsocio) "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If

    TotalFacturasSocios = TotalRegistros(SQL)

End Function

Private Function TotalFacturasHidrante(cTabla As String, cWhere As String) As Long
Dim SQL As String

    TotalFacturasHidrante = 0
    
    SQL = "SELECT  count(distinct rpozos.hidrante) "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If

    TotalFacturasHidrante = TotalRegistros(SQL)

End Function



Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim SQL As String
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
            '[Monica]29/05/2013: Solo para escalona y utxera obligamos a escribir el concepto o poner un blanco.
            If b Then
                If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                    If Len(txtcodigo(48).Text) = 0 Then
                        MsgBox "Debe introducir un valor en el concepto.", vbExclamation
                        PonerFoco txtcodigo(48)
                        b = False
                    End If
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
                    MsgBox "Debe introducir un valor en Euros/Acci�n.", vbExclamation
                    PonerFoco txtcodigo(8)
                    b = False
                End If
            End If
            If b Then
                If (txtcodigo(9).Text = "" And vParamAplic.Cooperativa <> 8 And vParamAplic.Cooperativa <> 10) Or (Len(txtcodigo(9).Text) = 0 And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)) Then
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
                    MsgBox "Si introduce un Importe para el Art�culo 1, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtcodigo(25)
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(28).Text <> "" And txtcodigo(27).Text = "" Then
                    MsgBox "Si introduce un Importe para el Art�culo 2, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtcodigo(27)
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(30).Text <> "" And txtcodigo(29).Text = "" Then
                    MsgBox "Si introduce un Importe para el Art�culo 3, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtcodigo(29)
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(32).Text <> "" And txtcodigo(31).Text = "" Then
                    MsgBox "Si introduce un Importe para el Art�culo 4, debe introducir un Concepto correspondiente.", vbExclamation
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
                MsgBox "El n�mero de etiquetas debe ser superior a 0. Revise."
                PonerFoco txtcodigo(44)
                b = False
            End If
        
            If b Then
                If Trim(txtcodigo(45).Text) = "" And Trim(txtcodigo(46).Text) = "" And Trim(txtcodigo(47).Text) = "" Then
                    MsgBox "Debe haber alg�n valor en alguna de las L�neas. Revise."
                    PonerFoco txtcodigo(45)
                    b = False
                End If
            End If
            
        Case 9 ' Rectificacion de Lecturas
            If txtcodigo(52).Text = "" Then
                MsgBox "Debe introducir un N� de Factura. Revise", vbExclamation
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
                SQL = "select count(*) from rrecibpozos where codtipom = " & DBSet(Mid(Combo1(2).Text, 1, 3), "T")
                SQL = SQL & " and numfactu = " & DBSet(txtcodigo(52).Text, "N")
                SQL = SQL & " and codsocio = " & DBSet(txtcodigo(56).Text, "N")
                SQL = SQL & " and hidrante = " & DBSet(txtcodigo(55).Text, "T")
                If TotalRegistros(SQL) = 0 Then
                    MsgBox "No existe ninguna factura con estos datos para rectificar. Revise.", vbExclamation
                    PonerFoco txtcodigo(52)
                    b = False
                Else
                    ' miramos si es la ultima factura de ese hidrante
                    ' en este caso no debemos hacer la rectificativa porque dejariamos el hidrante con las
                    ' lecturas incorrectas
                    SQL = "select max(fecfactu) from rrecibpozos where codtipom = " & DBSet(Mid(Combo1(2).Text, 1, 3), "T")
                    SQL = SQL & " and hidrante = " & DBSet(txtcodigo(55).Text, "T")
                    FecUlt = DevuelveValor(SQL)
                    
                    SQL = "select fecfactu from rrecibpozos where codtipom = " & DBSet(Mid(Combo1(2).Text, 1, 3), "T")
                    SQL = SQL & " and numfactu= " & DBSet(txtcodigo(52).Text, "N")
                    SQL = SQL & " and hidrante = " & DBSet(txtcodigo(55).Text, "T")
                    FecFac = DevuelveValor(SQL)
                    
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
            
            '[Monica]29/05/2013: Solo para escalona y utxera obligamos a escribir el concepto o poner un blanco.
            If b Then
                '[Monica]13/03/2014: a�adimos la condicion de opcionlistado = 11 pq sino pedia un concepto
                '                    en la bonificacion no pedimos concepto
                If (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) And OpcionListado = 11 Then
                    If Len(txtcodigo(76).Text) = 0 Then
                        MsgBox "Debe introducir un valor en el concepto.", vbExclamation
                        PonerFoco txtcodigo(76)
                        b = False
                    End If
                End If
            End If
            
            If OpcionListado = 11 Then
'[Monica]10/04/2013: los precios los saco de zonas
'                If CCur(ComprobarCero(txtNombre(0).Text)) = 0 Then
'                    MsgBox "Debe introducir un valor en cuotas para facturar. Revise.", vbExclamation
'                    PonerFoco txtcodigo(72)
'                    b = False
'                End If
'[Monica]10/04/2013: quito la comprobacion de zonas
'                'Comprobamos que existan las zonas pq no hay clave referencial
'                If b And txtcodigo(79).Text <> "" Then
'                    SQL = "select nomzonas from rzonas where codzonas = " & DBSet(txtcodigo(79).Text, "N")
'                    If DevuelveValor(SQL) = 0 Then
'                        MsgBox "No existe la zona " & txtcodigo(79).Text & ". Revise.", vbExclamation
'                        PonerFoco txtcodigo(79)
'                        b = False
'                    End If
'                End If
'                If b And txtcodigo(82).Text <> "" Then
'                    SQL = "select nomzonas from rzonas where codzonas = " & DBSet(txtcodigo(82).Text, "N")
'                    If DevuelveValor(SQL) = 0 Then
'                        MsgBox "No existe la zona " & txtcodigo(82).Text & ". Revise.", vbExclamation
'                        PonerFoco txtcodigo(82)
'                        b = False
'                    End If
'                End If
'                If b And txtcodigo(85).Text <> "" Then
'                    SQL = "select nomzonas from rzonas where codzonas = " & DBSet(txtcodigo(85).Text, "N")
'                    If DevuelveValor(SQL) = 0 Then
'                        MsgBox "No existe la zona " & txtcodigo(85).Text & ". Revise.", vbExclamation
'                        PonerFoco txtcodigo(85)
'                        b = False
'                    End If
'                End If
'                If b And txtcodigo(79).Text = "" And txtNombre(2).Text <> "" Then
'                    MsgBox "No ha introducido la zona. Revise.", vbExclamation
'                    PonerFoco txtcodigo(79)
'                    b = False
'                End If
'                If b And txtcodigo(82).Text = "" And txtNombre(4).Text <> "" Then
'                    MsgBox "No ha introducido la zona. Revise.", vbExclamation
'                    PonerFoco txtcodigo(82)
'                    b = False
'                End If
'                If b And txtcodigo(85).Text = "" And txtNombre(8).Text <> "" Then
'                    MsgBox "No ha introducido la zona. Revise.", vbExclamation
'                    PonerFoco txtcodigo(85)
'                    b = False
'                End If
            End If
            
            If OpcionListado = 12 Then
                'o metemos una bonificacion o un recargo o nada, pero no ambos a la vez
                If b Then
                    If ComprobarCero(txtcodigo(78).Text) <> 0 And ComprobarCero(txtcodigo(77).Text) <> 0 Then
                        MsgBox "No se permite introducir a la vez una Bonificacion y un Recargo. Revise.", vbExclamation
                        PonerFoco txtcodigo(78)
                        b = False
                    End If
                    If b And ComprobarCero(txtcodigo(78).Text) = 0 And ComprobarCero(txtcodigo(77).Text) = 0 Then
                        MsgBox "Debe introducir un porcentaje de Bonificacion o de Recargo. Revise.", vbExclamation
                        PonerFoco txtcodigo(78)
                        b = False
                    
                    End If
            
                End If
            End If
    
        Case 17 ' generacion de recibos de consumo a manta
            If txtcodigo(115).Text = "" Then
                MsgBox "Debe introducir un valor para el Socio. Reintroduzca.", vbExclamation
                PonerFoco txtcodigo(115)
                b = False
            End If
            
            If txtcodigo(114).Text = "" Then
                MsgBox "Debe introducir un valor para la Fecha del Ticket.", vbExclamation
                PonerFoco txtcodigo(114)
                b = False
            End If
            If b Then
                If txtcodigo(112).Text = "" Then
                    MsgBox "Debe introducir un valor en Euros/Acci�n.", vbExclamation
                    PonerFoco txtcodigo(112)
                    b = False
                End If
            End If
            If b Then
                If (txtcodigo(113).Text = "" And vParamAplic.Cooperativa <> 8 And vParamAplic.Cooperativa <> 10) Or (Len(txtcodigo(113).Text) = 0 And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)) Then
                    MsgBox "Debe introducir un valor en el concepto", vbExclamation
                    PonerFoco txtcodigo(113)
                    b = False
                End If
            End If
            
        Case 18 ' informe de recibos pendientes de cobro por bra�al
            If b Then ' socio
                If txtcodigo(104).Text <> "" And txtcodigo(105).Text <> "" Then
                    If CLng(txtcodigo(104).Text) > CLng(txtcodigo(105).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtcodigo(104)
                        b = False
                    End If
                End If
            End If
            If b Then 'fecha
                If txtcodigo(106).Text <> "" And txtcodigo(107).Text <> "" Then
                    If CDate(txtcodigo(106).Text) > CDate(txtcodigo(107).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtcodigo(106)
                        b = False
                    End If
                End If
            End If
            If b Then 'zona
                If txtcodigo(108).Text <> "" And txtcodigo(109).Text <> "" Then
                    If CLng(txtcodigo(108).Text) > CLng(txtcodigo(109).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtcodigo(108)
                        b = False
                    End If
                End If
            End If
            If b Then 'sector
                If txtcodigo(102).Text <> "" And txtcodigo(103).Text <> "" Then
                    If CLng(txtcodigo(102).Text) > CLng(txtcodigo(103).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtcodigo(102)
                        b = False
                    End If
                End If
            End If
            
        Case 20 ' informe de recibos de consumo pendientes de cobro
            If b Then ' socio
                If txtcodigo(124).Text <> "" And txtcodigo(125).Text <> "" Then
                    If CLng(txtcodigo(124).Text) > CLng(txtcodigo(125).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtcodigo(124)
                        b = False
                    End If
                End If
            End If
            If b Then 'fecha
                If txtcodigo(122).Text <> "" And txtcodigo(123).Text <> "" Then
                    If CDate(txtcodigo(122).Text) > CDate(txtcodigo(123).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtcodigo(122)
                        b = False
                    End If
                End If
            End If
            If b Then 'hidrante
                If txtcodigo(126).Text <> "" And txtcodigo(127).Text <> "" Then
                    If CLng(txtcodigo(126).Text) > CLng(txtcodigo(127).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtcodigo(126)
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
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String
    
    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Mantenimiento"
                    
    Nregs = TotalRegFacturasMto(nTabla, cadSelect)
    If Nregs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.Pb2.visible = True
        Me.Pb2.Max = Nregs
        Me.Pb2.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturaci�n Mantenimiento: " & vbCrLf & vbCrLf
        b = FacturacionMantenimiento(nTabla, cadSelect, txtcodigo(10).Text, Me.Pb2, Mens)
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtcodigo(10).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturaci�n Mantenimiento""|"
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
                'N� Factura
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
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresi�n de Facturas de Mantenimiento"
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
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Mantenimiento"
                    
    Nregs = TotalRegFacturasMtoUTXERA(nTabla, cadSelect)
    If Nregs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.Pb2.visible = True
        Me.Pb2.Max = Nregs
        Me.Pb2.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturaci�n Mantenimiento: " & vbCrLf & vbCrLf
        b = FacturacionMantenimientoUTXERA(nTabla, cadSelect, txtcodigo(10).Text, Me.Pb2, Mens)
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtcodigo(10).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturaci�n Mantenimiento""|"
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
                'N� Factura
'                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(1) & "])"
'                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
'                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
'                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                '[Monica]06/03/2013: solo lo facturado
                cadAux = "{rrecibpozos.codtipom} = 'RMP'"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub


                'Fecha de Factura
                FecFac = CDate(txtcodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresi�n de Facturas de Mantenimiento"
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
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Mantenimiento"
                    
    Nregs = TotalRegFacturasMtoUTXERA(nTabla, cadSelect)
    If Nregs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.Pb2.visible = True
        Me.Pb2.Max = Nregs
        Me.Pb2.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturaci�n Mantenimiento: " & vbCrLf & vbCrLf
        b = FacturacionMantenimientoESCALONA(nTabla, cadSelect, txtcodigo(10).Text, Me.Pb2, Mens)
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtcodigo(10).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturaci�n Mantenimiento""|"
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
                'N� Factura
'                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(1) & "])"
'                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
'                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
'                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                '[Monica]06/03/2013: solo lo facturado
                cadAux = "{rrecibpozos.codtipom} = 'RMP'"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                'Fecha de Factura
                FecFac = CDate(txtcodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresi�n de Facturas de Mantenimiento"
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
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String
Dim Cadena As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Talla"
                    
    Nregs = TotalRegFacturasTallaUTXERA(nTabla, cadSelect)
    If Nregs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.pb4.visible = True
        Me.pb4.Max = Nregs
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
                Cadena = "Bonificacion: " & CCur(ImporteSinFormato(txtcodigo(78).Text)) & "%"
            Else
                Cadena = "Recargo: " & CCur(ImporteSinFormato(txtcodigo(77).Text)) & "%"
            End If
        
            LOG.Insertar 8, vUsu, "Actualizaci�n Recibos Talla Pozos: " & vbCrLf & Cadena & vbCrLf & cadSelect
        End If
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        Mens = "Proceso Facturaci�n Talla: " & vbCrLf & vbCrLf
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
                CadParam = CadParam & "pFecFac= """ & txtcodigo(73).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturaci�n Talla""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtcodigo(73).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                ConSubInforme = False
                
                LlamarImprimir
            End If
          
'[Monica]02/09/2014: CONTADOSSSS  de momento quito la impresion de facturas de escalona pq hay contado y efectos mezclados con distinta impresion
'            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
'            If Me.Check1(6).Value Then
'                cadFormula = ""
'                cadSelect = ""
'                'N� Factura
''                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(3) & "])"
''                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
''                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
''                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
'
'                '[Monica]06/03/2013: solo lo facturado
'                cadAux = "{rrecibpozos.codtipom} = 'TAL'"
'                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
'                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
'
'                'Fecha de Factura
'                FecFac = CDate(txtcodigo(73).Text)
'                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
'                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
'                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
'                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
'
'                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
'                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
'                'Nombre fichero .rpt a Imprimir
'                cadNombreRPT = Replace(nomDocu, "Mto.", "Tal.")
'                'Nombre fichero .rpt a Imprimir
'                cadTitulo = "Reimpresi�n de Facturas de Talla"
'                ConSubInforme = True
'
'                LlamarImprimir
'
'                If frmVisReport.EstaImpreso Then
''                            ActualizarRegistrosFac "rrecibpozos", cadSelect
'                End If
'            End If
            'SALIR DE LA FACTURACION DE RECIBOS DE CONTADORES
            cmdCancel_Click (1)
        Else
            MsgBox "Error en el proceso" & vbCrLf & Mens, vbExclamation
        End If
    End If
End Sub

Private Sub ProcesoFacturacionConsumoMantaESCALONA(nTabla As String, cadSelect As String)
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


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Consumo a Manta"
                    
    Nregs = TotalRegFacturasMantaESCALONA(nTabla, cadSelect)
    If Nregs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.Pb7.visible = True
        Me.Pb7.Max = Nregs
        Me.Pb7.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturaci�n Mantenimiento: " & vbCrLf & vbCrLf
        b = FacturacionConsumoMantaESCALONA(nTabla, cadSelect, txtcodigo(114).Text, Me.Pb7, Mens)
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(9).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtcodigo(114).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturaci�n Consumo a Manta""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtcodigo(114).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                ConSubInforme = False
                
                LlamarImprimir
            End If
            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
            If Me.Check1(10).Value Then
                cadFormula = ""
                cadSelect = ""
                'N� Factura
                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(4) & "])"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                '[Monica]06/03/2013: solo lo facturado
                cadAux = "{rrecibpozos.codtipom} = 'RMT'"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                'Fecha de Factura
                FecFac = CDate(txtcodigo(114).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = Replace(nomDocu, "Mto.", "Manta.")
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresi�n Facturas Consumo a Manta"
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


Private Sub ProcesoFacturacionConsumoMantaESCALONANew(nTabla As String, cadSelect As String)
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


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Consumo a Manta"
                    
    Nregs = TotalRegFacturasMantaESCALONA(nTabla, cadSelect)
    If Nregs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.Pb7.visible = True
        Me.Pb7.Max = Nregs
        Me.Pb7.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturaci�n Mantenimiento: " & vbCrLf & vbCrLf
        b = FacturacionConsumoMantaESCALONANew(nTabla, cadSelect, txtcodigo(114).Text, Me.Pb7, Mens)
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
            If Me.Check1(10).Value Then
                cadFormula = ""
                cadSelect = ""
                'N� Factura
                cadAux = "({rpozticketsmanta.numalbar} IN [" & FacturasGeneradasPOZOS(5) & "])"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                'Fecha de Ticket
                FecFac = CDate(txtcodigo(114).Text)
                cadAux = "{rpozticketsmanta.fecalbar}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rpozticketsmanta.fecalbar}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                
                '[Monica]02/09/2014: antes escontado
                If EsSocioContadoPOZOS(txtcodigo(115).Text) Then
                    'Nombre fichero .rpt a Imprimir
                    cadNombreRPT = Replace(nomDocu, "Mto.", "TicketMantaCont.")
                Else
                    'Nombre fichero .rpt a Imprimir
                    cadNombreRPT = Replace(nomDocu, "Mto.", "TicketManta.")
                End If
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresi�n Tickets Consumo a Manta"
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
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Contabiliz As Boolean
Dim LEtra As String
Dim EstaEnTesoreria As String
Dim numasien As String

    On Error GoTo eHayFactContabilizadas

    Screen.MousePointer = vbHourglass

    SQL = "SELECT rrecibpozos.* "
    SQL = SQL & " FROM  " & Tabla

    If cSelect <> "" Then
        cSelect = QuitarCaracterACadena(cSelect, "{")
        cSelect = QuitarCaracterACadena(cSelect, "}")
        cSelect = QuitarCaracterACadena(cSelect, "_1")
        SQL = SQL & " WHERE " & cSelect
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Contabiliz = False
    While Not Rs.EOF And Not Contabiliz
    
        'Cojo la letra de serie
        LEtra = ObtenerLetraSerie2(DBLet(Rs!CodTipom))
        
        'Primero comprobaremos que esta el cobro en contabilidad
        EstaEnTesoreria = ""
        If ComprobarCobroArimoney(EstaEnTesoreria, LEtra, CLng(DBLet(Rs!numfactu)), CDate(DBLet(Rs!fecfactu))) Then
            MsgBox "La factura " & LEtra & " " & DBLet(Rs!numfactu) & " de fecha " & DBLet(Rs!fecfactu) & vbCrLf & EstaEnTesoreria & vbCrLf & vbCrLf & "Revise.", vbExclamation
            Contabiliz = True
        End If
    
        ' En Escalona no va a estar en registro de iva nunca
        If Not Contabiliz Then
            If LEtra <> "" Then
                numasien = DevuelveDesdeBDNew(cConta, "cabfact", "numasien", "numserie", LEtra, "T", , "codfaccl", CStr(Rs!numfactu), "N", "anofaccl", Year(Rs!fecfactu), "N")
                If Val(ComprobarCero(numasien)) <> 0 Then
                    
                Else
                    numasien = ""
                End If
            Else
                numasien = ""
            End If
            If numasien <> "" Then
                LEtra = "La factura esta en la contabilidad, " & DBLet(Rs!numfactu) & " de fecha " & DBLet(Rs!fecfactu)
                If numasien <> "" Then LEtra = LEtra & vbCrLf & "N� asiento: " & numasien
                
                numasien = String(50, "*") & vbCrLf
                numasien = numasien & numasien & vbCrLf & vbCrLf
                LEtra = numasien & LEtra & vbCrLf & vbCrLf & numasien
                MsgBox LEtra, vbInformation
                Contabiliz = True
            End If
        End If
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    HayFactContabilizadas = Contabiliz
    
    Screen.MousePointer = vbDefault
    Exit Function

eHayFactContabilizadas:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Hay Facturas Contabilizadas", Err.Description
End Function

Public Function TotalRegFacturasMto(cTabla As String, cWhere As String) As Long
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select rsocios_pozos.codsocio, sum(acciones) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1 having sum(acciones) <> 0"
    
    TotalRegFacturasMto = TotalRegistrosConsulta(SQL)
    
End Function


Public Function TotalRegFacturasMtoUTXERA(cTabla As String, cWhere As String) As Long
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select rpozos.codsocio, sum(hanegada) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1 having sum(hanegada) <> 0"
    
    TotalRegFacturasMtoUTXERA = TotalRegistrosConsulta(SQL)
    
End Function


Public Function TotalRegFacturasMantaESCALONA(cTabla As String, cWhere As String) As Long
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select rcampos.codsocio, sum(round(supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) hanegada FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1 having sum(round(supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) <> 0"
    
    TotalRegFacturasMantaESCALONA = TotalRegistrosConsulta(SQL)
    
End Function



Public Function TotalRegFacturasTallaUTXERA(cTabla As String, cWhere As String) As Long
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    '[Monica]19/09/2012: ahora se factura al propietario del campo no al socio | 13/03/2014: ahora se factura al socio no al propietario
    SQL = "Select rcampos.codsocio codsocio, rcampos.codzonas, sum(round(supcoope * " & DBSet(vParamAplic.Faneca, "N") & ",2)) hanegada  FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2  having sum(round(supcoope * " & DBSet(vParamAplic.Faneca, "N") & ",2)) <> 0"
    
    TotalRegFacturasTallaUTXERA = TotalRegistrosConsulta(SQL)
    
End Function




Private Function FacturacionMantenimiento(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
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
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT rsocios_pozos.codsocio, sum(acciones) acciones "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    SQL = SQL & " group by 1 having sum(acciones) <> 0 "
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " order by codsocio "
    
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And b
        HayReg = True
        
        numfactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                numfactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        
        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        Sql2 = "select sum(acciones) acciones from rsocios_pozos where codsocio = " & DBSet(Rs!Codsocio, "N")
        Acciones = DevuelveValor(Sql2)
        
        TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtcodigo(8).Text)), 2)
    
        IncrementarProgresNew Pb2, 1
        
        'insertar en la tabla de recibos de pozos
        SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        SQL = SQL & "concepto, contabilizado) "
        SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
        SQL = SQL & ValorNulo & "," & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & DBSet(txtcodigo(9).Text, "T") & ",0)"
        
        conn.Execute SQL
            
        If b Then b = InsertResumen(tipoMov, CStr(numfactu))
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        Rs.MoveNext
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


Private Function FacturacionMantenimientoUTXERA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
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
Dim cadMen As String

Dim PorcDto As Currency
Dim Descuento As Currency


    On Error GoTo eFacturacion

    FacturacionMantenimientoUTXERA = False
    
    tipoMov = "RMP"
    
    conn.BeginTrans
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT rpozos.codsocio, rpozos.hidrante, rpozos.hanegada  "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    
    SQL = SQL & " group by 1, 2 having hanegada <> 0 "
    
    ' ordenado por socio, hidrante
    SQL = SQL & " order by codsocio "
    
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And b
        HayReg = True
        
        numfactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                numfactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        
        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        Acciones = DBLet(Rs!hanegada, "N")
        
'        Brazas = (Int(Acciones) * 200) + ((Acciones - Int(Acciones)) * 1000)

'        TotalFac = Round2(Brazas * CCur(ImporteSinFormato(txtcodigo(8).Text)), 2)
        TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtcodigo(8).Text)), 2)
    
        
        '[Monica]14/05/2012: tambien a�adimos el poder poner una bonificacion o recargo (como en escalona)
        ' si hay bonificacion la calculamos
        If ComprobarCero(txtcodigo(53).Text) <> "0" Then
            PorcDto = CCur(ImporteSinFormato(txtcodigo(53).Text))
            Descuento = Round2(TotalFac * PorcDto / 100, 2)
            
            TotalFac = TotalFac + Descuento
        End If
    
    
        '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
        baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - baseimpo
    
        IncrementarProgresNew Pb2, 1
        
        'insertar en la tabla de recibos de pozos
        SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        SQL = SQL & "concepto, contabilizado, porcdto, impdto, precio, escontado) "
        SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
        SQL = SQL & DBSet(Rs!Hidrante, "T") & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        SQL = SQL & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & DBSet(txtcodigo(9).Text, "T") & ",0,"
        SQL = SQL & DBSet(PorcDto, "N") & ","
        SQL = SQL & DBSet(Descuento, "N") & ","
        SQL = SQL & DBSet(CCur(ImporteSinFormato(txtcodigo(8).Text)), "N") '& ")"
        
        '[Monica]02/09/2014: CONTADOSSSS
        If EsSocioContadoPOZOS(CStr(Rs!Codsocio)) Then
            SQL = SQL & ",1)"
        Else
            SQL = SQL & ",0)"
        End If
        
        conn.Execute SQL
            
        If b Then b = InsertResumen(tipoMov, CStr(numfactu))
        
        cadMen = ""
        If b Then b = RepartoCoopropietarios(tipoMov, CStr(numfactu), CStr(FecFac), cadMen, False)
        cadMen = "Reparto Coopropietarios: " & cadMen
        
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        Rs.MoveNext
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



Private Function FacturacionMantenimientoESCALONA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
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
Dim cadMen As String

Dim PorcDto As Currency
Dim Descuento As Currency
Dim CadValues As String


    On Error GoTo eFacturacion

    FacturacionMantenimientoESCALONA = False
    
    tipoMov = "RMP"
    
    conn.BeginTrans
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT rpozos.codsocio, sum(rpozos.hanegada) hanegada, count(*) nrohidrante  "
'    Sql = "SELECT rpozos.codsocio, round(sum(rcampos.supcoope) * 12.03, 2) hanegada, count(*) nrohidrante  "
    SQL = SQL & " FROM  " & cTabla ' & ") INNER JOIN rcampos On rpozos.codcampo = rcampos.codcampo "

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    
    SQL = SQL & " group by 1 having hanegada <> 0 "
    
    ' ordenado por socio, hidrante
    SQL = SQL & " order by codsocio "
    
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And b
        HayReg = True
        
        numfactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                numfactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        
        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        Acciones = DBLet(Rs!hanegada, "N")
        
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
        baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - baseimpo
    
        IncrementarProgresNew Pb2, 1
        
        'insertar en la tabla de recibos de pozos
        SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        SQL = SQL & "concepto, contabilizado, porcdto, impdto, precio, escontado) "
        SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
        SQL = SQL & ValorNulo & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        SQL = SQL & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & DBSet(txtcodigo(9).Text, "T") & ",0,"
        SQL = SQL & DBSet(PorcDto, "N") & ","
        SQL = SQL & DBSet(Descuento, "N") & ","
        SQL = SQL & DBSet(CCur(ImporteSinFormato(txtcodigo(8).Text)), "N") '& ")"
        
        '[Monica]02/09/2014: CONTADOSSSS
        If EsSocioContadoPOZOS(CStr(Rs!Codsocio)) Then
            SQL = SQL & ",1)"
        Else
            SQL = SQL & ",0)"
        End If
        
        conn.Execute SQL
            
            
        ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
'        Sql = "SELECT hidrante, round(rcampos.supcoope * 12.03, 2) hanegada "
        SQL = "SELECT hidrante, hanegada, nroorden "
        SQL = SQL & " FROM  " & cTabla '& ") INNER JOIN rcampos ON rpozos.codcampo = rcampos.codcampo"
'        Sql = Sql & " FROM  (" & cTabla & ") INNER JOIN rcampos ON rpozos.codcampo = rcampos.codcampo"
        If cWhere <> "" Then
            SQL = SQL & " WHERE " & cWhere
            SQL = SQL & " and rpozos.codsocio = " & DBSet(Rs!Codsocio, "N")
        Else
            SQL = SQL & " where rpozos.codsocio = " & DBSet(Rs!Codsocio, "N")
        End If
            
        Set Rs8 = New ADODB.Recordset
        Rs8.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = "insert into rrecibpozos_hid (codtipom, numfactu, fecfactu, hidrante, hanegada, nroorden) values  "
        CadValues = ""
        While Not Rs8.EOF
            CadValues = CadValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
            CadValues = CadValues & DBSet(Rs8!Hidrante, "T") & "," & DBSet(Rs8!hanegada, "N") & "," & DBSet(Rs8!nroorden, "N") & "),"
            Rs8.MoveNext
        Wend
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute SQL & CadValues
        End If
        Set Rs8 = Nothing
            
        If b Then b = InsertResumen(tipoMov, CStr(numfactu))
        
'[Monica]10/05/2012: no hay reparto de coopropietarios pq ese reparto va por hidrante, ya lo veremos
'        CadMen = ""
'        If b Then b = RepartoCoopropietarios(tipoMov, CStr(NumFactu), CStr(FecFac), CadMen, False)
'        CadMen = "Reparto Coopropietarios: " & CadMen
'
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        Rs.MoveNext
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


Private Function FacturacionTallaESCALONA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
Dim TotalFac As Currency
Dim TotalZona As Currency

Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long
Dim Brazas As Long
Dim cadMen As String

Dim PorcDto As Currency
Dim Descuento As Currency
Dim CadValues As String
Dim Precio As Currency

Dim PrecioBrz As Currency
Dim SocioAnt As Long
Dim SqlPrec As String


    On Error GoTo eFacturacion

    FacturacionTallaESCALONA = False
    
    tipoMov = "TAL"
    
    conn.BeginTrans
'[Monica]10/04/2013: quito la carga de la tabla intermedia de precios
'    b = CargarTablaPrecios
'    If Not b Then
'        conn.RollbackTrans
'        FacturacionTallaESCALONA = False
'        Exit Function
'    End If
    b = True
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    '[Monica]13/03/2014: ahora se factura al socio no al propietario
    SQL = "SELECT rcampos.codsocio codsocio, rcampos.codzonas, round(sum(rcampos.supcoope) / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    
    SQL = SQL & " group by 1, 2 having hanegada <> 0  "
    
    ' ordenado por socio, hidrante
    SQL = SQL & " order by codsocio, codzonas "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        SocioAnt = DBLet(Rs!Codsocio, "N")
        
    End If
    
    While Not Rs.EOF And b
        HayReg = True
        
        If SocioAnt <> DBLet(Rs!Codsocio, "N") Then
        
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
        
        
            baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
            ImpoIva = TotalFac - baseimpo
        
        
            'insertar en la tabla de recibos de pozos
            SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
            SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
            SQL = SQL & "concepto, contabilizado, porcdto, impdto, precio,escontado) "
            SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(SocioAnt, "N") & ","
            SQL = SQL & ValorNulo & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            SQL = SQL & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(txtcodigo(76).Text, "T") & ",0,"
            SQL = SQL & DBSet(PorcDto, "N") & ","
            SQL = SQL & DBSet(Descuento, "N") & ","
            SQL = SQL & DBSet(PrecioBrz, "N") '& ")"
            
            '[Monica]02/09/2014: CONTADOSSSS
            If EsSocioContadoPOZOS(CStr(SocioAnt)) Then
                SQL = SQL & ",1)"
            Else
                SQL = SQL & ",0)"
            End If
            
            conn.Execute SQL
            
            ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
            SQL = "SELECT rcampos.codcampo, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
            SQL = SQL & " FROM  " & cTabla
            If cWhere <> "" Then
                SQL = SQL & " WHERE " & cWhere
                '[Monica]13/03/2014: se factura al socio no al propietario
                SQL = SQL & " and rcampos.codsocio = " & DBSet(SocioAnt, "N")
            Else
                '[Monica]13/03/2014: se factura al socio no al propietario
                SQL = SQL & " where rcampos.codsocio = " & DBSet(SocioAnt, "N")
            End If
                
            Set Rs8 = New ADODB.Recordset
            Rs8.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = "insert into rrecibpozos_cam (codtipom, numfactu, fecfactu, codcampo, hanegada, precio1, precio2, codzonas, poligono, parcela, subparce) values  "
            CadValues = ""
            While Not Rs8.EOF
                Precio = DevuelvePrecio(DBLet(Rs8!codzonas, "N"))
                
                CadValues = CadValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                CadValues = CadValues & DBSet(Rs8!CodCampo, "N") & "," & DBSet(Rs8!hanegada, "N") & ","
                CadValues = CadValues & DBSet(PrecioTalla1, "N") & "," & DBSet(PrecioTalla2, "N") & "," & DBSet(ZonaTalla, "N") & ","
                CadValues = CadValues & DBSet(Rs8!poligono, "N") & "," & DBSet(Rs8!Parcela, "N") & "," & DBSet(Rs8!SubParce, "T") & "),"
                
                Rs8.MoveNext
            Wend
            
            If CadValues <> "" Then
                CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
                conn.Execute SQL & CadValues
            End If
            Set Rs8 = Nothing
                
            If b Then b = InsertResumen(tipoMov, CStr(numfactu))
            
            If b Then b = vTipoMov.IncrementarContador(tipoMov)
            
            
            baseimpo = 0
            ImpoIva = 0
            TotalFac = 0
        
            SocioAnt = DBLet(Rs!Codsocio, "N")
            
        End If
        
        Acciones = DBLet(Rs!hanegada, "N")
        
        Precio = DevuelvePrecio(Rs!codzonas)
        
        TotalFac = TotalFac + Round2(Acciones * Precio, 2)
        
        IncrementarProgresNew Pb1, 1
        
        Label2(78).Caption = "Socio: " & Format(Rs!Codsocio, "000000")
        DoEvents
        
        Rs.MoveNext
    Wend
    
    If HayReg And b Then
        numfactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                numfactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
        baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - baseimpo
    
        'insertar en la tabla de recibos de pozos
        SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        SQL = SQL & "concepto, contabilizado, porcdto, impdto, precio, escontado) "
        SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(SocioAnt, "N") & ","
        SQL = SQL & ValorNulo & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        SQL = SQL & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & DBSet(txtcodigo(76).Text, "T") & ",0,"
        SQL = SQL & DBSet(PorcDto, "N") & ","
        SQL = SQL & DBSet(Descuento, "N") & ","
        SQL = SQL & DBSet(PrecioBrz, "N") '& ")"
        
        '[Monica]02/09/2014: CONTADOSSSS
        If EsSocioContadoPOZOS(CStr(SocioAnt)) Then
            SQL = SQL & ",1)"
        Else
            SQL = SQL & ",0)"
        End If
            
        conn.Execute SQL
        
        ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
        SQL = "SELECT rcampos.codcampo, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
        SQL = SQL & " FROM  " & cTabla
        If cWhere <> "" Then
            SQL = SQL & " WHERE " & cWhere
            '[Monica]13/03/2014: se factura al socio no al propietario
            SQL = SQL & " and rcampos.codsocio = " & DBSet(SocioAnt, "N")
        Else
            '[Monica]13/03/2014: se factura al socio no al propietario
            SQL = SQL & " where rcampos.codsocio = " & DBSet(SocioAnt, "N")
        End If
            
        Set Rs8 = New ADODB.Recordset
        Rs8.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = "insert into rrecibpozos_cam (codtipom, numfactu, fecfactu, codcampo, hanegada, precio1, precio2, codzonas, poligono, parcela, subparce) values  "
        CadValues = ""
        While Not Rs8.EOF
            Precio = DevuelvePrecio(DBLet(Rs8!codzonas))
        
            CadValues = CadValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
            CadValues = CadValues & DBSet(Rs8!CodCampo, "N") & "," & DBSet(Rs8!hanegada, "N") & ","
            CadValues = CadValues & DBSet(PrecioTalla1, "N") & "," & DBSet(PrecioTalla2, "N") & "," & DBSet(ZonaTalla, "N") & ","
            CadValues = CadValues & DBSet(Rs8!poligono, "N") & "," & DBSet(Rs8!Parcela, "N") & "," & DBSet(Rs8!SubParce, "T") & "),"
            
            Rs8.MoveNext
        Wend
        
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute SQL & CadValues
        End If
        Set Rs8 = Nothing
            
        If b Then b = InsertResumen(tipoMov, CStr(numfactu))
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
    
    End If
    
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


Private Function ActualizacionTallaESCALONA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
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
Dim cadMen As String

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
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT rrecibpozos.* "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    
    ' ordenado por socio
    SQL = SQL & " order by rrecibpozos.codsocio, rrecibpozos.numfactu "
    
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And b
        HayReg = True
        
        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        TotalFac = DBLet(Rs!TotalFact, "N")
        
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
        baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - baseimpo
    
        IncrementarProgresNew Pb1, 1
        
        'modificamos la tabla de recibos de pozos
        SQL = "update rrecibpozos set baseimpo = " & DBSet(baseimpo, "N")
        SQL = SQL & ", tipoiva = " & DBSet(vParamAplic.CodIvaPOZ, "N")
        SQL = SQL & ", porc_iva = " & DBSet(PorcIva, "N")
        SQL = SQL & ", imporiva = " & DBSet(ImpoIva, "N")
        SQL = SQL & ", totalfact = " & DBSet(TotalFac, "N")
        SQL = SQL & ", porcdto = " & DBSet(PorcDto, "N")
        SQL = SQL & ", impdto = " & DBSet(Descuento, "N")
        SQL = SQL & " where codtipom = 'TAL'"
        SQL = SQL & " and numfactu = " & DBSet(Rs!numfactu, "N")
        SQL = SQL & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        SQL = SQL & " and codsocio = " & DBSet(Rs!Codsocio, "N")
        
        conn.Execute SQL
            
        ' Si el recibo est� contabilizado actualizaremos el arimoney
        LetraSerie = DevuelveValor("select letraser from usuarios.stipom where codtipom = 'TAL'")

        SQL = "update scobro set impvenci = " & DBSet(TotalFac, "N")
        SQL = SQL & " where numserie = " & DBSet(LetraSerie, "T")
        SQL = SQL & " and codfaccl = " & DBSet(Rs!numfactu, "N")
        SQL = SQL & " and fecfaccl = " & DBSet(Rs!fecfactu, "F")
        SQL = SQL & " and numorden = 1 "
            
        ConnConta.Execute SQL
        
        '[Monica]19/09/2012: al enlazar por el propietario y campos me salian todos los campos de ese propietario,
        '                    si el nro de factura, tipo ya existe no lo volvemos a insertar en el resumen
        '                    He a�adido: and totalregistros...
        If b And TotalRegistros("select count(*) from tmpinformes where codusu = " & vUsu.Codigo & " and nombre1 = 'TAL' and importe1 = " & DBSet(Rs!numfactu, "N")) = 0 Then b = InsertResumen("TAL", CStr(Rs!numfactu))
        
        Rs.MoveNext
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
Dim SQL As String
Dim RS1 As ADODB.Recordset
Dim Cad As String
    
    On Error GoTo eFacturasGeneradas

    FacturasGeneradasPOZOS = ""

    SQL = "select nombre1, importe1 from tmpinformes where codusu = " & vUsu.Codigo
    SQL = SQL & " and nombre1 = "
    Select Case Tipo
        Case 0 ' recibos de consumo de pozos
            SQL = SQL & "'RCP'"
        Case 1 ' recibos de mantenimiento de pozos
            SQL = SQL & "'RMP'"
        Case 2 ' recibos de contadores de pozos
            SQL = SQL & "'RVP'"
        Case 3
            SQL = SQL & "'TAL'"
        Case 4
            SQL = SQL & "'RMT'"
        Case 5
            SQL = SQL & "'ALV'"
    End Select
    
    Set RS1 = New ADODB.Recordset
    RS1.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Cad = ""
    While Not RS1.EOF
        Cad = Cad & DBLet(RS1.Fields(1).Value, "N") & ","
    
        RS1.MoveNext
    Wend
    Set RS1 = Nothing
    
    'si hay facturas quitamos la ultima coma
    If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
    
    FacturasGeneradasPOZOS = Cad
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
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
'    cadNombreRPT = "rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Contadores"
                    
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Nregs = TotalSocios(nTabla, cadSelect)
    Else
        Nregs = TotalRegFacturasMto(nTabla, cadSelect)
    End If
    If Nregs = 0 Then
        MsgBox "No hay registros a facturar.", vbExclamation
    Else
        Me.Pb3.visible = True
        Me.Pb3.Max = Nregs
        Me.Pb3.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturaci�n Contadores: " & vbCrLf & vbCrLf
        b = FacturacionContadores(nTabla, cadSelect, txtcodigo(22).Text, Me.Pb3, Mens)
        If b Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION de recibos de contadores
            If Me.Check1(5).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtcodigo(22).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturaci�n Contadores""|"
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
                'N� Factura
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
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                
                If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                    nomDocu = Replace(nomDocu, "Mto.", "Cont.")
                End If
                
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresi�n de Facturas de Contadores"
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

Private Function FacturacionContadores(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
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
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT rsocios.codsocio "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    SQL = SQL & " group by 1 "
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " order by codsocio "
    
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And b
        HayReg = True
        
        numfactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                numfactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        
        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        TotalFac = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(33).Text)))
    
        IncrementarProgresNew Pb3, 1
        
        'insertar en la tabla de recibos de pozos
        SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        SQL = SQL & "concepto, contabilizado, conceptomo, importemo, conceptoar1, importear1, conceptoar2, importear2, conceptoar3, "
        SQL = SQL & "importear3, conceptoar4, importear4"
        '[Monica]02/09/2014: CONTADOSSSS
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            SQL = SQL & ",escontado) "
        Else
            SQL = SQL & ") "
        End If
        SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
        SQL = SQL & ValorNulo & "," & DBSet(TotalFac, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        SQL = SQL & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & ValorNulo & ",0,"
        SQL = SQL & DBSet(txtcodigo(20).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtcodigo(21).Text))), "N", "S") & "," ' mano de obra
        SQL = SQL & DBSet(txtcodigo(25).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtcodigo(26).Text))), "N", "S") & "," ' articulo 1
        SQL = SQL & DBSet(txtcodigo(27).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtcodigo(28).Text))), "N", "S") & "," ' articulo 2
        SQL = SQL & DBSet(txtcodigo(29).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtcodigo(30).Text))), "N", "S") & "," ' articulo 3
        SQL = SQL & DBSet(txtcodigo(31).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtcodigo(32).Text))), "N", "S") '& ")" ' articulo 4
        
        '[Monica]02/09/2014: CONTADOSSSS
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            If EsSocioContadoPOZOS(CStr(Rs!Codsocio)) Then
                SQL = SQL & ",1)"
            Else
                SQL = SQL & ",0)"
            End If
        Else
            SQL = SQL & ")"
        End If
        
        
        conn.Execute SQL
            
        If b Then b = InsertResumen(tipoMov, CStr(numfactu))
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        Rs.MoveNext
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
Dim Rs As ADODB.Recordset
Dim SQL As String
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
        Combo1(0).AddItem "RMT-Consumo Manta"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 4
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
        Combo1(1).AddItem "RMT-Consumo Manta"
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
        Combo1(2).AddItem "RMT-Consumo Manta"
        Combo1(2).ItemData(Combo1(2).NewIndex) = 4
        
    End If
    
    
    
End Sub


Private Sub ProcesoFacturacionConsumoUTXERA(nTabla As String, cadSelect As String)
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


Dim Nregs As Long
Dim FecFac As Date

Dim Mens As String

Dim b As Boolean
Dim Sql2 As String

    '[Monica]29/08/2012: personalizamos la impresion de resumen de facturas pozos
    indRPT = 87 'Impresion de resumen de recibos de consumo de contadores de pozos
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    cadNombreRPT = nomDocu '"rResumFacturasPOZ.rpt"
    
    cadTitulo = "Resumen de Recibos de Contadores"
                    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        'comprobamos que los tipos de iva existen en la contabilidad de horto
                
        '[Monica]11/06/2013: semaforo para saber si hay facturas generadas
        HayFacturas = False
                
        '[Monica]11/06/2013: Mostramos si hay contadores con consumo inferior/superior al minimo/maximo
        MostrarContadoresANoFacturar nTabla, cadSelect
        If Not Continuar Then
            cmdCancel_Click (1)
            Exit Sub
        End If
                
        Nregs = TotalFacturasHidrante(nTabla, cadSelect)
        If Nregs <> 0 Then
                Me.Pb1.visible = True
                Me.Pb1.Max = Nregs
                Me.Pb1.Value = 0
                Me.Refresh
                Mens = "Proceso Facturaci�n Consumo: " & vbCrLf & vbCrLf
                b = FacturacionConsumoUTXERA(nTabla, cadSelect, txtcodigo(14).Text, Me.Pb1, Mens)
                If b Then
                    If Not HayFacturas Then
                        MsgBox "No se han generado de facturas de consumo.", vbExclamation
                        cmdCancel_Click (0)
                        Exit Sub
                    End If
                                
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                   
                    'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                    If Me.Check1(2).Value Then
                        cadFormula = ""
                        CadParam = CadParam & "pFecFac= """ & txtcodigo(14).Text & """|"
                        numParam = numParam + 1
                        CadParam = CadParam & "pTitulo= ""Resumen Facturaci�n de Contadores""|"
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
                        'N� Factura
'                        cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(0) & "])"
'                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
'                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
'                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        '[Monica]06/03/2013: solo lo facturado
                        cadAux = "{rrecibpozos.codtipom} = 'RCP'"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        'Fecha de Factura
                        FecFac = CDate(txtcodigo(14).Text)
                        cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        indRPT = 46 'Impresion de recibos de consumo de contadores de pozos
                        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        'Nombre fichero .rpt a Imprimir
                        cadTitulo = "Reimpresi�n de Facturas de Contadores"
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


Private Function FacturacionConsumoUTXERA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
Dim TotalFac As Currency


Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long

Dim cadMen As String

    On Error GoTo eFacturacion

    FacturacionConsumoUTXERA = False

    tipoMov = "RCP"

    conn.BeginTrans


    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT rpozos.codsocio,hidrante,nroorden,codparti,poligono,parcelas,hanegada,lect_ant,lect_act,fech_ant,fech_act,codpozo,consumo "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If

    ' ordenado por socio, hidrante
    SQL = SQL & " order by rpozos.codsocio, rpozos.hidrante "

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

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    baseimpo = 0
    ImpoIva = 0
    TotalFac = 0


    While Not Rs.EOF And b
        HayReg = True
            
            
        IncrementarProgresNew Pb1, 1

'If RS!CodSocio = 168 Then
'    MsgBox "168"
'End If



        '[Monica]17/05/2013: a�adida la condicion de que el consumo ha de ser superior o igual al m�nimo
        '[Monica]24/10/2011: a�adida esta condicion para que si no hay consumo se actualicen fechas
        ConsumoHidrante = DBLet(Rs!Consumo, "N")
        If DBLet(Rs!Consumo, "N") <> 0 And DBLet(Rs!Consumo, "N") >= vParamAplic.ConsumoMinPOZ And DBLet(ConsumoHidrante, "N") <= vParamAplic.ConsumoMaxPOZ Then
    
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
    
            ConsumoHidrante = DBLet(Rs!Consumo, "N") 'DBLet(RS!lect_act, "N") - DBLet(RS!lect_ant, "N")
            Consumo = ConsumoHidrante
    
            ConsTra1 = Consumo
            
            ' consumo de agua y consumo de electricidad
            
            baseimpo = Round2(ConsTra1 * CCur(ImporteSinFormato(txtcodigo(4).Text)), 2) + _
                       Round2(ConsTra1 * CCur(ImporteSinFormato(txtcodigo(5).Text)), 2)
    
    
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
            TotalFac = baseimpo + Round2(baseimpo * PorcIva / 100, 2)
    
    
            'insertar en la tabla de recibos de pozos
            SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, numlinea, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
            SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, codparti, parcelas, poligono, nroorden, escontado) "
            SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ",1,"
            SQL = SQL & DBSet(Rs!Hidrante, "T") & "," & DBSet(baseimpo, "N") & "," & vParamAplic.CodIvaPOZ & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            SQL = SQL & DBSet(TotalFac, "N") & "," & DBSet(ConsumoHidrante, "N") & "," & DBSet(0, "N") & ","
            SQL = SQL & DBSet(Rs!lect_ant, "N") & "," & DBSet(Rs!fech_ant, "F") & ","
            SQL = SQL & DBSet(Rs!lect_act, "N") & "," & DBSet(Rs!fech_act, "F") & ","
            SQL = SQL & DBSet(ConsTra1, "N") & "," & DBSet(ImporteSinFormato(txtcodigo(4).Text), "N") & ","
            SQL = SQL & DBSet(ConsTra1, "N") & "," & DBSet(ImporteSinFormato(txtcodigo(5).Text), "N") & ","
            
            '[Monica]22/10/2012: si nos han puesto un concepto guardammos el concepto
            ' antes :     Sql = Sql & "'Recibo de Consumo',0,"
            If txtcodigo(48).Text <> "" Then
                SQL = SQL & DBSet(txtcodigo(48).Text, "T") & ",0,"
            Else
                SQL = SQL & DBSet(vTipoMov.NombreMovimiento, "T") & ",0,"
            End If
            
            '[Monica]22/10/2012: guardamos tambien la partida [Monica]03/05/2013: ahora tb el poligono [Monica]22/07/2013: metemos el nro de orden
            SQL = SQL & DBSet(Rs!codparti, "N") & "," & DBSet(Rs!parcelas, "T") & "," & DBSet(Rs!poligono, "T") & "," & DBSet(Rs!nroorden, "N") '& ")"
    
            '[Monica]02/09/2014: CONTADOSSSS
            If EsSocioContadoPOZOS(CStr(Rs!Codsocio)) Then
                SQL = SQL & ",1)"
            Else
                SQL = SQL & ",0)"
            End If
    
            conn.Execute SQL
        
            If b Then b = RepartoCoopropietarios(tipoMov, CStr(numfactu), CStr(FecFac), cadMen)
            cadMen = "Reparto Coopropietarios: " & cadMen
        
            If b Then b = InsertResumen(tipoMov, CStr(numfactu))
        
            If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        End If

        If DBLet(Rs!fech_act, "F") <> "" Then
            '[Monica]11/06/2013: a�adida la condicion de que el consumo sea inferior o igual al consumo maximo de parametros
            If DBLet(ConsumoHidrante, "N") >= vParamAplic.ConsumoMinPOZ And DBLet(ConsumoHidrante, "N") <= vParamAplic.ConsumoMaxPOZ Then
            
                HayFacturas = True
            
                ' actualizar en los acumulados de hidrantes
                SQL = "update rpozos set acumconsumo = acumconsumo + " & DBSet(ConsumoHidrante, "N")
                SQL = SQL & ", lect_ant = lect_act "
                SQL = SQL & ", fech_ant = fech_act "
        '        sql = sql & ", lect_act = null "
                SQL = SQL & ", fech_act = null "
                SQL = SQL & ", consumo = 0 "
                SQL = SQL & " WHERE hidrante = " & DBSet(Rs!Hidrante, "T")
            Else
                '[Monica]17/05/2013: en el caso de que el consumo no supere el m�nimo
                '                    dejamos la lectura actual = a la que tenia la lectura anterior
                '                    la fecha anterior no se actualiza
                '                    y la fecha actual se deja a null
                SQL = "update rpozos set lect_act = lect_ant "
                SQL = SQL & ", fech_act = null "
                SQL = SQL & ", consumo = 0 "
                SQL = SQL & " WHERE hidrante = " & DBSet(Rs!Hidrante, "T")
            End If
            
            conn.Execute SQL
        End If
        
        Rs.MoveNext
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
Dim SQL As String

    SQL = "select count(*) from rpozos_cooprop where hidrante = " & DBSet(Hidrante, "T") & " and codsocio <> " & DBSet(Propietario, "N")
    
    TieneCopropietariosPOZOS = TotalRegistros(SQL) > 0

End Function

Private Function RepartoCoopropietarios(tipoMov As String, Factura As String, Fecha As String, cadErr As String, Optional SinIva As Boolean) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Dim numalbar As Long
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
Dim Numreg As Long
Dim campo As Long
Dim Porcentaje As Single
Dim numFac As Long
Dim vPorcen As String

    On Error GoTo eRepartoCoopropietarios

    RepartoCoopropietarios = False
    
    cadErr = ""
    
    b = True
    
    SQL = "select * from rrecibpozos where codtipom  = " & DBSet(tipoMov, "T")
    SQL = SQL & " and numfactu = " & DBSet(Factura, "N")
    SQL = SQL & " and fecfactu = " & DBSet(Fecha, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        
        If TieneCopropietariosPOZOS(CStr(Rs!Hidrante), CStr(Rs!Codsocio)) Then
            CodTipoMov = tipoMov
        
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(CodTipoMov) Then

                tBaseImpo = DBLet(Rs!baseimpo, "N")
                tImporIva = DBLet(Rs!ImporIva, "N")
                tTotalFact = DBLet(Rs!TotalFact, "N")

                Sql2 = "select * from rpozos_cooprop where hidrante = " & DBSet(Rs!Hidrante, "T")
                Sql2 = Sql2 & " and rpozos_cooprop.codsocio <> " & DBSet(Rs!Codsocio, "N")
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
                    
                    vBaseImpo = Round2(DBLet(Rs!baseimpo, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImporIva = Round2(DBLet(Rs!ImporIva, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vTotalFact = Round2(DBLet(Rs!TotalFact, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    
                    tBaseImpo = tBaseImpo - vBaseImpo
                    tImporIva = tImporIva - vImporIva
                    tTotalFact = tTotalFact - vTotalFact
                    
                    'insertar en la tabla de recibos de pozos
                    Sql4 = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, numlinea, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
                    Sql4 = Sql4 & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, "
                    '[Monica]29/04/2014: a�adidos campos que faltaban
                    Sql4 = Sql4 & "conceptomo,importemo,conceptoar1,importear1,conceptoar2,importear2,conceptoar3,importear3,conceptoar4,importear4,difdias,calibre,codpozo,porcdto,impdto,precio,pasaridoc,"
                    Sql4 = Sql4 & "codparti, parcelas, poligono, nroorden"
                    
                    '[Monica]02/09/2014: CONTADOSSSS
                    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                        Sql4 = Sql4 & ", escontado) "
                    Else
                        Sql4 = Sql4 & ") "
                    End If
                    
                    Sql4 = Sql4 & " values ('" & tipoMov & "'," & DBSet(numFac, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Rs2!Codsocio, "N") & ",1,"
                    Sql4 = Sql4 & DBSet(Rs!Hidrante, "T") & "," & DBSet(vBaseImpo, "N") & ","
                    
                    If SinIva Then
                        Sql4 = Sql4 & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    Else
                        Sql4 = Sql4 & vParamAplic.CodIvaPOZ & "," & DBSet(Rs!porc_iva, "N") & "," & DBSet(vImporIva, "N") & ","
                    End If
                    
                    Sql4 = Sql4 & DBSet(vTotalFact, "N") & "," & DBSet(Rs!Consumo, "N", "S") & "," & DBSet(0, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!lect_ant, "N") & "," & DBSet(Rs!fech_ant, "F") & ","
                    Sql4 = Sql4 & DBSet(Rs!lect_act, "N") & "," & DBSet(Rs!fech_act, "F") & ","
                    Sql4 = Sql4 & DBSet(Rs!Consumo1, "N") & "," & DBSet(Rs!Precio1, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!Consumo1, "N") & "," & DBSet(Rs!Precio2, "N") & ","
                    If tipoMov = "RCP" Then
                        Sql4 = Sql4 & DBSet(Rs!Concepto & " " & Format(DBLet(Rs2!Porcentaje, "N"), "##0.00") & "%", "T") & ",0,"
                    Else
                        Sql4 = Sql4 & DBSet(Rs!Concepto, "T") & ",0,"
                    End If
                    
                    '[Monica]29/04/2014: a�adidos los campos que faltaban
                    Sql4 = Sql4 & DBSet(Rs!conceptomo, "T") & "," & DBSet(Rs!importemo, "N") & "," & DBSet(Rs!Conceptoar1, "T") & "," & DBSet(Rs!importear1, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!Conceptoar2, "T") & "," & DBSet(Rs!importear2, "N") & "," & DBSet(Rs!conceptoar3, "T") & "," & DBSet(Rs!importear3, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!conceptoar4, "T") & "," & DBSet(Rs!importear4, "N") & "," & DBSet(Rs!difdias, "N") & "," & DBSet(Rs!calibre, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!Codpozo, "N") & "," & DBSet(Rs!PorcDto, "N") & "," & DBSet(Rs!ImpDto, "N") & "," & DBSet(Rs!Precio, "N") & "," & DBSet(Rs!pasaridoc, "N") & ","
                    
                    '[Monica]22/10/2012: guardamos tambien la partida [Monica]03/05/2013: ahora tb el poligono [Monica]22/07/2013: ahora tb metemos el nro de orden
                    Sql4 = Sql4 & DBSet(Rs!codparti, "N") & "," & DBSet(Rs!parcelas, "T") & "," & DBSet(Rs!poligono, "T") & "," & DBSet(Rs!nroorden, "N") '& ")"

                    '[Monica]02/09/2014: CONTADOSSSS
                    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                        If EsSocioContadoPOZOS(CStr(Rs2!Codsocio)) Then
                            Sql4 = Sql4 & ",1)"
                        Else
                            Sql4 = Sql4 & ",0)"
                        End If
                    Else
                        Sql4 = Sql4 & ")"
                    End If


                    conn.Execute Sql4
                    
                    If b Then b = InsertResumen(tipoMov, CStr(numFac))
                    
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                If b Then
                
                    vPorcen = DevuelveValor("select porcentaje from rpozos_cooprop where codsocio = " & DBSet(Rs!Codsocio, "N") & " and hidrante = " & DBSet(Rs!Hidrante, "T"))
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
    
    Set Rs = Nothing

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
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Inicio As Long
Dim Fin As Long
Dim NroDig As Integer
Dim Limite As Long


    On Error GoTo eCalculoConsumoHidrante


    CalculoConsumoHidrante = False
    
    SQL = "select * from rpozos where hidrante = " & DBSet(Hidrante, "T")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
       Inicio = 0
       Fin = 0
       NroDig = DBLet(Rs!Digcontrol, "N")
       Limite = 10 ^ NroDig
       
       Inicio = DBLet(Rs!lect_ant, "N")
       Fin = CLng(txtcodigo(51).Text)
    
       If Fin >= Inicio Then
          Consumo = Fin - Inicio
       Else
          If MsgBox("� Es un reinicio de contador ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Or (Inicio - Fin >= Limite) Then
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
    MuestraError Err.Number, "C�lculo Consumo Hidrante", Err.Description
End Function
             
             
'En vTesoreria pondremos como estaba el recibo
'Es decir. El  msgbox que pondra al final lo guardo en esta variable
Private Function ComprobarCobroArimoney(vTesoreria As String, LEtra As String, Codfaccl As Long, Fecha As Date) As Boolean
Dim vR As ADODB.Recordset
Dim Cad As String


On Error GoTo EComprobarCobroArimoney
    
    ComprobarCobroArimoney = False
    Set vR = New ADODB.Recordset
    Cad = "Select * from scobro where numserie='" & LEtra & "'"
    Cad = Cad & " AND codfaccl =" & Codfaccl
    Cad = Cad & " AND fecfaccl =" & DBSet(Fecha, "F")
    
    '
    vTesoreria = ""
    vR.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If vR.EOF Then
        vTesoreria = "" '"NO se ha encotrado ningun vencimiento en la tesoreria"
    Else
        While Not vR.EOF
            Cad = ""
            If DBLet(vR!codrem, "T") <> "" Then
                Cad = "El cobro asociado a la factura esta remesado(" & vR!codrem & ")"
            Else
                If DBLet(vR!recedocu, "N") = 1 Then
                    Cad = "Documento recibido"
                Else
                    If DBLet(vR!Estacaja, "N") = 1 Then
                        Cad = "Cobrado por caja"
                    Else
                        If DBLet(vR!transfer, "N") = 1 Then
                            Cad = "Esta en una transferencia"
                        Else
                           If DBLet(vR!impcobro, "N") > 0 Then Cad = "Esta parcialmente cobrado: " & vR!impcobro
                        
                            
                                    'Si hubeira que poner mas coas iria aqui
                        End If 'transfer
                    End If 'estacaja
                End If 'recdedocu
            End If 'remesado
            If Cad <> "" Then vTesoreria = vTesoreria & "Vto: " & vR!numorden & "      " & Cad & vbCrLf
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

Private Function DevuelvePrecio(Zona As Integer) As Currency
Dim SQL As String
Dim Precio As Currency
Dim Rs As ADODB.Recordset
Dim Prec1Zona0 As Currency
Dim Prec2Zona0 As Currency
    
    PrecioTalla1 = 0
    PrecioTalla2 = 0
    ZonaTalla = 0
    Prec1Zona0 = 0
    Prec2Zona0 = 0
    
    SQL = "select precio1, precio2 from rzonas where codzonas = 0"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Prec1Zona0 = DBLet(Rs.Fields(0).Value, "N")
        Prec2Zona0 = DBLet(Rs.Fields(1).Value, "N")
    End If
    
    Set Rs = Nothing
    
    SQL = "select precio1, precio2 from rzonas where codzonas = " & DBSet(Zona, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        PrecioTalla1 = DBLet(Rs.Fields(0).Value, "N")
        PrecioTalla2 = DBLet(Rs.Fields(1).Value, "N")
        Precio = PrecioTalla1 + PrecioTalla2
        If PrecioTalla1 <> Prec1Zona0 Or PrecioTalla2 <> Prec2Zona0 Then
            ZonaTalla = Zona
        End If
    End If

    Set Rs = Nothing
    
'[Monica]10/04/2013: el precio lo sacamos de la tabla de zonas
'    If Zona = CInt(ComprobarCero(txtcodigo(79).Text)) And txtcodigo(79).Text <> "" Then
'        PrecioTalla1 = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(80).Text)))
'        PrecioTalla2 = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(81).Text)))
'        Precio = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(80).Text))) + CCur(ImporteSinFormato(ComprobarCero(txtcodigo(81).Text)))
'        ZonaTalla = Zona
'    ElseIf Zona = CInt(ComprobarCero(txtcodigo(82).Text)) And txtcodigo(82).Text <> "" Then
'        PrecioTalla1 = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(83).Text)))
'        PrecioTalla2 = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(84).Text)))
'        Precio = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(83).Text))) + CCur(ImporteSinFormato(ComprobarCero(txtcodigo(84).Text)))
'    ElseIf Zona = CInt(ComprobarCero(txtcodigo(85).Text)) And txtcodigo(85).Text <> "" Then
'        PrecioTalla1 = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(86).Text)))
'        PrecioTalla2 = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(87).Text)))
'        Precio = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(86).Text))) + CCur(ImporteSinFormato(ComprobarCero(txtcodigo(87).Text)))
'        ZonaTalla = Zona
'    Else
'        PrecioTalla1 = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(72).Text)))
'        PrecioTalla2 = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(66).Text)))
'        Precio = CCur(ImporteSinFormato(ComprobarCero(txtcodigo(72).Text))) + CCur(ImporteSinFormato(ComprobarCero(txtcodigo(66).Text)))
'        ZonaTalla = 0
'    End If
    
    DevuelvePrecio = Precio

End Function


Private Function CargarTablaPrecios() As Boolean
Dim SQL As String
Dim SqlValues As String

    On Error GoTo eCargarTablaPrecios

    CargarTablaPrecios = False

    SQL = "delete from rpretallapoz "
    conn.Execute SQL
    
    SqlValues = ""
    
    SQL = "insert ignore into rpretallapoz (codzonas, precio1, precio2) values "
    
    SqlValues = SqlValues & "(0," & DBSet(txtcodigo(72).Text, "N") & "," & DBSet(txtcodigo(66).Text, "N") & "),"
    
    If txtcodigo(79).Text <> "" Then
        SqlValues = SqlValues & "(" & DBSet(txtcodigo(79).Text, "N") & "," & DBSet(txtcodigo(80).Text, "N") & "," & DBSet(txtcodigo(81).Text, "N") & "),"
    End If
    
    If txtcodigo(82).Text <> "" Then
        SqlValues = SqlValues & "(" & DBSet(txtcodigo(82).Text, "N") & "," & DBSet(txtcodigo(83).Text, "N") & "," & DBSet(txtcodigo(84).Text, "N") & "),"
    End If
    
    If txtcodigo(85).Text <> "" Then
        SqlValues = SqlValues & "(" & DBSet(txtcodigo(85).Text, "N") & "," & DBSet(txtcodigo(86).Text, "N") & "," & DBSet(txtcodigo(87).Text, "N") & "),"
    End If

    If SqlValues <> "" Then
        conn.Execute SQL & Mid(SqlValues, 1, Len(SqlValues) - 1)
    End If
    
    CargarTablaPrecios = True
    Exit Function

eCargarTablaPrecios:
    MuestraError Err.Number, "Cargar Tabla Precios", Err.Description

End Function

Private Sub EnviarEMailMulti(cadwhere As String, cadTit As String, cadRpt As String, cadTabla As String)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer

On Error GoTo EEnviar

    Screen.MousePointer = vbHourglass
    
    SQL = "SELECT distinct rsocios.codsocio,nomsocio,maisocio "
    SQL = SQL & "FROM " & cadTabla
    SQL = SQL & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ' Primero la borro por si acaso
    SQL = " DROP TABLE IF EXISTS tmpMail;"
    conn.Execute SQL
    
    'creamos una temporal donde guardamos para cada proveedor que SI tiene
    'e-mail, el mail1 o el mail2 al que vamos a enviar
    SQL = "CREATE TEMPORARY TABLE tmpMail ( "
    SQL = SQL & "codusu INT(7) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "codprove INT(6) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "nomprove varchar(40)  DEFAULT '' NOT NULL, "
    SQL = SQL & "email varchar(40)  DEFAULT '' NOT NULL) "
    conn.Execute SQL
    
    cont = 0
    lista = ""
    
    While Not Rs.EOF
    'para cada cliente/proveedor enviamos un e-mail
        Cad1 = DBLet(Rs.Fields(2), "T") 'e-mail administracion
        
        If Cad1 <> "" Then 'HAY email --> ENVIAMOS e-mail
            Label9(10).Caption = Format(Rs.Fields(0), "000000") & " - " & Rs.Fields(1) & " - " & Rs.Fields(2)
            DoEvents

            With frmImprimir
                .OtrosParametros = CadParam
                .NumeroParametros = numParam
                
                SQL = "{rsocios.codsocio}=" & Rs.Fields(0)

                .Opcion = 86
                .FormulaSeleccion = SQL
                .EnvioEMail = True
                CadenaDesdeOtroForm = "GENERANDO"
                .Titulo = "Cartas Talla"
                .NombreRPT = cadRpt
                .ConSubInforme = True
                .Show vbModal

                If CadenaDesdeOtroForm = "" Then
                'si se ha generado el .pdf para enviar
                    SQL = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                    SQL = SQL & " VALUES (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "T") & "," & DBSet(Cad1, "T") & ")"
                    conn.Execute SQL
            
                    Me.Refresh
                    espera 0.4
                    cont = cont + 1
                    'Se ha generado bien el documento
                    'Lo copiamos sobre app.path & \temp
                    SQL = Rs.Fields(0) & ".pdf"
                    FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & SQL
                End If
            End With
            Label9(10).Caption = ""
            DoEvents
        End If
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
      
    If cont > 0 Then
        espera 0.4
        SQL = "Carta de Talla" & "|"
       
       
        frmEMail.Opcion = 2
        frmEMail.DatosEnvio = SQL
        frmEMail.CodCryst = IndRptReport
        frmEMail.Ficheros = ""
        frmEMail.EsCartaTalla = True
        frmEMail.Show vbModal

        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
        
        'Borrar la carpeta con temporales
        Kill App.Path & "\temp\*.pdf"
    End If
    
    Screen.MousePointer = vbDefault
   
EEnviar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Enviando Carta de Talla por e-mail", Err.Description
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
    End If
End Sub

Private Function FacturacionTallaPreviaESCALONA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
Dim TotalFac As Currency
Dim TotalZona As Currency

Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long
Dim Brazas As Long
Dim cadMen As String

Dim PorcDto As Currency
Dim Descuento As Currency
Dim CadValues As String
Dim Precio As Currency

Dim PrecioBrz As Currency
Dim SocioAnt As Long
Dim SqlPrec As String
Dim Nregs As Integer


    On Error GoTo eFacturacion

    FacturacionTallaPreviaESCALONA = False
    
'    tipoMov = "TAL"
    
    b = True
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    SQL = "delete from tmpinformes2 where codusu = " & vUsu.Codigo
    conn.Execute SQL

    '[Monica]13/03/2014: se factura al socio no al propietario antes era codpropiet
    SQL = "SELECT rcampos.codsocio codsocio, rcampos.codzonas, round(sum(rcampos.supcoope) / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    
    SQL = SQL & " group by 1, 2 having hanegada <> 0  "
    
    ' ordenado por socio, hidrante
    SQL = SQL & " order by codsocio, codzonas "
    
    Me.Pb5.visible = True
    Nregs = TotalRegistrosConsulta(SQL)
    CargarProgresNew Pb5, Nregs
    DoEvents
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        SocioAnt = DBLet(Rs!Codsocio, "N")
        numfactu = 0
    End If
    
    While Not Rs.EOF And b
        HayReg = True
        
        If SocioAnt <> DBLet(Rs!Codsocio, "N") Then
        
            numfactu = numfactu + 1
        
            baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
            ImpoIva = TotalFac - baseimpo
        
            'insertar en la tabla de recibos de pozos tmpinformes
            '                               codusu, numfactu,fecfactu,codsocio,baseimpo,codivapoz,porciva,imporiva, totalfac, concepto
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, codigo1, importe2, campo1, porcen1, importe3, importe4, nombre1) "
            SQL = SQL & " values (" & vUsu.Codigo & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(SocioAnt, "N") & ","
            SQL = SQL & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            SQL = SQL & DBSet(TotalFac, "N") & ","
            SQL = SQL & DBSet(txtcodigo(97).Text, "T") & ")"
            
            conn.Execute SQL
            
            ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
            SQL = "SELECT rcampos.codcampo, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
            SQL = SQL & " FROM  " & cTabla
            If cWhere <> "" Then
                SQL = SQL & " WHERE " & cWhere
                '[Monica]13/03/2014: hidrantes del socio, antes eran hidrantes del propietario
                SQL = SQL & " and rcampos.codsocio = " & DBSet(SocioAnt, "N")
            Else
                '[Monica]13/03/2014: hidrantes del socio, antes eran hidrantes del propietario
                SQL = SQL & " where rcampos.codsocio = " & DBSet(SocioAnt, "N")
            End If
                
            Set Rs8 = New ADODB.Recordset
            Rs8.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            '                                       numfactu, fecfactu,codcampo,hanegadas,precio1, precio2, codzona
            SQL = "insert into tmpinformes2 (codusu, importe1, fecha1, importe2, importe3, precio1, precio2, campo1) values  "
            CadValues = ""
            While Not Rs8.EOF
                Precio = DevuelvePrecio(DBLet(Rs8!codzonas, "N"))
                
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                CadValues = CadValues & DBSet(Rs8!CodCampo, "N") & "," & DBSet(Rs8!hanegada, "N") & ","
                CadValues = CadValues & DBSet(PrecioTalla1, "N") & "," & DBSet(PrecioTalla2, "N") & "," & DBSet(ZonaTalla, "N") & "),"
                
                Rs8.MoveNext
            Wend
            
            If CadValues <> "" Then
                CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
                conn.Execute SQL & CadValues
            End If
            Set Rs8 = Nothing
                
            
            baseimpo = 0
            ImpoIva = 0
            TotalFac = 0
        
            SocioAnt = DBLet(Rs!Codsocio, "N")
            
        End If
        
        Acciones = DBLet(Rs!hanegada, "N")
        
        Precio = DevuelvePrecio(Rs!codzonas)
        
        TotalFac = TotalFac + Round2(Acciones * Precio, 2)
        
        IncrementarProgresNew Pb5, 1
        
'        Label2(78).Caption = "Socio: " & Format(Rs!Codsocio, "000000")
        DoEvents
        
        Rs.MoveNext
    Wend
    
    If HayReg And b Then
        numfactu = numfactu + 1
            
        baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - baseimpo
    
        'insertar en la tabla de recibos de pozos (intermedia)
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, codigo1, importe2, campo1, porcen1, importe3, importe4, nombre1) "
        SQL = SQL & " values (" & vUsu.Codigo & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(SocioAnt, "N") & ","
        SQL = SQL & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        SQL = SQL & DBSet(TotalFac, "N") & ","
        SQL = SQL & DBSet(txtcodigo(97).Text, "T") & ")"
        
        conn.Execute SQL
        
        ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
        SQL = "SELECT rcampos.codcampo, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
        SQL = SQL & " FROM  " & cTabla
        If cWhere <> "" Then
            SQL = SQL & " WHERE " & cWhere
            '[Monica]13/03/2014: hidrantes del socio, antes eran hidrantes del propietario
            SQL = SQL & " and rcampos.codsocio = " & DBSet(SocioAnt, "N")
        Else
            SQL = SQL & " where rcampos.codsocio = " & DBSet(SocioAnt, "N")
        End If
            
        Set Rs8 = New ADODB.Recordset
        Rs8.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        '                                       numfactu, fecfactu,codcampo,hanegadas,precio1, precio2, codzona
        SQL = "insert into tmpinformes2 (codusu, importe1, fecha1, importe2, importe3, precio1, precio2, campo1) values  "
        CadValues = ""
        While Not Rs8.EOF
            Precio = DevuelvePrecio(DBLet(Rs8!codzonas, "N"))
            
            CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
            CadValues = CadValues & DBSet(Rs8!CodCampo, "N") & "," & DBSet(Rs8!hanegada, "N") & ","
            CadValues = CadValues & DBSet(PrecioTalla1, "N") & "," & DBSet(PrecioTalla2, "N") & "," & DBSet(ZonaTalla, "N") & "),"
            
            Rs8.MoveNext
        Wend
        
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute SQL & CadValues
        End If
        Set Rs8 = Nothing
    
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        FacturacionTallaPreviaESCALONA = False
    Else
        FacturacionTallaPreviaESCALONA = True
    End If
    Me.Pb5.visible = False
End Function

Private Sub MostrarContadoresANoFacturar(cTabla As String, cSelect As String)
Dim SQL As String


    SQL = "select rpozos.hidrante from " & cTabla & " where (rpozos.consumo < " & DBSet(vParamAplic.ConsumoMinPOZ, "N") & " or rpozos.consumo > " & DBSet(vParamAplic.ConsumoMaxPOZ, "N") & ") "
    If cSelect <> "" Then SQL = SQL & " and " & cSelect
    
    If TotalRegistrosConsulta(SQL) <> 0 Then
        
        Set frmMens3 = New frmMensajes
        
        frmMens3.OpcionMensaje = 50
        frmMens3.cadwhere = " and rpozos.hidrante in (" & SQL & ")"
        frmMens3.Show vbModal
    
        Set frmMens3 = Nothing
        
    Else
    
        Continuar = True
    
    End If
    
End Sub



Private Sub InsertarTemporal(cadwhere As String, cadSelect As String)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer
Dim Sql2 As String

    
    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    'seleccionamos todos los socios a los que queremos enviar e-mail
    SQL = "SELECT distinct " & vUsu.Codigo & ", rsocios.codsocio, rrecibpozos.codtipom, rrecibpozos.numfactu, rrecibpozos.fecfactu  from rsocios, rrecibpozos where rrecibpozos.codsocio in (" & cadwhere & ")"
    SQL = SQL & " and rsocios.codsocio = rrecibpozos.codsocio "
    SQL = SQL & " and " & cadSelect
    
    Sql2 = "insert into tmpinformes (codusu, codigo1, nombre1, importe1, fecha1) " & SQL
    conn.Execute Sql2

End Sub



Private Function TotalSocios(cTabla As String, cWhere As String) As Long
Dim SQL As String

    TotalSocios = 0
    
    SQL = "SELECT  count(distinct rsocios.codsocio) "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If

    TotalSocios = TotalRegistros(SQL)

End Function


Private Function TotalRegistrosIndefa(vSQL As String) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalRegistrosIndefa = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalRegistrosIndefa = Rs.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        TotalRegistrosIndefa = 0
        Err.Clear
    End If
End Function


Private Function CargarTemporalCCCErroneas(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Cad2 As String
Dim Cad3 As String
Dim CadValues As String
Dim cadInsert As String
Dim Contador As String
Dim Nregs As Integer
Dim Fecha As Date
Dim DDCC As Integer
Dim CC As String
Dim Ent As String ' Entidad
Dim Suc As String ' Oficina
Dim DC As String ' Digitos de control
Dim I, i2, i3, i4 As Integer
Dim NumCC As String ' N�mero de cuenta propiamente dicho
Dim BuscaChekc As String
    
    On Error GoTo eCargarTemporal
    
    CargarTemporalCCCErroneas = False
    
    Screen.MousePointer = vbHourglass
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
    End If
    
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    Label2(102).visible = True
    DoEvents
    
    cadInsert = "insert into tmpinformes (codusu, codigo1, nombre1, nombre2)  VALUES "
    
    SQL = "select codsocio,codbanco,codsucur,digcontr,cuentaba, iban from rsocios "
    SQL = SQL & "where cuentaba <> '8888888888' "
    If cWhere <> "" Then SQL = SQL & " and  " & cWhere
    SQL = SQL & " order by codsocio "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    CadValues = ""

    While Not Rs.EOF
        Label2(102).Caption = "Socio : " & Format(DBLet(Rs!Codsocio), "000000")
        DoEvents
    
        If IsNumeric(DBLet(Rs!CuentaBa)) And IsNumeric(DBLet(Rs!CodBanco)) And IsNumeric(DBLet(Rs!CodSucur)) Then
            
            If Not IsNumeric(DBLet(Rs!digcontr)) Then
                DDCC = 0
            Else
                DDCC = DBLet(Rs!digcontr)
            End If
            
        
            CC = Format(DBLet(Rs!CodBanco), "0000") & Format(DBLet(Rs!CodSucur), "0000") & Format(DDCC, "00") & Format(DBLet(Rs!CuentaBa), "0000000000")
            
            If Not Comprueba_CC(CC) Then
                
                '-- Calculamos el primer d�gito de control
                I = Val(Mid(CC, 1, 1)) * 4
                I = I + Val(Mid(CC, 2, 1)) * 8
                I = I + Val(Mid(CC, 3, 1)) * 5
                I = I + Val(Mid(CC, 4, 1)) * 10
                I = I + Val(Mid(CC, 5, 1)) * 9
                I = I + Val(Mid(CC, 6, 1)) * 7
                I = I + Val(Mid(CC, 7, 1)) * 3
                I = I + Val(Mid(CC, 8, 1)) * 6
                i2 = Int(I / 11)
                i3 = I - (i2 * 11)
                i4 = 11 - i3
                Select Case i4
                    Case 11
                        i4 = 0
                    Case 10
                        i4 = 1
                End Select
                
                DC = i4
                
                '-- Calculamos el segundo d�gito de control
                I = Val(Mid(CC, 11, 1)) * 1
                I = I + Val(Mid(CC, 12, 1)) * 2
                I = I + Val(Mid(CC, 13, 1)) * 4
                I = I + Val(Mid(CC, 14, 1)) * 8
                I = I + Val(Mid(CC, 15, 1)) * 5
                I = I + Val(Mid(CC, 16, 1)) * 10
                I = I + Val(Mid(CC, 17, 1)) * 9
                I = I + Val(Mid(CC, 18, 1)) * 7
                I = I + Val(Mid(CC, 19, 1)) * 3
                I = I + Val(Mid(CC, 20, 1)) * 6
                i2 = Int(I / 11)
                i3 = I - (i2 * 11)
                i4 = 11 - i3
                Select Case i4
                    Case 11
                        i4 = 0
                    Case 10
                        i4 = 1
                End Select
                
                DC = DC & i4
            
                If DC <> DBLet(Rs!digcontr) Then
                    BuscaChekc = ""
                    If DBLet(Rs!Iban, "T") <> "" Then BuscaChekc = Mid(Rs!Iban, 1, 2)
                    
                    CC = Format(DBLet(Rs!CodBanco), "0000") & Format(DBLet(Rs!CodSucur), "0000") & Format(DC, "00") & Format(DBLet(Rs!CuentaBa), "0000000000")
                            
                    
                    If DevuelveIBAN2(BuscaChekc, CC, CC) Then
                        CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs!Codsocio, "N") & "," & DBSet(DC, "T") & "," & DBSet(BuscaChekc & CC, "T") & "),"
                    End If
                    
'22/11/2013:antes
'                    CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(RS!Codsocio, "N") & "," & DBSet(DC, "T") & "),"
                    
                End If
                
            Else
            
                '[Monica]22/11/2013: comprobamos el iban
                BuscaChekc = ""
                If DBLet(Rs!Iban, "T") <> "" Then BuscaChekc = Mid(Rs!Iban, 1, 2)
                
                If DevuelveIBAN2(BuscaChekc, CC, CC) Then
                    If BuscaChekc & CC <> DBLet(Rs!Iban, "T") Then
                        CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs!Codsocio, "N") & "," & ValorNulo & "," & DBSet(BuscaChekc & CC, "T") & "),"
                    End If
                End If
                
                
            End If
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute cadInsert & CadValues
    End If
        
    Label2(102).visible = False
    DoEvents
    CargarTemporalCCCErroneas = True
    Screen.MousePointer = vbDefault

    Exit Function

eCargarTemporal:
    Label2(102).visible = False
    DoEvents
    Screen.MousePointer = vbDefault
    MuestraError Err.Description, "Cargar Temporal CCC Err�neas", Err.Description
End Function


Private Function FacturacionConsumoMantaESCALONA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
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
Dim cadMen As String

Dim PorcDto As Currency
Dim Descuento As Currency
Dim CadValues As String
Dim I As Integer

    On Error GoTo eFacturacion

    FacturacionConsumoMantaESCALONA = False
    
    tipoMov = "RMT"
    
    conn.BeginTrans
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    'hacemos una factura por socio campo
    SQL = "SELECT rcampos.codsocio, rcampos.codcampo, rpozauxmanta.nroimpresion, sum(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) hanegada  "
    SQL = SQL & " FROM  (" & cTabla & ") INNER JOIN rpozauxmanta On rcampos.codcampo = rpozauxmanta.codcampo "

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    
    SQL = SQL & " group by 1, 2, 3 having sum(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) <> 0 "
    
    ' ordenado por socio, hidrante
    SQL = SQL & " order by codsocio, codcampo "
    
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And b
        HayReg = True
        For I = 1 To DBLet(Rs!nroimpresion, "N")
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            baseimpo = 0
            ImpoIva = 0
            TotalFac = 0
            
            Acciones = DBLet(Rs!hanegada, "N")
            
    '        Brazas = (Int(Acciones) * 200) + ((Acciones - Int(Acciones)) * 1000)
    '        Brazas = Acciones * 200
    
            TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtcodigo(112).Text)), 2)
            
        
            '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
            baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
            ImpoIva = TotalFac - baseimpo
        
            IncrementarProgresNew Pb7, 1
            
            'insertar en la tabla de recibos de pozos
            SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
            SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
            SQL = SQL & "concepto, contabilizado, porcdto, impdto, precio) "
            SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
            SQL = SQL & ValorNulo & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            SQL = SQL & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(txtcodigo(113).Text, "T") & ",0,"
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(CCur(ImporteSinFormato(txtcodigo(112).Text)), "N") & ")"
            
            conn.Execute SQL
                
                
            ' Introducimos en la tabla de lineas de campos que intervienen en la factura para la impresion
            ' SOLO HABRA UN CAMPO
            SQL = "SELECT rcampos.codcampo, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
            SQL = SQL & " FROM  " & cTabla '& ") INNER JOIN rcampos ON rpozos.codcampo = rcampos.codcampo"
            If cWhere <> "" Then
                SQL = SQL & " WHERE " & cWhere
                SQL = SQL & " and rcampos.codsocio = " & DBSet(Rs!Codsocio, "N")
                SQL = SQL & " and rcampos.codcampo = " & DBSet(Rs!CodCampo, "N")
            Else
                SQL = SQL & " where rcampos.codsocio = " & DBSet(Rs!Codsocio, "N")
                SQL = SQL & " and rcampos.codcampo = " & DBSet(Rs!CodCampo, "N")
            End If
                
            Set Rs8 = New ADODB.Recordset
            Rs8.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = "insert into rrecibpozos_cam (codtipom, numfactu, fecfactu, codcampo, hanegada, precio1, codzonas, poligono, parcela, subparce) values  "
            CadValues = ""
            While Not Rs8.EOF
                CadValues = CadValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                CadValues = CadValues & DBSet(Rs8!CodCampo, "N") & "," & DBSet(Rs8!hanegada, "N") & "," & DBSet(txtcodigo(112).Text, "N") & ","
                CadValues = CadValues & DBSet(Rs8!codzonas, "N") & "," & DBSet(Rs8!poligono, "N") & "," & DBSet(Rs8!Parcela, "N") & "," & DBSet(Rs8!SubParce, "T")
                CadValues = CadValues & "),"
                Rs8.MoveNext
            Wend
            If CadValues <> "" Then
                CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
                conn.Execute SQL & CadValues
            End If
            Set Rs8 = Nothing
                
            If b Then b = InsertResumen(tipoMov, CStr(numfactu))
            
            If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        Next I
        Rs.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionConsumoMantaESCALONA = False
    Else
        conn.CommitTrans
        FacturacionConsumoMantaESCALONA = True
    End If
End Function


Private Function FacturacionConsumoMantaESCALONANew(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
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
Dim cadMen As String

Dim PorcDto As Currency
Dim Descuento As Currency
Dim CadValues As String
Dim I As Integer

Dim Precio As Currency

    On Error GoTo eFacturacion

    FacturacionConsumoMantaESCALONANew = False
    
    tipoMov = "ALV"
    
    conn.BeginTrans
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
'[Monica]20/04/2015: ya no actualizamos nada �?
'    ' actualizamos el precio de recibo a manta
'    Sql = "update rtipopozos set imporcuotahda = " & DBSet(CCur(ImporteSinFormato(txtCodigo(112).Text)), "N")
'    Sql = Sql & " where codpozo = 1"
'    conn.Execute Sql


    'hacemos una factura por socio campo
    SQL = "SELECT rcampos.codsocio, rcampos.codcampo, rpozauxmanta.nroimpresion, rpozauxmanta.hanegadas, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, rzonas.preciomanta  "
    SQL = SQL & " FROM  (" & cTabla & ") INNER JOIN rpozauxmanta On rcampos.codcampo = rpozauxmanta.codcampo "

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    
    SQL = SQL & " group by 1, 2, 3 having rpozauxmanta.hanegadas <> 0 "
    
    ' ordenado por socio, hidrante
    SQL = SQL & " order by codsocio, codcampo "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    b = True
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And b
        HayReg = True
        For I = 1 To DBLet(Rs!nroimpresion, "N")
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rpozticketsmanta", "numalbar", "numalbar", CStr(numfactu), "N", , "fecalbar", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            baseimpo = 0
            ImpoIva = 0
            TotalFac = 0
            
            Acciones = DBLet(Rs!Hanegadas, "N")
            Precio = DBLet(Rs!preciomanta, "N")
            '[Monica]20/04/2015: el precio lo cogemos de las zonas
            'TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtCodigo(112).Text)), 2)
            TotalFac = Round2(Acciones * Precio, 2)
        
            '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
            baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
            ImpoIva = TotalFac - baseimpo
        
            IncrementarProgresNew Pb7, 1
            
            'insertar en la tabla de tickets de pozos
            SQL = "insert into rpozticketsmanta (numalbar,fecalbar,codsocio,codcampo,hanegada,precio1,importe,codzonas,poligono,parcela,subparce,fecriego,fecpago,concepto) "
            SQL = SQL & " values (" & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
            SQL = SQL & DBSet(Rs!CodCampo, "N") & "," & DBSet(Rs!Hanegadas, "N") & "," & DBSet(Precio, "N") & "," ' DBSet(CCur(ImporteSinFormato(txtCodigo(112).Text)), "N") & ","
            SQL = SQL & DBSet(TotalFac, "N") & "," & DBSet(Rs!codzonas, "N") & "," & DBSet(Rs!poligono, "N") & "," & DBSet(Rs!Parcela, "N") & "," & DBSet(Rs!SubParce, "T")
            SQL = SQL & "," & ValorNulo & "," & ValorNulo & "," & DBSet(txtcodigo(113).Text, "T") & ")"
            
            conn.Execute SQL
                
            SQL = "insert into tmpinformes (codusu, nombre1, importe1) values ( " & vUsu.Codigo & ",'ALV'," & DBSet(numfactu, "N") & ")"
            conn.Execute SQL
                
                
            If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        Next I
        Rs.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionConsumoMantaESCALONANew = False
    Else
        conn.CommitTrans
        FacturacionConsumoMantaESCALONANew = True
    End If
End Function




