VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
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
   Begin VB.Frame FrameComprobacion 
      Height          =   3885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6945
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
         Left            =   5610
         TabIndex        =   6
         Top             =   3195
         Width           =   1095
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
         Left            =   4410
         TabIndex        =   5
         Top             =   3210
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1590
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1200
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2640
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2250
         Width           =   1350
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
         Left            =   810
         TabIndex        =   13
         Top             =   1200
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
         Left            =   810
         TabIndex        =   12
         Top             =   1605
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hidrante"
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
         Left            =   540
         TabIndex        =   11
         Top             =   870
         Width           =   825
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
         Left            =   840
         TabIndex        =   9
         Top             =   2280
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
         Left            =   840
         TabIndex        =   8
         Top             =   2685
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
         Left            =   570
         TabIndex        =   7
         Top             =   1980
         Width           =   600
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1545
         Picture         =   "frmPOZListado.frx":000C
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1560
         Picture         =   "frmPOZListado.frx":0097
         Top             =   2250
         Width           =   240
      End
   End
   Begin VB.Frame FrameExporLecturas 
      Height          =   3735
      Left            =   0
      TabIndex        =   501
      Top             =   0
      Width           =   7350
      Begin VB.TextBox txtcodigo 
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
         Index           =   134
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   503
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1560
         Width           =   4830
      End
      Begin VB.CommandButton CmdAcepExportar 
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
         Left            =   4680
         TabIndex        =   504
         Top             =   3060
         Width           =   1095
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
         Index           =   16
         Left            =   5850
         TabIndex        =   505
         Top             =   3045
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   121
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   502
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1065
         Width           =   1230
      End
      Begin MSComctlLib.ProgressBar Pb8 
         Height          =   255
         Left            =   315
         TabIndex        =   511
         Top             =   2070
         Visible         =   0   'False
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Fichero"
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
         Height          =   255
         Index           =   60
         Left            =   330
         TabIndex        =   510
         Top             =   1560
         Width           =   780
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   25
         Left            =   1890
         Picture         =   "frmPOZListado.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   1065
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Lectura"
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
         Height          =   255
         Index           =   58
         Left            =   315
         TabIndex        =   509
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label22 
         Caption         =   "Exportacion de Lecturas"
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
         Left            =   315
         TabIndex        =   508
         Top             =   315
         Width           =   3900
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   507
         Top             =   2685
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   315
         Index           =   2
         Left            =   360
         TabIndex        =   506
         Top             =   2385
         Width           =   6555
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   34
         Left            =   1890
         MouseIcon       =   "frmPOZListado.frx":01AD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar fichero"
         Top             =   1575
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
         Caption         =   "€/hanegada"
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
         Caption         =   "€/hanegada"
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
   Begin VB.Frame FrameCartaTallas 
      Height          =   8415
      Left            =   -30
      TabIndex        =   217
      Top             =   30
      Width           =   8250
      Begin VB.TextBox txtcodigo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   97
         Left            =   2505
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   221
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
         Top             =   2520
         Width           =   5400
      End
      Begin VB.Frame Frame13 
         Height          =   900
         Left            =   360
         TabIndex        =   317
         Top             =   6510
         Width           =   7575
         Begin VB.OptionButton OptMail 
            Caption         =   "Enviar por e-mail e imprimir a los socios sin correo"
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
            Index           =   0
            Left            =   165
            TabIndex        =   319
            Top             =   360
            Width           =   5325
         End
         Begin VB.OptionButton OptMail 
            Caption         =   "Imprimir Todos"
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
            Left            =   5670
            TabIndex        =   318
            Top             =   360
            Value           =   -1  'True
            Width           =   1785
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
            Caption         =   "Administración"
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
         Caption         =   "Datos de Impresión"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2955
         Left            =   360
         TabIndex        =   305
         Top             =   3480
         Width           =   7590
         Begin VB.TextBox txtcodigo 
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
            Index           =   94
            Left            =   2190
            MaxLength       =   35
            TabIndex        =   228
            Top             =   2520
            Width           =   5220
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   93
            Left            =   2190
            MaxLength       =   35
            TabIndex        =   227
            Top             =   2160
            Width           =   5220
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   92
            Left            =   2190
            MaxLength       =   35
            TabIndex        =   226
            Top             =   1800
            Width           =   5220
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   91
            Left            =   2190
            MaxLength       =   40
            TabIndex        =   225
            Top             =   1440
            Width           =   5220
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   90
            Left            =   2190
            MaxLength       =   40
            TabIndex        =   224
            Top             =   1080
            Width           =   5220
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   89
            Left            =   2190
            MaxLength       =   40
            TabIndex        =   223
            Top             =   720
            Width           =   5220
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   88
            Left            =   2190
            MaxLength       =   40
            TabIndex        =   222
            Top             =   360
            Width           =   5220
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Recargos"
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
            Index           =   88
            Left            =   210
            TabIndex        =   312
            Top             =   2520
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Período voluntario"
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
            Index           =   87
            Left            =   210
            TabIndex        =   311
            Top             =   2160
            Width           =   1770
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bonificación"
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
            Index           =   86
            Left            =   210
            TabIndex        =   310
            Top             =   1800
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Prohibición"
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
            Index           =   85
            Left            =   210
            TabIndex        =   309
            Top             =   1440
            Width           =   1725
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Fin Comunic."
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
            Index           =   84
            Left            =   210
            TabIndex        =   308
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio Pago"
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
            Index           =   83
            Left            =   210
            TabIndex        =   307
            Top             =   720
            Width           =   1755
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Junta"
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
            Index           =   82
            Left            =   210
            TabIndex        =   306
            Top             =   360
            Width           =   1215
         End
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
         Index           =   68
         Left            =   3525
         Locked          =   -1  'True
         TabIndex        =   237
         Text            =   "Text5"
         Top             =   1515
         Width           =   4410
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
         Index           =   67
         Left            =   3525
         Locked          =   -1  'True
         TabIndex        =   236
         Text            =   "Text5"
         Top             =   1110
         Width           =   4410
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   69
         Left            =   2505
         MaxLength       =   10
         TabIndex        =   220
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   67
         Left            =   2490
         MaxLength       =   10
         TabIndex        =   218
         Top             =   1110
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   68
         Left            =   2490
         MaxLength       =   10
         TabIndex        =   219
         Top             =   1500
         Width           =   960
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
         Left            =   5640
         TabIndex        =   229
         Top             =   7785
         Width           =   1095
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
         Left            =   6840
         TabIndex        =   230
         Top             =   7770
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar Pb5 
         Height          =   255
         Left            =   360
         TabIndex        =   342
         Top             =   7440
         Visible         =   0   'False
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
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
         Index           =   93
         Left            =   570
         TabIndex        =   341
         Top             =   2535
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precios Zona"
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
         Index           =   92
         Left            =   570
         TabIndex        =   340
         Top             =   3030
         Width           =   1275
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   2145
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
         Left            =   2175
         MouseIcon       =   "frmPOZListado.frx":02FF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1515
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   2175
         MouseIcon       =   "frmPOZListado.frx":0451
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   2175
         Picture         =   "frmPOZListado.frx":05A3
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
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
         Index           =   60
         Left            =   570
         TabIndex        =   235
         Top             =   1980
         Width           =   1320
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
         Index           =   57
         Left            =   540
         TabIndex        =   233
         Top             =   960
         Width           =   540
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
         Index           =   56
         Left            =   1425
         TabIndex        =   232
         Top             =   1515
         Width           =   555
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
         Index           =   55
         Left            =   1425
         TabIndex        =   231
         Top             =   1110
         Width           =   600
      End
   End
   Begin VB.Frame FrameComprobacionDatos 
      Height          =   5685
      Left            =   30
      TabIndex        =   343
      Top             =   0
      Width           =   8400
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
         Left            =   120
         TabIndex        =   346
         Top             =   780
         Width           =   8085
         Begin VB.TextBox txtcodigo 
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
            Index           =   98
            Left            =   1275
            MaxLength       =   10
            TabIndex        =   354
            Top             =   540
            Width           =   1395
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   99
            Left            =   1275
            MaxLength       =   10
            TabIndex        =   355
            Top             =   900
            Width           =   1395
         End
         Begin VB.Frame Frame14 
            Caption         =   "Tipo Informe"
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
            Height          =   2625
            Left            =   105
            TabIndex        =   347
            Top             =   1290
            Width           =   7875
            Begin VB.OptionButton Option4 
               Caption         =   "Diferencias entre Datos de Contadores y Campos"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   0
               Left            =   195
               TabIndex        =   353
               Top             =   1272
               Width           =   5280
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Diferencias entre Escalona e Indefa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   1
               Left            =   195
               TabIndex        =   352
               Top             =   255
               Width           =   4065
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Contadores que existen en Indefa y no en Escalona"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   2
               Left            =   195
               TabIndex        =   351
               Top             =   594
               Width           =   5685
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Contadores con Socio Bloqueado"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   4
               Left            =   195
               TabIndex        =   350
               Top             =   1611
               Width           =   3885
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Contadores con consumo que están en Inelcom y no están en Escalona"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   5
               Left            =   195
               TabIndex        =   349
               Top             =   1950
               Width           =   7635
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Contadores que existen en Escalona y no en Indefa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   3
               Left            =   195
               TabIndex        =   348
               Top             =   933
               Width           =   5505
            End
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
            Index           =   94
            Left            =   600
            TabIndex        =   360
            Top             =   570
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
            Index           =   95
            Left            =   600
            TabIndex        =   359
            Top             =   930
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hidrante"
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
            Index           =   96
            Left            =   240
            TabIndex        =   358
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.CommandButton CmdAceptarComp 
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
         Left            =   5850
         TabIndex        =   356
         Top             =   5130
         Width           =   1095
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
         Index           =   9
         Left            =   7065
         TabIndex        =   357
         Top             =   5115
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar pb6 
         Height          =   255
         Left            =   210
         TabIndex        =   361
         Top             =   4800
         Width           =   7935
         _ExtentX        =   13996
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
         Caption         =   "Informe de Comprobación de Datos"
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
         Left            =   270
         TabIndex        =   344
         Top             =   300
         Width           =   5925
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
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1305
         Left            =   3600
         TabIndex        =   267
         Top             =   1920
         Visible         =   0   'False
         Width           =   2145
         Begin VB.OptionButton Option1 
            Caption         =   "Contado"
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
            Index           =   2
            Left            =   240
            TabIndex        =   270
            Top             =   345
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Efecto"
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
            Index           =   3
            Left            =   240
            TabIndex        =   269
            Top             =   615
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
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
            Index           =   4
            Left            =   240
            TabIndex        =   268
            Top             =   885
            Width           =   1380
         End
      End
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
         Index           =   0
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   1365
         Width           =   2250
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   39
         Left            =   1530
         MaxLength       =   7
         TabIndex        =   119
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1755
         Width           =   1290
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   38
         Left            =   1530
         MaxLength       =   7
         TabIndex        =   118
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1365
         Width           =   1290
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   37
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   123
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1290
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   36
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   121
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1290
      End
      Begin VB.CommandButton cmdCancelReimp 
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
         Left            =   5295
         TabIndex        =   117
         Top             =   4275
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptarReimp 
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
         Left            =   4125
         TabIndex        =   116
         Top             =   4275
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1530
         MaxLength       =   6
         TabIndex        =   127
         Top             =   3765
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1530
         MaxLength       =   6
         TabIndex        =   125
         Top             =   3390
         Width           =   1050
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
         Index           =   35
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "Text5"
         Top             =   3765
         Width           =   3810
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
         Index           =   34
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "Text5"
         Top             =   3390
         Width           =   3810
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
         Left            =   360
         TabIndex        =   134
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
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
         Index           =   3
         Left            =   585
         TabIndex        =   133
         Top             =   1755
         Width           =   645
      End
      Begin VB.Label Label4 
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
         Index           =   1
         Left            =   585
         TabIndex        =   132
         Top             =   1395
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Factura"
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
         Index           =   2
         Left            =   360
         TabIndex        =   131
         Top             =   1110
         Width           =   1170
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
         Height          =   255
         Index           =   16
         Left            =   330
         TabIndex        =   130
         Top             =   2115
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         Left            =   555
         TabIndex        =   129
         Top             =   2415
         Width           =   690
      End
      Begin VB.Label Label4 
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
         Index           =   14
         Left            =   555
         TabIndex        =   128
         Top             =   2775
         Width           =   645
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1260
         Picture         =   "frmPOZListado.frx":062E
         ToolTipText     =   "Buscar fecha"
         Top             =   2415
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   1260
         Picture         =   "frmPOZListado.frx":06B9
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   13
         Left            =   540
         TabIndex        =   126
         Top             =   3405
         Width           =   690
      End
      Begin VB.Label Label4 
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
         Index           =   12
         Left            =   555
         TabIndex        =   124
         Top             =   3780
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   11
         Left            =   375
         TabIndex        =   122
         Top             =   3165
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1245
         MouseIcon       =   "frmPOZListado.frx":0744
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3765
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1245
         MouseIcon       =   "frmPOZListado.frx":0896
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3390
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
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
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   120
         Top             =   1110
         Width           =   1815
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
            Name            =   "Verdana"
            Size            =   9.75
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
            Index           =   7
            Left            =   420
            TabIndex        =   274
            Top             =   840
            Width           =   1245
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Efecto"
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
            Index           =   6
            Left            =   420
            TabIndex        =   273
            Top             =   570
            Width           =   1365
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Contado"
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
            Index           =   5
            Left            =   420
            TabIndex        =   272
            Top             =   300
            Width           =   1365
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ordenado por"
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
         ForeColor       =   &H00972E0B&
         Height          =   705
         Left            =   390
         TabIndex        =   180
         Top             =   5130
         Width           =   3600
         Begin VB.OptionButton Option3 
            Caption         =   "Nro.Factura"
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
            Left            =   1800
            TabIndex        =   182
            Top             =   315
            Width           =   1530
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Socio"
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
            Left            =   210
            TabIndex        =   181
            Top             =   315
            Value           =   -1  'True
            Width           =   1305
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Resumen Facturación"
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
         Left            =   375
         TabIndex        =   144
         Top             =   4740
         Width           =   2490
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   50
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   142
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3825
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   49
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   141
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3450
         Width           =   1050
      End
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
         Index           =   40
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   137
         Top             =   1230
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   41
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   138
         Top             =   1605
         Width           =   830
      End
      Begin VB.CommandButton cmdAceptarListFact 
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
         Left            =   4230
         TabIndex        =   145
         Top             =   5355
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelList 
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
         Left            =   5400
         TabIndex        =   146
         Top             =   5340
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   42
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   139
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   43
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   140
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1050
      End
      Begin VB.Label Label4 
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
         Index           =   59
         Left            =   765
         TabIndex        =   513
         Top             =   3465
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Index           =   55
         Left            =   765
         TabIndex        =   512
         Top             =   3825
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "Nro. Factura"
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
         Height          =   255
         Index           =   6
         Left            =   345
         TabIndex        =   179
         Top             =   3180
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo Factura"
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
         Height          =   255
         Index           =   4
         Left            =   345
         TabIndex        =   156
         Top             =   4290
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":09E8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":0B3A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   19
         Left            =   345
         TabIndex        =   155
         Top             =   1005
         Width           =   540
      End
      Begin VB.Label Label4 
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
         Index           =   18
         Left            =   780
         TabIndex        =   154
         Top             =   1620
         Width           =   555
      End
      Begin VB.Label Label4 
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
         Index           =   17
         Left            =   765
         TabIndex        =   153
         Top             =   1245
         Width           =   600
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1485
         Picture         =   "frmPOZListado.frx":0C8C
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   1485
         Picture         =   "frmPOZListado.frx":0D17
         ToolTipText     =   "Buscar fecha"
         Top             =   2415
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Left            =   780
         TabIndex        =   152
         Top             =   2775
         Width           =   555
      End
      Begin VB.Label Label4 
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
         Index           =   9
         Left            =   780
         TabIndex        =   151
         Top             =   2415
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
         Height          =   255
         Index           =   8
         Left            =   345
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
         Left            =   345
         TabIndex        =   149
         Top             =   315
         Width           =   5160
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
         Index           =   100
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   373
         Text            =   "Text5"
         Top             =   1290
         Width           =   3645
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
         Index           =   101
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   372
         Text            =   "Text5"
         Top             =   1695
         Width           =   3645
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   101
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   368
         Top             =   1695
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   100
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   367
         Top             =   1290
         Width           =   960
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
         Index           =   10
         Left            =   5385
         TabIndex        =   364
         Top             =   2475
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptarCompCCC 
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
         Left            =   4140
         TabIndex        =   363
         Top             =   2490
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Procesando"
         Height          =   195
         Index           =   102
         Left            =   570
         TabIndex        =   374
         Top             =   2220
         Visible         =   0   'False
         Width           =   5925
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":0DA2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1725
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":0EF4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   100
         Left            =   570
         TabIndex        =   371
         Top             =   990
         Width           =   540
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
         Index           =   99
         Left            =   795
         TabIndex        =   370
         Top             =   1725
         Width           =   600
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
         Index           =   98
         Left            =   795
         TabIndex        =   369
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label16 
         Caption         =   "Informe de Cuentas Bancarias Erróneas"
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
            Height          =   660
            Left            =   0
            TabIndex        =   169
            Top             =   1590
            Width           =   6525
            Begin VB.TextBox txtcodigo 
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
               Left            =   1320
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
               Index           =   35
               Left            =   330
               TabIndex        =   170
               Top             =   180
               Width           =   945
            End
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   62
            Top             =   240
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
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
            Left            =   4080
            TabIndex        =   63
            Top             =   0
            Width           =   2235
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Recibo"
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
            Index           =   3
            Left            =   4080
            TabIndex        =   64
            Top             =   330
            Width           =   2265
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   1290
            Left            =   270
            TabIndex        =   171
            Top             =   480
            Width           =   3555
            Begin VB.TextBox txtcodigo 
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
               Left            =   1290
               MaxLength       =   10
               TabIndex        =   175
               Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
               Top             =   525
               Width           =   1005
            End
            Begin VB.TextBox txtcodigo 
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
               Left            =   1290
               MaxLength       =   10
               TabIndex        =   174
               Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
               Top             =   915
               Width           =   1005
            End
            Begin VB.TextBox txtcodigo 
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
               Left            =   2400
               MaxLength       =   10
               TabIndex        =   173
               Top             =   525
               Width           =   1005
            End
            Begin VB.TextBox txtcodigo 
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
               Left            =   2400
               MaxLength       =   10
               TabIndex        =   172
               Top             =   915
               Width           =   1005
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Rango Consumo"
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
               Index           =   4
               Left            =   60
               TabIndex        =   178
               Top             =   90
               Width           =   1560
            End
            Begin VB.Label Label2 
               Caption         =   "Hasta m3"
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
               Left            =   1290
               TabIndex        =   177
               Top             =   300
               Width           =   945
            End
            Begin VB.Label Label2 
               Caption         =   "Precio"
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
            Left            =   330
            TabIndex        =   76
            Top             =   -30
            Width           =   1320
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1305
            Picture         =   "frmPOZListado.frx":1046
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   61
         Top             =   2370
         Width           =   1320
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   60
         Top             =   1980
         Width           =   1320
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   59
         Top             =   1425
         Width           =   1275
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   58
         Top             =   1020
         Width           =   1275
      End
      Begin VB.CommandButton CmdAceptarRecCons 
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
         Left            =   4350
         TabIndex        =   66
         Top             =   5610
         Width           =   1095
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
         Left            =   5505
         TabIndex        =   67
         Top             =   5595
         Width           =   1095
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
         Picture         =   "frmPOZListado.frx":10D1
         Top             =   2370
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1530
         Picture         =   "frmPOZListado.frx":115C
         Top             =   1980
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
         Index           =   22
         Left            =   570
         TabIndex        =   74
         Top             =   1800
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
         Index           =   21
         Left            =   840
         TabIndex        =   73
         Top             =   2370
         Width           =   600
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
         Index           =   20
         Left            =   840
         TabIndex        =   72
         Top             =   2010
         Width           =   645
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
         TabIndex        =   71
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hidrante"
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
         Index           =   16
         Left            =   540
         TabIndex        =   70
         Top             =   870
         Width           =   825
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
         Index           =   15
         Left            =   810
         TabIndex        =   69
         Top             =   1470
         Width           =   600
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
         Index           =   14
         Left            =   810
         TabIndex        =   68
         Top             =   1110
         Width           =   645
      End
   End
   Begin VB.Frame FrameAsignacionPrecios 
      Height          =   5025
      Left            =   0
      TabIndex        =   320
      Top             =   0
      Width           =   7395
      Begin VB.TextBox txtcodigo 
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
         Index           =   120
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   328
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
         Top             =   3390
         Width           =   1245
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   334
         Text            =   "Text5"
         Top             =   2970
         Width           =   1245
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   71
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   327
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
         Top             =   2550
         Width           =   1245
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   70
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   326
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
         Top             =   2130
         Width           =   1245
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
         Index           =   96
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   332
         Text            =   "Text5"
         Top             =   1650
         Width           =   3735
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   96
         Left            =   2700
         MaxLength       =   4
         TabIndex        =   325
         Tag             =   "Zona|N|N|1|9999|rcampos|codzonas|0000||"
         Top             =   1650
         Width           =   585
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
         Index           =   95
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   331
         Text            =   "Text5"
         Top             =   1290
         Width           =   3735
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   95
         Left            =   2700
         MaxLength       =   4
         TabIndex        =   324
         Tag             =   "Zona|N|N|1|9999|rcampos|codzonas|0000||"
         Top             =   1290
         Width           =   585
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
         Left            =   5910
         TabIndex        =   330
         Top             =   4290
         Width           =   1095
      End
      Begin VB.CommandButton CmdAcepAsigPrec 
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
         Left            =   4770
         TabIndex        =   329
         Top             =   4290
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Riego a Manta"
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
         Index           =   104
         Left            =   660
         TabIndex        =   483
         Top             =   3420
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "€/hanegada"
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
         Index           =   103
         Left            =   3960
         TabIndex        =   482
         Top             =   3420
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "€/hanegada"
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
         Index           =   77
         Left            =   3960
         TabIndex        =   339
         Top             =   2160
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "€/hanegada"
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
         Index           =   76
         Left            =   3960
         TabIndex        =   338
         Top             =   2580
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL "
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
         Index           =   75
         Left            =   660
         TabIndex        =   337
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Talla Ordinaria"
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
         Index           =   62
         Left            =   660
         TabIndex        =   336
         Top             =   2580
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Amortizacion Canal"
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
         Index           =   61
         Left            =   660
         TabIndex        =   335
         Top             =   2160
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Zonas"
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
         Index           =   89
         Left            =   660
         TabIndex        =   333
         Top             =   1050
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   2370
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Zona"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   2370
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Zona"
         Top             =   1320
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
         Index           =   91
         Left            =   1230
         TabIndex        =   323
         Top             =   1320
         Width           =   1035
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
         Index           =   90
         Left            =   1230
         TabIndex        =   322
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label Label14 
         Caption         =   "Asignación de precios"
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
   Begin VB.Frame FrameRecPdtesCobro 
      Height          =   6960
      Left            =   0
      TabIndex        =   392
      Top             =   -30
      Width           =   6675
      Begin VB.Frame Frame16 
         Caption         =   "Tipo de Listado"
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
         Height          =   705
         Left            =   480
         TabIndex        =   425
         Top             =   4980
         Width           =   5835
         Begin VB.OptionButton Option9 
            Caption         =   "Ambos"
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
            Left            =   4140
            TabIndex        =   428
            Top             =   345
            Width           =   1305
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Sector"
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
            Left            =   330
            TabIndex        =   427
            Top             =   345
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Braçal"
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
            Left            =   2340
            TabIndex        =   426
            Top             =   345
            Width           =   1305
         End
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   109
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   409
         Top             =   4575
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   107
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   405
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1320
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   106
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   404
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Text            =   "1234567890"
         Top             =   2415
         Width           =   1320
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
         Index           =   12
         Left            =   5310
         TabIndex        =   411
         Top             =   6435
         Width           =   1095
      End
      Begin VB.CommandButton CmdAcepRecPdtesCob 
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
         Left            =   4140
         TabIndex        =   410
         Top             =   6450
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   105
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   403
         Top             =   1605
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   104
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   402
         Text            =   "000000"
         Top             =   1230
         Width           =   830
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
         Height          =   360
         Index           =   103
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   407
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3645
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   102
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   406
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3270
         Width           =   1050
      End
      Begin VB.Frame Frame19 
         Caption         =   "Agrupado por"
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
         Height          =   705
         Left            =   480
         TabIndex        =   397
         Top             =   5700
         Width           =   3825
         Begin VB.OptionButton Option6 
            Caption         =   "Socio"
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
            Left            =   2325
            TabIndex        =   399
            Top             =   345
            Width           =   1305
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Sector/Braçal"
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
            Left            =   330
            TabIndex        =   398
            Top             =   300
            Value           =   -1  'True
            Width           =   1710
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Tipo Pago"
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
         ForeColor       =   &H00972E0B&
         Height          =   1215
         Left            =   4110
         TabIndex        =   393
         Top             =   2160
         Visible         =   0   'False
         Width           =   2145
         Begin VB.OptionButton Option1 
            Caption         =   "Contado"
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
            Index           =   10
            Left            =   420
            TabIndex        =   396
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Efecto"
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
            Index           =   9
            Left            =   420
            TabIndex        =   395
            Top             =   570
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
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
            Index           =   8
            Left            =   420
            TabIndex        =   394
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Label Label4 
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
         Index           =   34
         Left            =   705
         TabIndex        =   424
         Top             =   4230
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Index           =   33
         Left            =   705
         TabIndex        =   423
         Top             =   4590
         Width           =   555
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   1470
         MouseIcon       =   "frmPOZListado.frx":11E7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar braçal"
         Top             =   4590
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1470
         MouseIcon       =   "frmPOZListado.frx":1339
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar braçal"
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
            Name            =   "Verdana"
            Size            =   9.75
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
         Left            =   690
         TabIndex        =   418
         Top             =   2415
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Index           =   30
         Left            =   690
         TabIndex        =   417
         Top             =   2775
         Width           =   555
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1470
         Picture         =   "frmPOZListado.frx":148B
         ToolTipText     =   "Buscar fecha"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1470
         Picture         =   "frmPOZListado.frx":1516
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Left            =   720
         TabIndex        =   416
         Top             =   1245
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Index           =   26
         Left            =   735
         TabIndex        =   415
         Top             =   1620
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   25
         Left            =   510
         TabIndex        =   414
         Top             =   960
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":15A1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":16F3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Braçal"
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
         Height          =   255
         Index           =   24
         Left            =   525
         TabIndex        =   413
         Top             =   3945
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Sector"
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
         Height          =   255
         Index           =   23
         Left            =   510
         TabIndex        =   412
         Top             =   3180
         Width           =   1815
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
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   111
         Top             =   6270
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Artículos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2265
         Left            =   270
         TabIndex        =   108
         Top             =   3900
         Width           =   7815
         Begin VB.TextBox txtcodigo 
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
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   96
            Text            =   "frmPOZListado.frx":1845
            Top             =   1755
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   97
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   1755
            Width           =   1245
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   94
            Text            =   "frmPOZListado.frx":188E
            Top             =   1350
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   95
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   1350
            Width           =   1245
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   92
            Text            =   "frmPOZListado.frx":18D7
            Top             =   945
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   93
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   945
            Width           =   1245
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   90
            Text            =   "frmPOZListado.frx":1920
            Top             =   540
            Width           =   6105
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   6420
            MaxLength       =   10
            TabIndex        =   91
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   540
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
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
            Height          =   240
            Index           =   33
            Left            =   6420
            TabIndex        =   110
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
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
            Height          =   240
            Index           =   32
            Left            =   240
            TabIndex        =   109
            Top             =   270
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mano Obra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1095
         Left            =   300
         TabIndex        =   105
         Top             =   2670
         Width           =   7785
         Begin VB.TextBox txtcodigo 
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
            Left            =   6390
            MaxLength       =   10
            TabIndex        =   89
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   540
            Width           =   1245
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   210
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   88
            Text            =   "frmPOZListado.frx":1969
            Top             =   540
            Width           =   6105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
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
            Height          =   240
            Index           =   31
            Left            =   6390
            TabIndex        =   107
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
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
            Height          =   240
            Index           =   8
            Left            =   210
            TabIndex        =   106
            Top             =   270
            Width           =   945
         End
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
         Left            =   6945
         TabIndex        =   100
         Top             =   7095
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptarRecCont 
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
         Left            =   5760
         TabIndex        =   98
         Top             =   7110
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   86
         Top             =   1440
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
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
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   87
            Top             =   240
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
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
            Index           =   5
            Left            =   4290
            TabIndex        =   83
            Top             =   60
            Width           =   2280
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Recibo"
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
            Index           =   4
            Left            =   4290
            TabIndex        =   82
            Top             =   420
            Width           =   2310
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recibo"
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
            Index           =   11
            Left            =   330
            TabIndex        =   84
            Top             =   -30
            Width           =   1320
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   1350
            Picture         =   "frmPOZListado.frx":19B2
            Top             =   240
            Width           =   240
         End
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
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe  Recibo"
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
         Index           =   12
         Left            =   4890
         TabIndex        =   112
         Top             =   6300
         Width           =   1560
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
         Index           =   30
         Left            =   810
         TabIndex        =   104
         Top             =   1110
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
         Index           =   23
         Left            =   810
         TabIndex        =   103
         Top             =   1470
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   13
         Left            =   540
         TabIndex        =   102
         Top             =   870
         Width           =   540
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
         TabIndex        =   101
         Top             =   300
         Width           =   5925
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1590
         MouseIcon       =   "frmPOZListado.frx":1A3D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1590
         MouseIcon       =   "frmPOZListado.frx":1B8F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame FrameReciboMantenimiento 
      Height          =   7505
      Left            =   30
      TabIndex        =   26
      Top             =   0
      Width           =   6945
      Begin VB.TextBox txtcodigo 
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
         Index           =   65
         Left            =   4215
         MaxLength       =   10
         TabIndex        =   36
         Top             =   3750
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   64
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   35
         Top             =   3720
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   63
         Left            =   4215
         MaxLength       =   10
         TabIndex        =   30
         Top             =   1995
         Width           =   1080
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   62
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   29
         Text            =   "      "
         Top             =   1995
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   60
         Left            =   4215
         MaxLength       =   25
         TabIndex        =   34
         Text            =   "      "
         Top             =   3135
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   59
         Left            =   1845
         MaxLength       =   25
         TabIndex        =   33
         Top             =   3135
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   58
         Left            =   4215
         MaxLength       =   6
         TabIndex        =   32
         Text            =   "      "
         Top             =   2550
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   57
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   31
         Top             =   2550
         Width           =   960
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
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text5"
         Top             =   1485
         Width           =   3375
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   210
         TabIndex        =   49
         Top             =   4215
         Width           =   6375
         Begin VB.TextBox txtcodigo 
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
            Index           =   61
            Left            =   3990
            MaxLength       =   10
            TabIndex        =   40
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   1215
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   53
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   39
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   1215
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Recibo"
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
            Index           =   0
            Left            =   4020
            TabIndex        =   51
            Top             =   555
            Width           =   2220
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
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
            Left            =   4020
            TabIndex        =   50
            Top             =   195
            Width           =   2190
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   38
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   660
            Width           =   1335
         End
         Begin VB.TextBox txtcodigo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   9
            Left            =   1650
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
            Top             =   1620
            Width           =   4725
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   37
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   4980
            TabIndex        =   216
            Top             =   1245
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   53
            Left            =   2685
            TabIndex        =   215
            Top             =   1245
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Recargo"
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
            Left            =   3165
            TabIndex        =   208
            Top             =   1245
            Width           =   795
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   1365
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bonificación"
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
            Index           =   37
            Left            =   330
            TabIndex        =   201
            Top             =   1020
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
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
            Left            =   330
            TabIndex        =   54
            Top             =   1530
            Width           =   945
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   1350
            Picture         =   "frmPOZListado.frx":1CE1
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recibo"
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
            Index           =   10
            Left            =   330
            TabIndex        =   53
            Top             =   -30
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Euros/Acción"
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
            Left            =   330
            TabIndex        =   52
            Top             =   660
            Width           =   1290
         End
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   27
         Top             =   1080
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   28
         Top             =   1485
         Width           =   960
      End
      Begin VB.CommandButton CmdAceptarRecMto 
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
         TabIndex        =   42
         Top             =   6810
         Width           =   1095
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
         Left            =   5505
         TabIndex        =   43
         Top             =   6795
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   255
         Left            =   540
         TabIndex        =   44
         Top             =   6450
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
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
         Index           =   52
         Left            =   840
         TabIndex        =   214
         Top             =   3780
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
         Index           =   51
         Left            =   3270
         TabIndex        =   213
         Top             =   3750
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alta"
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
         Index           =   50
         Left            =   540
         TabIndex        =   212
         Top             =   3510
         Width           =   1065
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   3945
         Picture         =   "frmPOZListado.frx":1D6C
         Top             =   3750
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   1560
         Picture         =   "frmPOZListado.frx":1DF7
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hidrante"
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
         Index           =   49
         Left            =   540
         TabIndex        =   211
         Top             =   1785
         Width           =   825
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
         Left            =   810
         TabIndex        =   210
         Top             =   2025
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
         Left            =   3270
         TabIndex        =   209
         Top             =   2025
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
         Index           =   45
         Left            =   3270
         TabIndex        =   207
         Top             =   3165
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
         Index           =   44
         Left            =   810
         TabIndex        =   206
         Top             =   3165
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
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
         Left            =   540
         TabIndex        =   205
         Top             =   2925
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
         Index           =   42
         Left            =   3270
         TabIndex        =   204
         Top             =   2580
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
         Index           =   41
         Left            =   810
         TabIndex        =   203
         Top             =   2580
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Polígono"
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
         TabIndex        =   202
         Top             =   2340
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":1E82
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1560
         MouseIcon       =   "frmPOZListado.frx":1FD4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1485
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
         Index           =   19
         Left            =   540
         TabIndex        =   47
         Top             =   870
         Width           =   540
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
         Index           =   18
         Left            =   810
         TabIndex        =   46
         Top             =   1470
         Width           =   645
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
         Index           =   17
         Left            =   810
         TabIndex        =   45
         Top             =   1110
         Width           =   690
      End
   End
   Begin VB.Frame FrameReciboTalla 
      Height          =   5085
      Left            =   0
      TabIndex        =   238
      Top             =   30
      Width           =   7530
      Begin VB.Frame FrameCuota 
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   315
         TabIndex        =   301
         Top             =   2925
         Width           =   6990
         Begin VB.TextBox txtcodigo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   76
            Left            =   1590
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   302
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
            Top             =   150
            Width           =   5130
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   1620
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
            Index           =   81
            Left            =   270
            TabIndex        =   304
            Top             =   660
            Width           =   1275
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
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
            Index           =   67
            Left            =   270
            TabIndex        =   303
            Top             =   210
            Width           =   945
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen"
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
         Index           =   7
         Left            =   4410
         TabIndex        =   259
         Top             =   1950
         Width           =   2190
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Recibo"
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
         Index           =   6
         Left            =   4410
         TabIndex        =   258
         Top             =   2310
         Width           =   2220
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
         Left            =   5925
         TabIndex        =   247
         Top             =   4455
         Width           =   1095
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
         Left            =   4725
         TabIndex        =   244
         Top             =   4470
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   75
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   242
         Top             =   1500
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   74
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   241
         Top             =   1110
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   73
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   243
         Text            =   "0000000000"
         Top             =   2160
         Width           =   1350
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
         Index           =   74
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   240
         Text            =   "Text5"
         Top             =   1110
         Width           =   4080
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
         Index           =   75
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   239
         Text            =   "Text5"
         Top             =   1500
         Width           =   4080
      End
      Begin MSComctlLib.ProgressBar pb4 
         Height          =   255
         Left            =   480
         TabIndex        =   260
         Top             =   4080
         Width           =   6570
         _ExtentX        =   11589
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
         Top             =   2670
         Width           =   6705
         Begin VB.CheckBox Check1 
            Caption         =   "Sólo efectos"
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
            Index           =   8
            Left            =   5115
            TabIndex        =   261
            Top             =   420
            Width           =   1875
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   78
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   245
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   77
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   246
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bonificación"
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
            Index           =   71
            Left            =   150
            TabIndex        =   257
            Top             =   390
            Width           =   1170
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   1
            Left            =   1365
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
            Index           =   70
            Left            =   3030
            TabIndex        =   256
            Top             =   390
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   69
            Left            =   2640
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
            Left            =   4890
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
         Width           =   3825
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
         Index           =   66
         Left            =   930
         TabIndex        =   252
         Top             =   1110
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
         Index           =   65
         Left            =   930
         TabIndex        =   251
         Top             =   1470
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   64
         Left            =   540
         TabIndex        =   250
         Top             =   870
         Width           =   540
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
         TabIndex        =   249
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
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
         Index           =   63
         Left            =   570
         TabIndex        =   248
         Top             =   1890
         Width           =   1320
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   1590
         Picture         =   "frmPOZListado.frx":2126
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1635
         MouseIcon       =   "frmPOZListado.frx":21B1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1635
         MouseIcon       =   "frmPOZListado.frx":2303
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1515
         Width           =   240
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
         Index           =   11
         Left            =   5460
         TabIndex        =   391
         Top             =   4395
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptarRecManta 
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
         Left            =   4290
         TabIndex        =   389
         Top             =   4410
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   115
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   383
         Top             =   1200
         Width           =   960
      End
      Begin VB.Frame Frame17 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   180
         TabIndex        =   377
         Top             =   1710
         Width           =   6375
         Begin VB.TextBox txtcodigo 
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
            Index           =   114
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   384
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox txtcodigo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   113
            Left            =   1650
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   387
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
            Top             =   1095
            Width           =   4725
         End
         Begin VB.TextBox txtcodigo 
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
            Height          =   360
            Index           =   112
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   385
            Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|###,##0.00||"
            Top             =   660
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
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
            Index           =   9
            Left            =   4050
            TabIndex        =   379
            Top             =   120
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Ticket"
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
            Index           =   10
            Left            =   4050
            TabIndex        =   378
            Top             =   480
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Euros/Acción"
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
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   109
            Left            =   330
            TabIndex        =   382
            Top             =   660
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ticket"
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
            Index           =   108
            Left            =   330
            TabIndex        =   381
            Top             =   -30
            Width           =   1290
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   16
            Left            =   1350
            Picture         =   "frmPOZListado.frx":2455
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
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
            Index           =   107
            Left            =   330
            TabIndex        =   380
            Top             =   1050
            Width           =   945
         End
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
         Index           =   122
         Left            =   540
         TabIndex        =   390
         Top             =   1170
         Width           =   540
      End
      Begin VB.Label Label17 
         Caption         =   "Generación Tickets Consumo a Manta"
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
         MouseIcon       =   "frmPOZListado.frx":24E0
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
         Left            =   450
         TabIndex        =   481
         Top             =   5100
         Width           =   2565
      End
      Begin VB.Frame Frame24 
         Caption         =   "Ordenado por"
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
         Height          =   750
         Left            =   450
         TabIndex        =   466
         Top             =   4305
         Width           =   3690
         Begin VB.OptionButton Option14 
            Caption         =   "Hidrante"
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
            Left            =   360
            TabIndex        =   470
            Top             =   360
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Socio"
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
            Left            =   1890
            TabIndex        =   468
            Top             =   360
            Width           =   1305
         End
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   127
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   465
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3795
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   126
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   464
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3375
         Width           =   1350
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
         Index           =   125
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   459
         Top             =   1620
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   124
         Left            =   1770
         MaxLength       =   6
         TabIndex        =   458
         Top             =   1230
         Width           =   830
      End
      Begin VB.CommandButton CmdAcepRecConsPdtes 
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
         Left            =   4095
         TabIndex        =   467
         Top             =   5190
         Width           =   1095
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
         Left            =   5265
         TabIndex        =   469
         Top             =   5175
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   123
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   463
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   122
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   461
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2400
         Width           =   1350
      End
      Begin VB.Label Label4 
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
         Index           =   45
         Left            =   720
         TabIndex        =   480
         Top             =   3795
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Index           =   44
         Left            =   720
         TabIndex        =   479
         Top             =   3435
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Hidrante"
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
         Height          =   255
         Index           =   53
         Left            =   510
         TabIndex        =   478
         Top             =   3165
         Width           =   825
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   32
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":2632
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   31
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":2784
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Left            =   510
         TabIndex        =   477
         Top             =   1005
         Width           =   540
      End
      Begin VB.Label Label4 
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
         Index           =   50
         Left            =   780
         TabIndex        =   476
         Top             =   1620
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Index           =   49
         Left            =   765
         TabIndex        =   475
         Top             =   1245
         Width           =   645
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   1470
         Picture         =   "frmPOZListado.frx":28D6
         ToolTipText     =   "Buscar fecha"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   1470
         Picture         =   "frmPOZListado.frx":2961
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   48
         Left            =   735
         TabIndex        =   474
         Top             =   2775
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Index           =   47
         Left            =   735
         TabIndex        =   473
         Top             =   2415
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
   Begin VB.Frame FrameInfMantaFechaRiego 
      Height          =   5910
      Left            =   0
      TabIndex        =   429
      Top             =   0
      Width           =   6675
      Begin VB.Frame Frame21 
         Caption         =   "Agrupado por"
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
         Height          =   705
         Left            =   480
         TabIndex        =   454
         Top             =   4440
         Width           =   5895
         Begin VB.OptionButton Option1 
            Caption         =   "Fecha de Riego"
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
            Index           =   14
            Left            =   765
            TabIndex        =   456
            Top             =   300
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Fecha de Pago"
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
            Index           =   15
            Left            =   3525
            TabIndex        =   455
            Top             =   300
            Width           =   1815
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Tipo Pago"
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
         Height          =   1215
         Left            =   4200
         TabIndex        =   450
         Top             =   2160
         Width           =   2145
         Begin VB.OptionButton Option1 
            Caption         =   "Contado"
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
            Index           =   13
            Left            =   420
            TabIndex        =   453
            Top             =   300
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Efecto"
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
            Index           =   12
            Left            =   420
            TabIndex        =   452
            Top             =   570
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
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
            Index           =   11
            Left            =   420
            TabIndex        =   451
            Top             =   840
            Width           =   1155
         End
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   111
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   437
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3945
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   110
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   436
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3570
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   119
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   435
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2790
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   118
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   434
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1365
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
         Index           =   13
         Left            =   5265
         TabIndex        =   439
         Top             =   5295
         Width           =   1095
      End
      Begin VB.CommandButton CmdAcepLisTicFecRiego 
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
         Left            =   4095
         TabIndex        =   438
         Top             =   5310
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   117
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   433
         Top             =   1605
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
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
         Picture         =   "frmPOZListado.frx":29EC
         ToolTipText     =   "Buscar fecha"
         Top             =   3945
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   43
         Left            =   795
         TabIndex        =   449
         Top             =   3930
         Width           =   555
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   1500
         Picture         =   "frmPOZListado.frx":2A77
         ToolTipText     =   "Buscar fecha"
         Top             =   3555
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   36
         Left            =   780
         TabIndex        =   448
         Top             =   3570
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Pago"
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
            Name            =   "Verdana"
            Size            =   9.75
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
         Index           =   41
         Left            =   780
         TabIndex        =   444
         Top             =   2415
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Index           =   40
         Left            =   780
         TabIndex        =   443
         Top             =   2775
         Width           =   555
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   1500
         Picture         =   "frmPOZListado.frx":2B02
         ToolTipText     =   "Buscar fecha"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   1500
         Picture         =   "frmPOZListado.frx":2B8D
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   39
         Left            =   810
         TabIndex        =   442
         Top             =   1245
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Index           =   38
         Left            =   825
         TabIndex        =   441
         Top             =   1620
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   37
         Left            =   510
         TabIndex        =   440
         Top             =   1005
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":2C18
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":2D6A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
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
            Name            =   "Verdana"
            Size            =   9.75
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
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   300
            TabIndex        =   25
            Top             =   660
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Contador"
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
            Index           =   0
            Left            =   300
            TabIndex        =   24
            Top             =   330
            Width           =   1725
         End
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1665
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1260
         Width           =   1350
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
         Left            =   3465
         TabIndex        =   18
         Top             =   3195
         Width           =   1095
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
         Index           =   0
         Left            =   4680
         TabIndex        =   20
         Top             =   3210
         Width           =   1095
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
         Index           =   2
         Left            =   600
         TabIndex        =   21
         Top             =   1080
         Width           =   825
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
         Left            =   960
         TabIndex        =   19
         Top             =   1680
         Width           =   600
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
         Index           =   0
         Left            =   960
         TabIndex        =   17
         Top             =   1320
         Width           =   645
      End
   End
   Begin VB.Frame FrameEtiquetasContadores 
      Height          =   3885
      Left            =   0
      TabIndex        =   157
      Top             =   0
      Width           =   6945
      Begin VB.TextBox txtcodigo 
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
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   160
         Top             =   1920
         Width           =   5175
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   44
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   162
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|##,###,##0||"
         Top             =   2850
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   45
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   158
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   1050
         Width           =   5175
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   159
         Top             =   1470
         Width           =   5175
      End
      Begin VB.CommandButton CmdAceptarEtiq 
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
         Left            =   4320
         TabIndex        =   164
         Top             =   3210
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelEtiq 
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
         Left            =   5520
         TabIndex        =   166
         Top             =   3195
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Línea 3"
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
         Index           =   34
         Left            =   570
         TabIndex        =   168
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Línea 2"
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
         Index           =   39
         Left            =   570
         TabIndex        =   167
         Top             =   1500
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Nro.Etiquetas"
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
         Height          =   195
         Index           =   38
         Left            =   570
         TabIndex        =   165
         Top             =   2550
         Width           =   1590
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
         Caption         =   "Línea 1"
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
         Left            =   570
         TabIndex        =   161
         Top             =   1080
         Width           =   705
      End
   End
   Begin VB.Frame FrameRectificacion 
      Height          =   4680
      Left            =   0
      TabIndex        =   183
      Top             =   0
      Width           =   7035
      Begin VB.Frame Frame9 
         Caption         =   "Datos para Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1995
         Left            =   240
         TabIndex        =   187
         Top             =   870
         Width           =   6675
         Begin VB.TextBox txtcodigo 
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
            Index           =   55
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   191
            Text            =   "0000000000"
            Top             =   1350
            Width           =   1305
         End
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
            Index           =   2
            Left            =   4485
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   196
            Top             =   450
            Width           =   2070
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   52
            Left            =   1440
            MaxLength       =   7
            TabIndex        =   189
            Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
            Top             =   450
            Width           =   1305
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   56
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   190
            Top             =   900
            Width           =   1065
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
            Index           =   56
            Left            =   2550
            Locked          =   -1  'True
            TabIndex        =   188
            Text            =   "Text5"
            Top             =   900
            Width           =   3990
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hidrante"
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
            Index           =   7
            Left            =   120
            TabIndex        =   200
            Top             =   1350
            Width           =   825
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nro.Factura"
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
            Index           =   21
            Left            =   120
            TabIndex        =   199
            Top             =   465
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Socio"
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
            Left            =   120
            TabIndex        =   198
            Top             =   900
            Width           =   540
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   1140
            MouseIcon       =   "frmPOZListado.frx":2EBC
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   900
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Factura"
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
         Index           =   51
         Left            =   1695
         MaxLength       =   6
         TabIndex        =   192
         Top             =   3060
         Width           =   1290
      End
      Begin VB.CommandButton CmdAceptarRectif 
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
         Left            =   4620
         TabIndex        =   194
         Top             =   3915
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelRectif 
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
         Left            =   5790
         TabIndex        =   195
         Top             =   3915
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   54
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   193
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3480
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lectura"
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
         Left            =   375
         TabIndex        =   186
         Top             =   3060
         Width           =   750
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   1425
         Picture         =   "frmPOZListado.frx":300E
         ToolTipText     =   "Buscar fecha"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "F.Factura"
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
         Height          =   255
         Index           =   22
         Left            =   375
         TabIndex        =   185
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
         Left            =   375
         TabIndex        =   184
         Top             =   360
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
         Left            =   5430
         TabIndex        =   266
         Top             =   2580
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
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
         TabIndex        =   264
         Top             =   2580
         Width           =   1095
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
   Begin VB.Frame FrameImporLecturas 
      Height          =   4725
      Left            =   0
      TabIndex        =   485
      Top             =   0
      Width           =   6675
      Begin VB.TextBox txtcodigo 
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
         Index           =   131
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   490
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2055
         Width           =   1230
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
         Index           =   15
         Left            =   5220
         TabIndex        =   494
         Top             =   4035
         Width           =   1095
      End
      Begin VB.CommandButton CmdAcepImportar 
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
         Left            =   4050
         TabIndex        =   492
         Top             =   4050
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   130
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   489
         Top             =   1560
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   129
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   488
         Top             =   1095
         Width           =   830
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
         Index           =   130
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   487
         Text            =   "Text5"
         Top             =   1560
         Width           =   3675
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
         Index           =   129
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   486
         Text            =   "Text5"
         Top             =   1095
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1755
         MaxLength       =   100
         TabIndex        =   491
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2550
         Width           =   4560
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   33
         Left            =   1485
         MouseIcon       =   "frmPOZListado.frx":3099
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar fichero"
         Top             =   2565
         Width           =   240
      End
      Begin VB.Label lblProgres 
         Height          =   405
         Index           =   0
         Left            =   180
         TabIndex        =   500
         Top             =   3015
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   499
         Top             =   3540
         Width           =   6195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
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
         Index           =   57
         Left            =   315
         TabIndex        =   498
         Top             =   1575
         Width           =   945
      End
      Begin VB.Label Label21 
         Caption         =   "Importacion de Lecturas"
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
         TabIndex        =   497
         Top             =   315
         Width           =   3900
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   61
         Left            =   315
         TabIndex        =   496
         Top             =   2070
         Width           =   780
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   26
         Left            =   1485
         Picture         =   "frmPOZListado.frx":31EB
         ToolTipText     =   "Buscar fecha"
         Top             =   2055
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comunidad"
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
         Index           =   56
         Left            =   330
         TabIndex        =   495
         Top             =   1095
         Width           =   1065
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   30
         Left            =   1485
         MouseIcon       =   "frmPOZListado.frx":3276
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   1575
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   29
         Left            =   1500
         MouseIcon       =   "frmPOZListado.frx":33C8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar comunidad"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fichero"
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
         Height          =   255
         Index           =   54
         Left            =   330
         TabIndex        =   493
         Top             =   2595
         Width           =   780
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
    
    ' 13.-
    ' 14.- Asignacion de Precios de Talla
    
    ' 15.- Listado de Diferencias con Indefa
    ' 16.- Listado de cuentas bancarias de socios erróneas
    
    ' 17.- Generacion de recibos a manta
    
    ' 18.- Informe de recibos pendientes de cobro por braçal y por sector
    ' 19.- Informe de recibos de riego a manta por fecha de riego
    ' 20.- Informe de recibos de consumo pendientes de cobro
    
    ' 21.- Importacion de lecturas de Monasterios
    ' 22.- Exportacion de lecturas de Monasterios
    
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
Private WithEvents frmZon  As frmManZonas 'mantenimiento de zonas
Attribute frmZon.VB_VarHelpID = -1
Private WithEvents frmPoz As frmPOZPozos  ' comunidades en monasterios
Attribute frmPoz.VB_VarHelpID = -1
Private WithEvents frmCon As frmBasico  ' conceptos de monasterios
Attribute frmCon.VB_VarHelpID = -1
 
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

Private WithEvents frmMensSoc As frmMensajes ' seleccion de socios en los recibos de talla de escalona
Attribute frmMensSoc.VB_VarHelpID = -1

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
Dim indFrame As Single 'nº de frame en el que estamos
 
Dim IndRptReport As Integer
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
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
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check2_Click()
    Me.Frame7.Enabled = (Check2.Value = 1)
End Sub

Private Sub CmdAcepAsigPrec_Click()
Dim Sql As String

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
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H ZONA
    cDesde = Trim(txtCodigo(95).Text)
    cHasta = Trim(txtCodigo(96).Text)
    nDesde = txtNombre(95).Text
    nHasta = txtNombre(96).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codzonas}"
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

    vSQL = "update rzonas set precio1 =  " & DBSet(txtCodigo(70).Text, "N")
    vSQL = vSQL & ", precio2 = " & DBSet(txtCodigo(71).Text, "N")
    vSQL = vSQL & ", preciomanta = " & DBSet(txtCodigo(120).Text, "N")
    If cadSelect <> "" Then vSQL = vSQL & " where " & cadSelect

    conn.Execute vSQL
    
       
    conn.CommitTrans
    ProcesarCambios = True
    Exit Function
    
eProcesarCambios:
    conn.RollbackTrans
    MuestraError Err.Number, "Procesar Cambios", Err.Description
End Function



Private Function ProcesarFichero(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim Sql1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim NomFic As String

    ProcesarFichero = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    B = True
    While Not EOF(NF)
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        B = InsertarLinea(cad)
        
        If B = False Then
            ProcesarFichero = False
            Exit Function
        End If
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        B = InsertarLinea(cad)

        If B = False Then
            ProcesarFichero = False
            Exit Function
        End If
    End If
    
    ProcesarFichero = B
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

End Function


Private Function InsertarLinea(cad As String) As Boolean
Dim Sql As String
Dim Comunidad As String
Dim Calle As String
Dim Propiedad As String
Dim Tipo As String
Dim Contador As String
Dim observa As String
Dim Concepto As String
Dim Lectura As String

    On Error GoTo EInsertarLinea

    InsertarLinea = False
    
    
    Tipo = Mid(cad, 1, 2)
    
    If Tipo <> "40" And Tipo <> "60" Then
        InsertarLinea = True
        Exit Function
    End If
    
    Comunidad = Mid(cad, 3, 3)
    Calle = Mid(cad, 6, 4)
    Propiedad = Mid(cad, 10, 4)


    Contador = Format(Comunidad, "00") & Format(Calle, "0000") & Format(Propiedad, "0000")


    Select Case Tipo
        Case "40"
            observa = Mid(cad, 14, 27)
            
            Sql = "update rcampos set observac = " & DBSet(observa, "T") & " where codcampo = " & DBSet(Propiedad, "N")
            
            conn.Execute Sql
            
        Case "60"
            Concepto = Mid(cad, 14, 2)
            Lectura = Mid(cad, 16, 7)
            
            
            Sql = "update rpozos set fech_ant = " & DBSet(txtCodigo(131), "F") & ", "
            Sql = Sql & " lect_ant = " & DBSet(Lectura, "N")
            Sql = Sql & ", lect_act = " & ValorNulo
            Sql = Sql & ", fech_act = " & ValorNulo
            Sql = Sql & ", consumo = " & ValorNulo
            Sql = Sql & " where hidrante = " & DBSet(Contador, "T")
            
            conn.Execute Sql
            
        Case "90"
        
    End Select
    
    InsertarLinea = True
EInsertarLinea:
    If Err.Number <> 0 Then
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function






Private Function ProcesarFichero2(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim Sql1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean

    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO VENTAS.TXT

    B = True

    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        B = ComprobarRegistro(cad)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        B = ComprobarRegistro(cad)
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFichero2 = B
    Exit Function

eProcesarFichero2:
    ProcesarFichero2 = False
End Function


Private Function ComprobarRegistro(cad As String) As Boolean
Dim Sql As String
Dim Tipo As String
Dim Comunidad As String
Dim Calle As String
Dim Propiedad As String
Dim observa As String
Dim Concepto As String
Dim Lectura As String
Dim Aux As String
Dim Contador As String

Dim Mens As String


    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True


    Tipo = Mid(cad, 1, 2)
    
    If Tipo <> "40" And Tipo <> "60" Then Exit Function
    
    Comunidad = Mid(cad, 3, 3)
    Calle = Mid(cad, 6, 4)
    Propiedad = Mid(cad, 10, 4)

    
    Select Case Tipo
        Case "40"
            observa = Mid(cad, 14, 27)
        Case "60"
            Concepto = Mid(cad, 14, 2)
            Lectura = Mid(cad, 16, 7)
        Case "90"
        
    End Select
    
    Contador = Format(Comunidad, "00") & Format(Calle, "0000") & Format(Propiedad, "0000")
    

    'Comprobar comunidad
    Aux = DevuelveDesdeBD("codpozo", "rtipopozos", "codpozo", Comunidad, "N")
    If Aux = "" Then
        Mens = "Comunidad no existe"
        Sql = "insert into tmpinformes (codusu, importe1, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Comunidad, "N") & "," & DBSet(Contador, "T") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute Sql
    End If
    
    'Comprobar calles
    Aux = DevuelveDesdeBD("codparti", "rpartida", "codparti", Calle, "N")
    If Aux = "" Then
        Mens = "Calle no existe"
        Sql = "insert into tmpinformes (codusu, importe1, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Calle, "N") & "," & DBSet(Contador, "T") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute Sql
    End If
    
    'Comprobar propiedad
    Aux = DevuelveDesdeBD("codcampo", "rcampos", "codcampo", Propiedad, "N")
    If Aux = "" Then
        Mens = "Propiedad no existe"
        Sql = "insert into tmpinformes (codusu, importe1, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Propiedad, "N") & "," & DBSet(Contador, "T") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute Sql
    End If
    
    'Contador no existe
    Aux = DevuelveDesdeBD("codcampo", "rpozos", "hidrante", Contador, "T")
    If Aux = "" Then
        Mens = "Contador no existe"
        Sql = "insert into tmpinformes (codusu, importe1, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(0, "N") & "," & DBSet(Contador, "T") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute Sql
    End If
    
    
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function


Private Sub InicializarTabla()
Dim Sql As String
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    conn.Execute Sql
End Sub


Private Sub CmdAcepExportar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim CodTipom As String
Dim Sql As String
Dim B As Boolean
Dim NFich As Integer
Dim ComunidadAnt As String
Dim Consumo As Long
Dim Importe As Currency
Dim TImporte As Currency
Dim Rs As ADODB.Recordset
Dim Nregs As Long
Dim cad As String
Dim NF As Integer
Dim cImporte As String
Dim SqlRec As String
Dim numfactu As Long
Dim tipoMov As String
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim Sql2 As String



    If Not DatosOK Then Exit Sub
    
    Sql = "select rpozos.*, rtipopozos.nompozo,rtipopozos.precio1 from rpozos inner join rtipopozos on rpozos.codpozo = rtipopozos.codpozo order by rpozos.hidrante "
    
    Sql2 = "select count(*) from (" & Sql & ") aaaaaa"
    
    If TotalRegistros(Sql2) > 0 Then
    
        lblProgres(2).visible = True
        
        Me.Pb8.visible = True
        Me.Pb8.Max = TotalRegistros(Sql2)
        Me.Pb8.Value = 0
        
    
        NF = FreeFile
        Open App.Path & "\exportacion.txt" For Output As #NF
    
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        conn.BeginTrans
        
        
        tipoMov = "RCP"
        
        If Not Rs.EOF Then
            TImporte = 0
            Nregs = 0
        
            ComunidadAnt = DBLet(Rs!codpozo, "N")
            ' cabecera de comunidad
            cad = "50"
            cad = cad & Format(ComunidadAnt, "000")
            cad = cad & "00000"
            cad = cad & RellenaABlancos(DBLet(Rs!nompozo, "T"), True, 40)
            cad = cad & txtCodigo(121).Text
            cad = cad & Format(DBLet(Rs!fech_ant, "F"), "dd/mm/yyyy")
            cad = cad & Format(DBLet(Rs!fech_act, "F"), "dd/mm/yyyy")
            
            Print #NF, cad
            
            
            Set vTipoMov = New CTiposMov
            
            
        End If
        
        
        While Not Rs.EOF
' registro detalle por propietario guardado para cuando se sepa como calcular importes
'

            Consumo = DBLet(Rs!Consumo, "N")
'            Precio = DBLet(Rs!Precio1, "N")
'            Importe = Round2(Consumo * Precio, 2)
'            TImporte = TImporte + Importe
            
            lblProgres(2).Caption = "Procesando contador : " & Rs!Hidrante
            Me.Refresh
            DoEvents
            
            IncrementarProgresNew Pb8, 1
            
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", txtCodigo(121).Text, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            
            
            
            Importe = 0
            
            Nregs = Nregs + 1
            
            cad = "60"
            cad = cad & Format(Rs!codpozo, "000")
            cad = cad & Format(Rs!codcampo, "00000")
            
            ' agua fria
            If Importe >= 0 Then
                cImporte = Format(Importe, "0000000.00")
            Else
                cImporte = Format(Abs(Importe), "0000000.00")
            End If
            
            cad = cad & Replace(cImporte, ",", "")
            If Importe >= 0 Then
                cad = cad & "+"
            Else
                cad = cad & "-"
            End If
            
            ' agua caliente
            cad = cad & "000000000+"
            
            ' otro contador
            cad = cad & "000000000+"
            
            ' consumo agua fria
            cad = cad & Format(Consumo, "00000")
            ' consumo agua caliente
            cad = cad & "00000"
            ' consumo agua otro contador
            cad = cad & "00000"
            
            'observaciones
            cad = cad & Space(160)
            
            Print #NF, cad
            
            
            SqlRec = "insert into rrecibpozos (codtipom,numfactu,fecfactu,numlinea,codsocio,hidrante,baseimpo,"
            SqlRec = SqlRec & "imporiva,totalfact,consumo,lect_ant,fech_ant,lect_act,fech_act,consumo1,precio1) values ("
            SqlRec = SqlRec & "'RCP'," & DBSet(numfactu, "N") & "," & DBSet(txtCodigo(121).Text, "F") & ",1," & DBSet(Rs!Codsocio, "N") & ","
            SqlRec = SqlRec & DBSet(Rs!Hidrante, "T") & "," & DBSet(Importe, "N") & ","
            SqlRec = SqlRec & DBSet(Importe, "N") & "," & DBSet(Importe, "N") & "," & DBSet(Consumo, "N") & "," & DBSet(Rs!lect_ant, "N") & "," & DBSet(Rs!fech_ant, "F") & ","
            SqlRec = SqlRec & DBSet(Rs!lect_act, "N") & "," & DBSet(txtCodigo(121).Text, "F") & "," & DBSet(Consumo, "N") & ",0)"
            
            conn.Execute SqlRec
            
            
            SqlRec = "update rpozos set lect_ant = lect_act, fech_ant = " & DBSet(txtCodigo(121).Text, "F") & ", lect_act = " & ValorNulo & ", fech_act = " & ValorNulo & ", consumo = 0 "
            SqlRec = SqlRec & " where hidrante = " & DBSet(Rs!Hidrante, "T")
            
            conn.Execute SqlRec
            
            vTipoMov.IncrementarContador tipoMov
            
            Rs.MoveNext
        
            If Rs.EOF Then
                ' imprimimos el registro de total por comunidad
                cad = "80"
                cad = cad & Format(ComunidadAnt, "000")
                cad = cad & "00000"
                ' agua fria
                If TImporte >= 0 Then
                    cImporte = Format(TImporte, "0000000.00")
                Else
                    cImporte = Format(Abs(TImporte), "0000000.00")
                End If
                
                cad = cad & Replace(cImporte, ",", "")
                If Importe >= 0 Then
                    cad = cad & "+"
                Else
                    cad = cad & "-"
                End If
                ' agua caliente
                cad = cad & "000000000+"
                ' otro contador
                cad = cad & "000000000+"
                ' nuero de registros
                cad = cad & Format(Nregs, "00000")
                
                Print #NF, cad
            
            Else
        
                If ComunidadAnt <> DBLet(Rs!codpozo, "N") Then
                    ' imprimimos el registro de total por comunidad
                    cad = "80"
                    cad = cad & Format(ComunidadAnt, "000")
                    cad = cad & "00000"
                    ' agua fria
                    If TImporte >= 0 Then
                        cImporte = Format(TImporte, "0000000.00")
                    Else
                        cImporte = Format(Abs(TImporte), "0000000.00")
                    End If
                    
                    cad = cad & Replace(cImporte, ",", "")
                    If Importe >= 0 Then
                        cad = cad & "+"
                    Else
                        cad = cad & "-"
                    End If
                    ' agua caliente
                    cad = cad & "000000000+"
                    ' otro contador
                    cad = cad & "000000000+"
                    ' nuero de registros
                    cad = cad & Format(Nregs, "00000")
                    
                    Print #NF, cad
                
                    If Not Rs.EOF Then
                        ComunidadAnt = DBLet(Rs!codpozo, "N")
                
                        TImporte = 0
                        Nregs = 0
                        
                        ' cabecera de comunidad
                        cad = "50"
                        cad = cad & Format(Rs!codpozo, "000")
                        cad = cad & "00000"
                        cad = cad & RellenaABlancos(DBLet(Rs!nompozo, "T"), True, 40)
                        cad = cad & txtCodigo(121).Text
                        cad = cad & Format(DBLet(Rs!fech_ant, "F"), "dd/mm/yyyy")
                        cad = cad & Format(DBLet(Rs!fech_act, "F"), "dd/mm/yyyy")
                    
                        Print #NF, cad
                    End If
                End If
            
            End If
        Wend
        Set Rs = Nothing
    
        Close #NF
    
        FileCopy App.Path & "\exportacion.txt", txtCodigo(134).Text
    
    Else
        MsgBox "No hay registros para procesar", vbExclamation
        Exit Sub
    End If
             
eError:
    If Err.Number <> 0 Then
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
        lblProgres(2).Caption = ""
        conn.RollbackTrans
        Pb8.visible = False
    Else
        MsgBox "Proceso realizado correctamente.", vbExclamation
        conn.CommitTrans
        
        Pb1.visible = False
        lblProgres(2).Caption = ""
        lblProgres(3).Caption = ""
        cmdCancel_Click (16)
        Pb8.visible = False
    End If

End Sub

Private Sub CmdAcepImportar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim CodTipom As String
Dim Sql As String
Dim B As Boolean

    
    If Not DatosOK Then Exit Sub
    
'    Me.cd1.DefaultExt = "TXT"
'    Me.cd1.FileName = "lecthp.txt"
'    Me.cd1.ShowOpen
    
    cd1.FileName = txtCodigo(128).Text
    
    If Me.cd1.FileName <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

'fin
        If ProcesarFichero2(Me.cd1.FileName) Then
              cadTabla = "tmpinformes"
              cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
              
              Sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
              
              If TotalRegistros(Sql) <> 0 Then
                  MsgBox "Hay errores en el Traspaso de Lecturas. Debe corregirlos previamente.", vbExclamation
                  cadTitulo = "Errores de Traspaso de Lecturas"
                  cadNombreRPT = "rErroresLecturas.rpt"
                  LlamarImprimir
                  Exit Sub
              Else
                  conn.BeginTrans
                  B = ProcesarFichero(Me.cd1.FileName)
              End If
        End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Err.Number <> 0 Or Not B Then
        If Err.Number = 32755 Then Exit Sub
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
        cmdCancel_Click (15)
    End If
End Sub

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
    
    tabla = "rpozticketsmanta"
    
    If Not DatosOK Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtCodigo(116).Text)
    cHasta = Trim(txtCodigo(117).Text)
    nDesde = txtNombre(116).Text
    nHasta = txtNombre(117).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha riego
    cDesde = Trim(txtCodigo(118).Text)
    cHasta = Trim(txtCodigo(119).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecriego}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    If Not AnyadirAFormula(cadSelect, "not " & tabla & ".fecriego is null") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "not isnull({" & tabla & ".fecriego})") Then Exit Sub
    
    '++
    '[Monica]25/09/2014: añadimos la fecha de pago y el tipo de pago que es
    'D/H Fecha pago
    cDesde = Trim(txtCodigo(110).Text)
    cHasta = Trim(txtCodigo(111).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecpago}"
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
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarAlbaranes

    Screen.MousePointer = vbHourglass


    CargarAlbaranes = False

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "Select " & vUsu.Codigo & ",numalbar, fecalbar FROM rpozticketsmanta inner join rsocios on rpozticketsmanta.codsocio = rsocios.codsocio "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    If Option1(11).Value Then
        ' no hacemos nada
    
    Else
        If cWhere <> "" Then
            Sql = Sql & " and "
        Else
            Sql = Sql & " where "
        End If
        
        If Option1(13).Value Then ' contado
            Sql = Sql & " ((numalbar, fecalbar) in (select numalbar, fecalbar from rrecibpozos where codtipom = 'RMT' and escontado = 1)  or        "
            Sql = Sql & " (not (numalbar, fecalbar) in (select numalbar, fecalbar from rrecibpozos where codtipom = 'RMT') and rsocios.cuentaba='8888888888'))"
        Else ' banco
            Sql = Sql & " ((numalbar, fecalbar) in (select numalbar, fecalbar from rrecibpozos where codtipom = 'RMT' and escontado = 0)  or       "
            Sql = Sql & " (not (numalbar, fecalbar) in (select numalbar, fecalbar from rrecibpozos where codtipom = 'RMT') and rsocios.cuentaba<>'8888888888'))"
        End If
    End If
    
    Sql2 = "insert into tmpinformes (codusu, importe1, fecha1) "
    conn.Execute Sql2 & Sql
    
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
Dim Sql As String
Dim Sql3 As String
Dim Sql2 As String
Dim ctabla1 As String
Dim cad As String

Dim SqlZonas As String
    
Dim cadSelect1 As String
Dim cadFormula1 As String
Dim cadena As String
    
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
    If Not DatosOK Then Exit Sub
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H Socio
    If txtCodigo(124).Text <> "" Then
        Sql1 = Sql1 & " and rr.codsocio >= " & DBSet(txtCodigo(124).Text, "N")
    End If
    If txtCodigo(125).Text <> "" Then
        Sql1 = Sql1 & " and rr.codsocio <= " & DBSet(txtCodigo(125).Text, "N")
    End If
    If txtCodigo(124).Text <> "" Or txtCodigo(125).Text <> "" Then
        cad = ""
        If txtCodigo(124).Text <> "" Then cad = cad & " DESDE: " & txtCodigo(124).Text & " " & txtNombre(124).Text
        If txtCodigo(125).Text <> "" Then cad = cad & "  HASTA: " & txtCodigo(125).Text & " " & txtNombre(125).Text
        CadParam = CadParam & "pDHSocio=""" & cad & """|"
        numParam = numParam + 1
    End If
    
    'D/H fecha
    If txtCodigo(122).Text <> "" Then
        Sql1 = Sql1 & " and rr.fecfactu >= " & DBSet(txtCodigo(122).Text, "F")
    End If
    If txtCodigo(123).Text <> "" Then
        Sql1 = Sql1 & " and rr.fecfactu <= " & DBSet(txtCodigo(123).Text, "F")
    End If
    If txtCodigo(122).Text <> "" Or txtCodigo(123).Text <> "" Then
        cad = ""
        If txtCodigo(122).Text <> "" Then cad = cad & " DESDE: " & txtCodigo(122).Text
        If txtCodigo(123).Text <> "" Then cad = cad & "  HASTA: " & txtCodigo(123).Text
        CadParam = CadParam & "pDHFecha=""" & cad & """|"
        numParam = numParam + 1
    End If

    ' hidrante
    If txtCodigo(126).Text <> "" Then Sql1 = Sql1 & " and rr.hidrante >= " & DBSet(txtCodigo(126).Text, "N")
    If txtCodigo(127).Text <> "" Then Sql1 = Sql1 & " and rr.hidrante <= " & DBSet(txtCodigo(127).Text, "N")
    If txtCodigo(102).Text <> "" Or txtCodigo(103).Text <> "" Then
        cad = ""
        If txtCodigo(126).Text <> "" Then cad = cad & " DESDE: " & txtCodigo(126).Text
        If txtCodigo(127).Text <> "" Then cad = cad & "  HASTA: " & txtCodigo(127).Text
        CadParam = CadParam & "pDHHidrante=""" & cad & """|"
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
Dim Sql As String
Dim Sql3 As String
Dim Sql2 As String
Dim ctabla1 As String
Dim cad As String

Dim SqlZonas As String
    
Dim cadSelect1 As String
Dim cadFormula1 As String
Dim cadena As String
    
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
    Sql = "where ((cc.impvenci + coalesce(cc.gastos,0) - coalesce(cc.impcobro,0) <> 0)) "
    Sql = Sql & " and cc.impvenci > 0 "
    Sql = Sql & " and rr.codtipom = ll.codtipom "
    Sql = Sql & " and rr.numfactu = ll.numfactu "
    Sql = Sql & " and rr.fecfactu = ll.fecfactu "
    Sql = Sql & " and mid(cc.codmacta,5,6) = rr.codsocio"
    Sql = Sql & " and rr.codtipom = tt.codtipom "
    Sql = Sql & " and cc.numserie = tt.letraser "
    Sql = Sql & " and cc.codfaccl = rr.numfactu"
    Sql = Sql & " and cc.fecfaccl = rr.fecfactu"
    Sql = Sql & " and ll.codcampo = cam.codcampo"
    Sql = Sql & " and rr.codtipom = 'TAL'"
    Sql = Sql & " and  mid(cc.codmacta,5,6) = ss.codsocio"
    Sql = Sql & " and cam.codzonas = zz.codzonas"
    
    
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
    
    If Not DatosOK Then Exit Sub
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H Socio
    If txtCodigo(104).Text <> "" Then
        Sql = Sql & " and rr.codsocio >= " & DBSet(txtCodigo(104).Text, "N")
        Sql1 = Sql1 & " and rr.codsocio >= " & DBSet(txtCodigo(104).Text, "N")
        Sql2 = Sql2 & " and rr.codsocio >= " & DBSet(txtCodigo(104).Text, "N")
    End If
    If txtCodigo(105).Text <> "" Then
        Sql = Sql & " and rr.codsocio <= " & DBSet(txtCodigo(105).Text, "N")
        Sql1 = Sql1 & " and rr.codsocio <= " & DBSet(txtCodigo(105).Text, "N")
        Sql2 = Sql2 & " and rr.codsocio <= " & DBSet(txtCodigo(105).Text, "N")
    End If
    If txtCodigo(104).Text <> "" Or txtCodigo(105).Text <> "" Then
        cad = ""
        If txtCodigo(104).Text <> "" Then cad = cad & " DESDE: " & txtCodigo(104).Text & " " & txtNombre(104).Text
        If txtCodigo(105).Text <> "" Then cad = cad & "  HASTA: " & txtCodigo(105).Text & " " & txtNombre(105).Text
        CadParam = CadParam & "pDHSocio=""" & cad & """|"
        numParam = numParam + 1
    End If
    
    'D/H fecha
    If txtCodigo(106).Text <> "" Then
        Sql = Sql & " and rr.fecfactu >= " & DBSet(txtCodigo(106).Text, "F")
        Sql1 = Sql1 & " and rr.fecfactu >= " & DBSet(txtCodigo(106).Text, "F")
        Sql2 = Sql2 & " and rr.fecfactu >= " & DBSet(txtCodigo(106).Text, "F")
    End If
    If txtCodigo(107).Text <> "" Then
        Sql = Sql & " and rr.fecfactu <= " & DBSet(txtCodigo(107).Text, "F")
        Sql1 = Sql1 & " and rr.fecfactu <= " & DBSet(txtCodigo(107).Text, "F")
        Sql2 = Sql2 & " and rr.fecfactu <= " & DBSet(txtCodigo(107).Text, "F")
    End If
    If txtCodigo(106).Text <> "" Or txtCodigo(107).Text <> "" Then
        cad = ""
        If txtCodigo(107).Text <> "" Then cad = cad & " DESDE: " & txtCodigo(106).Text
        If txtCodigo(108).Text <> "" Then cad = cad & "  HASTA: " & txtCodigo(107).Text
        CadParam = CadParam & "pDHFecha=""" & cad & """|"
        numParam = numParam + 1
    End If

    ' braçal
    If txtCodigo(108).Text <> "" Then
        Sql = Sql & " and cam.codzonas >= " & DBSet(txtCodigo(108).Text, "N")
        Sql2 = Sql2 & " and cam.codzonas >= " & DBSet(txtCodigo(108).Text, "N")
    End If
    If txtCodigo(109).Text <> "" Then
        Sql = Sql & " and cam.codzonas <= " & DBSet(txtCodigo(109).Text, "N")
        Sql2 = Sql2 & " and cam.codzonas <= " & DBSet(txtCodigo(109).Text, "N")
    End If
    If txtCodigo(108).Text <> "" Or txtCodigo(109).Text <> "" Then
        cad = ""
        If txtCodigo(108).Text <> "" Then cad = cad & " DESDE: " & txtCodigo(108).Text & " " & txtNombre(108).Text
        If txtCodigo(109).Text <> "" Then cad = cad & "  HASTA: " & txtCodigo(109).Text & " " & txtNombre(109).Text
        CadParam = CadParam & "pDHZona=""" & cad & """|"
        numParam = numParam + 1
    End If

    ' sector
    If txtCodigo(102).Text <> "" Then Sql1 = Sql1 & " and mid(rr.hidrante,1,2) >= " & DBSet(txtCodigo(102).Text, "N")
    If txtCodigo(103).Text <> "" Then Sql1 = Sql1 & " and mid(rr.hidrante,1,2) <= " & DBSet(txtCodigo(103).Text, "N")
    If txtCodigo(102).Text <> "" Or txtCodigo(103).Text <> "" Then
        cad = ""
        If txtCodigo(102).Text <> "" Then cad = cad & " DESDE: " & txtCodigo(102).Text
        If txtCodigo(103).Text <> "" Then cad = cad & "  HASTA: " & txtCodigo(103).Text
        CadParam = CadParam & "pDHSector=""" & cad & """|"
        numParam = numParam + 1
    End If
    
    If Option7.Value Then CadParam = CadParam & "pTipo=2|" 'sector
    If Option8.Value Then CadParam = CadParam & "pTipo=1|" 'braçal
    If Option9.Value Then CadParam = CadParam & "pTipo=0|" 'ambos
    numParam = numParam + 1
    
    If CargarTemporalRecibosPdtes(cTabla, Sql, Sql2, ctabla1, Sql1) Then
        If HayRegParaInforme("tmpinformes", "tmpinformes.codusu = " & vUsu.Codigo) Then
        
            cadTitulo = "Recibos Pendientes de Cobro"

        
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            If Option7.Value Then cadFormula = cadFormula & " and {tmpinformes.campo1} = 2" 'sector
            If Option8.Value Then cadFormula = cadFormula & " and {tmpinformes.campo1} = 1" 'braçal
            
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
Dim Sql As String
Dim Sql3 As String

Dim SqlZonas As String
    
Dim cadSelect1 As String
Dim cadFormula1 As String
Dim cadena As String
    
Dim CadSelect0 As String
Dim SqlZonas0 As String
    
    InicializarVbles
    
    If Not DatosOK Then Exit Sub
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '======== FORMULA  ====================================
    'D/H Hidrante
    cDesde = Trim(txtCodigo(98).Text)
    cHasta = Trim(txtCodigo(99).Text)
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
    
    If CargarTemporalDiferencias(tabla, cadSelect) Then
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
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Cad2 As String
Dim Cad3 As String
Dim CadValues As String
Dim CadInsert As String
Dim Contador As String
Dim Nregs As Integer
Dim Fecha As Date

    On Error GoTo eCargarTemporal
    
    CargarTemporalDiferencias = False
    
    Screen.MousePointer = vbHourglass
    
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    If Me.Option4(0).Value Then
                                                'tipo,   h.contador,h.socio,c.socio, h.campo, c.campo,h.poligono,c.polig,  h.parce, c.parcela  h.hda,  c.hda
        Sql = "insert into tmpinformes (codusu, codigo1, nombre1, importe1, importe2,importe3,importe4, nombre2, importe5, nombre3, importeb1, precio1, precio2) "
    
        Sql = Sql & "SELECT " & vUsu.Codigo & ",0 tipo, rpozos.hidrante,rpozos.codsocio,rcampos.codsocio, rpozos.codcampo, rcampos.codcampo, rpozos.poligono, rcampos.poligono, rpozos.parcelas, rcampos.parcela, rpozos.hanegada, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4)"
        Sql = Sql & " FROM rpozos,rcampos "
        Sql = Sql & "  WHERE rpozos.poligono=rcampos.poligono AND rpozos.parcelas=rcampos.parcela AND rpozos.codcampo<>rcampos.codcampo "
        Sql = Sql & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtCodigo(98).Text <> "" Then Sql = Sql & " and rpozos.hidrante >= " & DBSet(txtCodigo(98).Text, "T")
        If txtCodigo(99).Text <> "" Then Sql = Sql & " and rpozos.hidrante <= " & DBSet(txtCodigo(99).Text, "T")
        Sql = Sql & " union "
        Sql = Sql & "SELECT " & vUsu.Codigo & ",1 tipo, rpozos.hidrante,rpozos.codsocio,rcampos.codsocio, rpozos.codcampo, rcampos.codcampo, rpozos.poligono, rcampos.poligono, rpozos.parcelas, rcampos.parcela, rpozos.hanegada, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4)"
        Sql = Sql & " FROM rpozos,rcampos "
        Sql = Sql & " WHERE  rpozos.codcampo=rcampos.codcampo AND (rpozos.poligono<>rcampos.poligono)  "
        Sql = Sql & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtCodigo(98).Text <> "" Then Sql = Sql & " and rpozos.hidrante >= " & DBSet(txtCodigo(98).Text, "T")
        If txtCodigo(99).Text <> "" Then Sql = Sql & " and rpozos.hidrante <= " & DBSet(txtCodigo(99).Text, "T")
        Sql = Sql & " union "
        Sql = Sql & "SELECT " & vUsu.Codigo & ",2 tipo, rpozos.hidrante,rpozos.codsocio,rcampos.codsocio, rpozos.codcampo, rcampos.codcampo, rpozos.poligono, rcampos.poligono, rpozos.parcelas, rcampos.parcela, rpozos.hanegada, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4)"
        Sql = Sql & " FROM rpozos,rcampos "
        Sql = Sql & " WHERE  rpozos.codcampo=rcampos.codcampo AND rpozos.codsocio <> rcampos.codsocio "
        Sql = Sql & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtCodigo(98).Text <> "" Then Sql = Sql & " and rpozos.hidrante >= " & DBSet(txtCodigo(98).Text, "T")
        If txtCodigo(99).Text <> "" Then Sql = Sql & " and rpozos.hidrante <= " & DBSet(txtCodigo(99).Text, "T")
        Sql = Sql & " union "
        Sql = Sql & "SELECT " & vUsu.Codigo & ",3 tipo, rpozos.hidrante,rpozos.codsocio,rcampos.codsocio, rpozos.codcampo, rcampos.codcampo, rpozos.poligono, rcampos.poligono, rpozos.parcelas, rcampos.parcela, rpozos.hanegada, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4)"
        Sql = Sql & " FROM rpozos,rcampos "
        Sql = Sql & " WHERE  rpozos.codcampo=rcampos.codcampo and "
        Sql = Sql & " truncate(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4),0) <> truncate(rpozos.hanegada,0) "
        Sql = Sql & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtCodigo(98).Text <> "" Then Sql = Sql & " and rpozos.hidrante >= " & DBSet(txtCodigo(98).Text, "T")
        If txtCodigo(99).Text <> "" Then Sql = Sql & " and rpozos.hidrante <= " & DBSet(txtCodigo(99).Text, "T")
        Sql = Sql & " union "
        Sql = Sql & "SELECT " & vUsu.Codigo & ",4 tipo, rpozos.hidrante,rpozos.codsocio,rcampos.codsocio, rpozos.codcampo, rcampos.codcampo, rpozos.poligono, rcampos.poligono, rpozos.parcelas, rcampos.parcela, rpozos.hanegada, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",4)"
        Sql = Sql & " FROM rpozos,rcampos "
        Sql = Sql & " WHERE  rpozos.codcampo=rcampos.codcampo AND (rpozos.parcelas<>rcampos.parcela)  "
        Sql = Sql & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtCodigo(98).Text <> "" Then Sql = Sql & " and rpozos.hidrante >= " & DBSet(txtCodigo(98).Text, "T")
        If txtCodigo(99).Text <> "" Then Sql = Sql & " and rpozos.hidrante <= " & DBSet(txtCodigo(99).Text, "T")
        
        conn.Execute Sql
    
    End If
    
    ' listado de discrepancias indefa
    If Me.Option4(1).Value Then
        If AbrirConexionIndefa() = False Then
            MsgBox "No se ha podido acceder a los datos de Indefa. ", vbExclamation
            Exit Function
        End If
        
                                            '   h.contador,h.poligono,h.parcelas, h.hdas     h.socio_revisado toma
        CadInsert = "insert into tmpinformes (codusu,  nombre1, nombre2, nombre3,   precio1, importe1,      importe2) values "

        Sql = "SELECT " & vUsu.Codigo & ", rpozos.hidrante,rpozos.poligono,rpozos.parcelas,rpozos.hanegada,rpozos.codsocio, rpozos.nroorden "
        Sql = Sql & " FROM rpozos "
        Sql = Sql & "  WHERE length(hidrante) = 6 and cast(hidrante as unsigned) "
        Sql = Sql & " and (rpozos.fechabaja is null or rpozos.fechabaja = '')"
        If txtCodigo(98).Text <> "" Then Sql = Sql & " and rpozos.hidrante >= " & DBSet(txtCodigo(98).Text, "T")
        If txtCodigo(99).Text <> "" Then Sql = Sql & " and rpozos.hidrante <= " & DBSet(txtCodigo(99).Text, "T")

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        Nregs = TotalRegistrosConsulta(Sql)
        If Nregs <> 0 Then
            Pb6.visible = True
            Label2(97).visible = True
            CargarProgres Pb6, Nregs
            DoEvents
        End If

        CadValues = ""
        While Not Rs.EOF
            IncrementarProgres Pb6, 1
            
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
                If Trim(DBLet(Rs!Poligono, "T")) <> Trim(DBLet(Rs2!Poligono, "T")) Then
                    If DBLet(Rs2!Poligono, "T") = "" Then
                        Cad2 = Cad2 & "'',"
                    Else
                        Cad2 = Cad2 & DBSet(Rs2!Poligono, "T") & ","
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
                
                '[Monica]30/10/2013: añadimos la parte de la toma
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
            conn.Execute CadInsert & CadValues
        End If
        CerrarConexionIndefa
        Pb6.visible = False
        Label2(97).visible = False
    End If
    
    'listado de contadores que existen en indefa y no en Escalona
    If Me.Option4(2).Value Then
        If AbrirConexionIndefa() = False Then
            MsgBox "No se ha podido acceder a los datos de Indefa. ", vbExclamation
            Exit Function
        End If
                                            '         h.contador,poligono,parcelas,hanegadas, indicamos si tiene fecha de baja
        CadInsert = "insert into tmpinformes (codusu,  nombre1, importe1, nombre2, precio1, importe2) values "
        
        CadValues = ""
        '[Monica]18/07/2013
        Sql2 = "select sector, hidrante, salida_tch, poligono, parcelas, hanegadas "
        Sql2 = Sql2 & " from rae_visitas_hidtomas "
        Sql2 = Sql2 & " where sector < '8' "
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
            
        Nregs = TotalRegistrosIndefa("select count(*) from rae_visitas_hidtomas")
        If Nregs <> 0 Then
            Pb6.visible = True
            Label2(97).visible = True
            CargarProgres Pb6, Nregs
            DoEvents
        End If
            
        While Not Rs2.EOF
            '[Monica]18/07/2013
            Contador = Format(Rs2!sector, "00") & Format(Rs2!Hidrante, "00") & Format(Rs2!salida_tch, "00")
            
            IncrementarProgres Pb6, 1
            Label2(97).Caption = "Procesando contador: " & Contador
            DoEvents
            
            Sql = "select count(*) from rpozos where hidrante = " & DBSet(Contador, "T")
            If TotalRegistros(Sql) = 0 Then
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Contador, "T") & ","
                CadValues = CadValues & DBSet(Rs2!Poligono, "N") & "," & DBSet(Rs2!parcelas, "T") & ","
                CadValues = CadValues & DBSet(Rs2!Hanegadas, "N") & ",0),"
            Else
                ' estan en escalona pero tienen fecha de baja pongo una marca para identificarlos
                Sql = "select count(*) from rpozos where hidrante = " & DBSet(Contador, "T") & " and not fechabaja is null"
                If TotalRegistros(Sql) = 1 Then
                    CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Contador, "T") & ","
                    CadValues = CadValues & DBSet(Rs2!Poligono, "N") & "," & DBSet(Rs2!parcelas, "T") & ","
                    CadValues = CadValues & DBSet(Rs2!Hanegadas, "N") & ",1),"
                End If
                
            End If
            
            Rs2.MoveNext
        Wend
        
        Set Rs2 = Nothing
        
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute CadInsert & CadValues
        End If
        CerrarConexionIndefa
        Pb6.visible = False
        Label2(97).visible = False
    End If
        
        
    'listado de contadores que existen en Escalona y no en Indefa
    If Me.Option4(3).Value Then
        If AbrirConexionIndefa() = False Then
            MsgBox "No se ha podido acceder a los datos de Indefa. ", vbExclamation
            Exit Function
        End If
                                            '   h.contador
        CadInsert = "insert into tmpinformes (codusu,  nombre1) values "
        
        CadValues = ""
        
        Sql = "select hidrante from rpozos where length(hidrante) = 6 and cast(hidrante as unsigned) "
        Sql = Sql & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "
        If txtCodigo(98).Text <> "" Then Sql = Sql & " and rpozos.hidrante >= " & DBSet(txtCodigo(98).Text, "T")
        If txtCodigo(99).Text <> "" Then Sql = Sql & " and rpozos.hidrante <= " & DBSet(txtCodigo(99).Text, "T")
        Sql = Sql & " order by hidrante "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Nregs = TotalRegistrosConsulta(Sql)
        If Nregs <> 0 Then
            Pb6.visible = True
            Label2(97).visible = True
            CargarProgres Pb6, Nregs
            DoEvents
        End If
        
        While Not Rs.EOF
            IncrementarProgres Pb6, 1
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
            conn.Execute CadInsert & CadValues
        End If
        CerrarConexionIndefa
        Pb6.visible = False
        Label2(97).visible = False
    End If
        
        
    'listado de contadores con socio bloqueado
    If Me.Option4(4).Value Then
                                            '   h.contador
        Sql = "insert into tmpinformes (codusu,  nombre1) "
        Sql = Sql & " select " & vUsu.Codigo & ",hidrante"
        Sql = Sql & " from rpozos "
        Sql = Sql & " where codsocio in (select codsocio from rsocios where codsitua > 1)"
        Sql = Sql & " and (rpozos.fechabaja is null or rpozos.fechabaja = '') "

        If txtCodigo(98).Text <> "" Then Sql = Sql & " and rpozos.hidrante >= " & DBSet(txtCodigo(98).Text, "T")
        If txtCodigo(99).Text <> "" Then Sql = Sql & " and rpozos.hidrante <= " & DBSet(txtCodigo(99).Text, "T")
        
        conn.Execute Sql
    End If
        
        
    'listado de contadores con consumo y no existencia en escalona
    If Me.Option4(5).Value Then
    
        Fecha = DevuelveValor("select max(fecproceso) from rpozos_lectura")
                                            
                                            '    h.contador consumo
        Sql = "insert into tmpinformes (codusu,  nombre1, importe1) "
        Sql = Sql & "select " & vUsu.Codigo & ", contador, "
        If vParamAplic.TipoLecturaPoz Then
            Sql = Sql & "lectura_bd "
        Else
            Sql = Sql & "lectura_equipo "
        End If
        Sql = Sql & " from rpozos_lectura "
        Sql = Sql & " where "
        If vParamAplic.TipoLecturaPoz Then
            Sql = Sql & " lectura_bd <> 0"
        Else
            Sql = Sql & " lectura_equipo <> 0 "
        End If
        
        Sql = Sql & " and (fecproceso is null or fecproceso = " & DBSet(Fecha, "F") & ")"
        Sql = Sql & " and not right(concat('00',contador),6) in (select hidrante from rpozos where (1=1) "
        If txtCodigo(98).Text <> "" Then Sql = Sql & " and rpozos.hidrante >= " & DBSet(txtCodigo(98).Text, "T")
        If txtCodigo(99).Text <> "" Then Sql = Sql & " and rpozos.hidrante <= " & DBSet(txtCodigo(99).Text, "T")
        Sql = Sql & ")"
  
        conn.Execute Sql
        
        
    End If
        
    CargarTemporalDiferencias = True
    Screen.MousePointer = vbDefault

    Exit Function

eCargarTemporal:
    CerrarConexionIndefa
    Pb6.visible = False
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
Dim Sql As String
Dim Sql3 As String

Dim SqlZonas As String
    
Dim cadSelect1 As String
Dim cadFormula1 As String
Dim cadena As String
    
Dim CadSelect0 As String
Dim SqlZonas0 As String
    
    InicializarVbles
    
    If Not DatosOK Then Exit Sub
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtCodigo(100).Text)
    cHasta = Trim(txtCodigo(101).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rsocios.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If

    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null") Then Exit Sub

    If CargarTemporalCCCErroneas(tabla, cadSelect) Then
        If HayRegParaInforme("tmpinformes", "tmpinformes.codusu = " & vUsu.Codigo) Then
        
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo

            indRPT = 98
            ConSubInforme = True
            cadTitulo = "Cuentas Bancarias de Socios erróneas"
        
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

    If Not DatosOK Then Exit Sub


    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtCodigo(115).Text)
'    cHasta = "" 'Trim(txtcodigo(116).Text)
'    nDesde = ""
'    nHasta = ""
'    If Not (cDesde = "" And cHasta = "") Then
'        'Cadena para seleccion Desde y Hasta
'        '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
'        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
'            Codigo = "{rcampos.codsocio}"
'        Else
'            Codigo = "{rsocios_pozos.codsocio}"
'        End If
'        TipCod = "N"
'        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
'    End If

    If Not AnyadirAFormula(cadSelect, "{rcampos.codsocio} = " & DBSet(txtCodigo(115).Text, "N")) Then Exit Sub


    vSQL = ""
    If txtCodigo(115).Text <> "" Then vSQL = vSQL & " and rcampos.codsocio = " & DBSet(txtCodigo(115).Text, "N")


'09/09/2010 : solo socios que no tengan fecha de baja
'    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
    
    '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        '[Monica]20/04/2015: añadimos la union a rzonas
        tabla = "(rcampos INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio) INNER JOIN rzonas ON rcampos.codzonas = rzonas.codzonas "
        
        If vParamAplic.Cooperativa = 10 Then
            tabla = "(" & tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
    
            If Not AnyadirAFormula(cadSelect, "{rsituacion.bloqueo} = 0") Then Exit Sub
            
        End If
    Else
        tabla = "rsocios_pozos INNER JOIN rsocios ON rsocios_pozos.codsocio = rsocios.codsocio "
        tabla = "(" & tabla & ") INNER JOIN rpozos ON rsocios_pozos.codsocio = rpozos.codsocio "
    End If

    '[Monica]08/05/2012: solo para Utxera y Escalona pq en turis se va a rsocios_pozos
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        If HayRegParaInforme(tabla, cadSelect) Then
            If ProcesoCarga(tabla, cadSelect) Then
                
                frmPOZMantaAux.Show vbModal
                
                TotalRegs = DevuelveValor("select sum(nroimpresion) from rpozauxmanta")
                If TotalRegs <> 0 Then
                    If MsgBox("¿ Desea continuar con el proceso ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        ProcesoFacturacionConsumoMantaESCALONANew tabla, cadSelect
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
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoCarga
    
    ProcesoCarga = False
    
    Sql = "delete from rpozauxmanta "
    conn.Execute Sql

    Sql = "select rcampos.codsocio, rcampos.codcampo, rcampos.codvarie, rcampos.codparti, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, "
    Sql = Sql & " round(supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) hanegadas, "
    '[Monica]20/04/2015: el preciomanta viene de rzonas
'    Sql = Sql & " round(round(supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) * " & DBSet(txtCodigo(112).Text, "N") & ",2) importe, 0  from " & cTabla
    Sql = Sql & " round(round(supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) * preciomanta,2) importe, 0  from " & cTabla
    If cWhere <> "" Then Sql = Sql & " where " & cWhere

    Sql2 = "insert into rpozauxmanta (codsocio, codcampo, codvarie, codparti, codzonas, poligono, parcela, subparce, hanegadas, importe, nroimpresion) "
    Sql2 = Sql2 & Sql
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
    
    If Not DatosOK Then Exit Sub
    
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    tabla = "rpozos"

    cadSelect = " rpozos.hidrante = " & DBSet(txtCodigo(55).Text, "T")     ' Hidrante
    
    
    '[Monica]23/09/2011: de momento solo rectifico las facturas de quatretonda
    Select Case vParamAplic.Cooperativa
        '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
        Case 8, 10 ' UTXERA
    
    
        Case 7 ' Quatretonda
            Dim B As Boolean
            
            Consumo = 0
            B = CalculoConsumoHidrante(txtCodigo(55).Text, txtCodigo(51).Text, Consumo)
             
            If B Then
                Check1(2).Value = 1
                Check1(3).Value = 1
                ProcesoFacturacionConsumo tabla, cadSelect, txtCodigo(54).Text, Consumo, True
            End If
        
        Case Else ' MALLAES
    
    
    End Select

End Sub

Private Sub CmdAceptarEtiq_Click()
Dim campo As String
Dim tabla As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim Sql As String
Dim Sql2 As String
Dim I As Long

    InicializarVbles
    
    If Not DatosOK Then Exit Sub
    
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
    If txtCodigo(45).Text <> "" Then
        CadParam = CadParam & "pLinea1="" " & txtCodigo(45).Text & """|"
    Else
        CadParam = CadParam & "pLinea1=""""|"
    End If
    numParam = numParam + 1
    
    'Parametro Linea 2
    If txtCodigo(46).Text <> "" Then
        CadParam = CadParam & "pLinea2="" " & txtCodigo(46).Text & """|"
    Else
        CadParam = CadParam & "pLinea2=""""|"
    End If
    numParam = numParam + 1
    
    'Parametro Linea 3
    If txtCodigo(47).Text <> "" Then
        CadParam = CadParam & "pLinea3="" " & txtCodigo(47).Text & """|"
    Else
        CadParam = CadParam & "pLinea3=""""|"
    End If
    numParam = numParam + 1
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = ""
    For I = 1 To CLng(txtCodigo(44).Text)
        Sql = Sql & "(" & vUsu.Codigo & "," & I & "),"
    Next I
    
    Sql2 = "insert into tmpinformes (codusu,codigo1) values "
    Sql2 = Sql2 & Mid(Sql, 1, Len(Sql) - 1) ' quitamos la ultima coma
    
    conn.Execute Sql2
    
    cadFormula = ""
    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu}=" & vUsu.Codigo) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{tmpinformes.codusu}=" & vUsu.Codigo) Then Exit Sub
    
    tabla = "tmpinformes"
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
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
Dim Sql As String
Dim Sql3 As String

Dim SqlZonas As String
    
Dim cadSelect1 As String
Dim cadFormula1 As String
Dim cadena As String
    
Dim CadSelect0 As String
Dim SqlZonas0 As String
    
    InicializarVbles
    
    If Not DatosOK Then Exit Sub
    
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    Select Case Index
        Case 0 ' Listado de Toma de lectura de contador
            'D/H Hidrante
            cDesde = Trim(txtCodigo(0).Text)
            cHasta = Trim(txtCodigo(1).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rpozos.hidrante}"
                TipCod = "T"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHHidrante=""") Then Exit Sub
            End If
            
            cTabla = tabla
            
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
            cDesde = Trim(txtCodigo(18).Text)
            cHasta = Trim(txtCodigo(19).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rpozos.hidrante}"
                TipCod = "T"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHHidrante=""") Then Exit Sub
            End If
            
            'D/H fecha
            cDesde = Trim(txtCodigo(16).Text)
            cHasta = Trim(txtCodigo(17).Text)
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
        
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = nomDocu
            
            If vParamAplic.Cooperativa = 7 Then
                If CargarTemporal(tabla, cadSelect) Then
                    If HayRegParaInforme("tmpinformes", "tmpinformes.codusu = " & vUsu.Codigo) Then
                        CadParam = CadParam & "pUsu=" & vUsu.Codigo & "|"
                        numParam = numParam + 1
                        ConSubInforme = True
                        LlamarImprimir
                    End If
                End If
            Else
                If HayRegParaInforme(tabla, cadSelect) Then
                    LlamarImprimir
                End If
            End If
    
        Case 2  ' opcionlistado = 10 --> cartas de tallas
            
            '[Monica]10/06/2013: Cambiamos, las cartas de talla no tienen que estar generadas para crearlas
            

            '======== FORMULA  ====================================
            'D/H Socio
            cDesde = Trim(txtCodigo(67).Text)
            cHasta = Trim(txtCodigo(68).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rsocios.codsocio}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
            End If
            
            vSQL = ""
            If txtCodigo(67).Text <> "" Then vSQL = vSQL & " and rsocios.codsocio >= " & DBSet(txtCodigo(67).Text, "N")
            If txtCodigo(68).Text <> "" Then vSQL = vSQL & " and rsocios.codsocio <= " & DBSet(txtCodigo(68).Text, "N")
        
        
            '[Monica]19/09/2012: se factura al propietario de los campos | 13/03/2014:se factura al socio antes al propietario
            tabla = "rcampos INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
            tabla = "(" & tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
            
        
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
            cadena = "select count(*) from " & tabla & " where " & CadSelect0
            If TotalRegistros(cadena) <> 0 Then
                
                Set frmMens2 = New frmMensajes
                
                frmMens2.OpcionMensaje = 49
                frmMens2.cadena = "select distinct rcampos.codzonas, rzonas.nomzonas from (" & tabla & ") inner join rzonas on rcampos.codzonas = rzonas.codzonas where " & CadSelect0
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
            
            
            If Not FacturacionTallaPreviaESCALONA(tabla, cadSelect, txtCodigo(69).Text, Me.pb5, "Prefacturacion Talla") Then Exit Sub
            
            '[Monica]11/04/2013: añadimos los textos de la carta parametrizados
            CadParam = CadParam & "pFJunta=""" & txtCodigo(88).Text & """|"
            CadParam = CadParam & "pFInicio=""" & txtCodigo(89).Text & """|"
            CadParam = CadParam & "pFinCom=""" & txtCodigo(90).Text & """|"
            CadParam = CadParam & "pFProhib=""" & txtCodigo(91).Text & """|"
            CadParam = CadParam & "pBonif=""" & txtCodigo(92).Text & """|"
            CadParam = CadParam & "pPerVol=""" & txtCodigo(93).Text & """|"
            CadParam = CadParam & "pRecarg=""" & txtCodigo(94).Text & """|"
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
            
                Sql = "select count(*) from " & tabla & " where " & cadSelect
            
                If TotalRegistros(Sql) <> 0 Then
                    'Enviarlo por e-mail
                    IndRptReport = indRPT
                    EnviarEMailMulti cadSelect, Titulo, nomDocu, tabla ' "rSocioCarta.rpt", Tabla  'email para socios
                Else
                    If MsgBox("No hay socios a enviar carta por email. ¿ Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
                End If
            
                If Not AnyadirAFormula(cadSelect1, "rsocios.maisocio is null or rsocios.maisocio=''") Then Exit Sub
                If Not AnyadirAFormula(cadFormula1, "isnull({rsocios.maisocio}) or {rsocios.maisocio}=''") Then Exit Sub
            
                Sql = "select count(*) from " & tabla & " where " & cadSelect1
                
                If TotalRegistros(Sql) <> 0 Then
                    cadFormula = cadFormula1
                    LlamarImprimir
                Else
                    MsgBox "No hay Socios para imprimir cartas.", vbExclamation
                End If
            
            Else
            
                If HayRegParaInforme(tabla, cadSelect) Then
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
                cDesde = Trim(txtCodigo(74).Text)
                cHasta = Trim(txtCodigo(75).Text)
                nDesde = ""
                nHasta = ""
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{rsocios.codsocio}"
                    TipCod = "N"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
                End If
                
                vSQL = ""
                If txtCodigo(74).Text <> "" Then vSQL = vSQL & " and rsocios.codsocio >= " & DBSet(txtCodigo(74).Text, "N")
                If txtCodigo(75).Text <> "" Then vSQL = vSQL & " and rsocios.codsocio <= " & DBSet(txtCodigo(75).Text, "N")
            
            
                '[Monica]19/09/2012: se factura al propietario de los campos | 13/03/2014: se factura al socio antes al propietario
                tabla = "rcampos INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
                tabla = "(" & tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
                
            
                If Not AnyadirAFormula(cadSelect, "{rsituacion.bloqueo} = 0") Then Exit Sub
                
'[Monica]25/03/2013: el campo no puede tener fecha de baja
                If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null") Then Exit Sub
'[Monica]25/03/2013: la situacion del campo debe de ser 1
                If Not AnyadirAFormula(cadSelect, "{rcampos.codsitua} = 1") Then Exit Sub
                
'
'[Monica]16/05/2017: ahora pedimos que socios vamos a facturar talla, por defecto todos marcados
'
'       SELECCIONA SOCIOS A FACTURAR POR SI ALGUN DIA LO QUIEREN
'
'                cadSelect = Replace(Replace(cadSelect, "{", ""), "}", "")
'                cadena = "select count(*) from " & Tabla & " where " & cadSelect
'
'                If TotalRegistros(cadena) <> 0 Then
'                    Set frmMensSoc = New frmMensajes
'
'                    frmMensSoc.OpcionMensaje = 67
'                    frmMensSoc.cadWHERE = "select distinct rsocios.codsocio, rsocios.nomsocio, rsocios.nifsocio from (" & Tabla & ") where " & cadSelect
'                    frmMensSoc.Show vbModal
'
'                    Set frmMensSoc = Nothing
'                End If
                
                
                '[Monica]31/05/2013: Comprobamos si de las zonas que vamos a facturar hay alguna con la suma de precios a 0 para ver si quieren o no continuar
                
                CadSelect0 = cadSelect
                CadSelect0 = Replace(Replace(CadSelect0, "{", ""), "}", "")
                SqlZonas0 = "rcampos.codzonas in (select codzonas from rzonas where if(precio1 is null,0,precio1) + if(precio2 is null,0,precio2) =  0)"
                If Not AnyadirAFormula(CadSelect0, SqlZonas0) Then Exit Sub
                cadena = "select count(*) from " & tabla & " where " & CadSelect0
                If TotalRegistros(cadena) <> 0 Then
                    
                    Set frmMens2 = New frmMensajes
                    
                    frmMens2.OpcionMensaje = 49
                    frmMens2.cadena = "select distinct rcampos.codzonas, rzonas.nomzonas from (" & tabla & ") inner join rzonas on rcampos.codzonas = rzonas.codzonas where " & CadSelect0
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
                
                
                ProcesoFacturacionTallaESCALONA tabla, cadSelect
            Else
                'D/H Socio
                cDesde = Trim(txtCodigo(74).Text)
                cHasta = Trim(txtCodigo(75).Text)
                nDesde = ""
                nHasta = ""
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{rsocios.codsocio}"
                    TipCod = "N"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
                End If
                
                '[Monica]19/09/2012: se actualiza la factura del propietario de los campos | 13/03/2014: el calculo de talla es para el socio no para el propietario
                tabla = "rcampos INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio "
                tabla = "(" & tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
                tabla = "(" & tabla & ") INNER JOIN rrecibpozos ON rsocios.codsocio = rrecibpozos.codsocio and rrecibpozos.codtipom = 'TAL' "
                
                If Not AnyadirAFormula(cadSelect, "{rrecibpozos.fecfactu} = " & DBSet(txtCodigo(73).Text, "F")) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rrecibpozos.fecfactu} = date(" & txtCodigo(73).Text & ")") Then Exit Sub
                
                
                If Check1(8).Value Then
                    Sql3 = "{rsocios.codbanco} <> '8888888888' and not {rsocios.codbanco} is null"
                    If Not AnyadirAFormula(cadSelect, Sql3) Then Exit Sub
                    Sql3 = "{rsocios.codbanco} <> '8888888888' and not isnull({rsocios.codbanco})"
                    If Not AnyadirAFormula(cadSelect, Sql3) Then Exit Sub
                End If
                    
                ProcesoFacturacionTallaESCALONA tabla, cadSelect
            
            End If
    End Select
End Sub

Private Function CargarTemporal(cTabla As String, cWhere As String) As Boolean
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
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
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


Private Function CargarTemporalRecibosPdtes(cTabla As String, cWhere As String, cWhere2 As String, ctabla1 As String, cwhere1 As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Sql2 As String
Dim SqlInsert As String


    On Error GoTo eCargarTemporal
    
    CargarTemporalRecibosPdtes = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    SqlInsert = "insert into tmpinformes (codusu, campo1, codigo1, nombre1,importe1, nombre2, importe2, importe3, importe4, nombre3, importeb2, fecha1, importe5) "
    
    Sql = "select " & vUsu.Codigo & ",1, mid(cc.codmacta,5,6) codsocio,ss.nomsocio, cam.codzonas, zz.nomzonas, cam.codcampo, cam.poligono, cam.parcela, rr.codtipom, rr.numfactu, rr.fecfactu, sum(round((coalesce(ll.precio1,0) + coalesce(ll.precio2,0)) * ll.hanegada,2)) importe"
    Sql = Sql & " from " & cTabla
    Sql = Sql & " " & cWhere
    Sql = Sql & " group by 1,2,3,4,5,6,7,8,9,10,11,12 "
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",1, mid(cc.codmacta,5,6) codsocio,ss.nomsocio, cam.codzonas, zz.nomzonas, cam.codcampo, cam.poligono, cam.parcela, rr.codtipom, rr.numfactu, rr.fecfactu, sum(round((coalesce(ll.precio1,0) + coalesce(ll.precio2,0)) * ll.hanegada,2)) importe"
    Sql = Sql & " from " & cTabla
    Sql = Sql & " " & cWhere2
    Sql = Sql & " group by 1,2,3,4,5,6,7,8,9,10,11,12 "
    
    
    conn.Execute SqlInsert & Sql
    
    
    SqlInsert = "insert into tmpinformes (codusu, campo1, codigo1, nombre1,importeb1, nombre3, importeb2, fecha1, importe5) "
    
    Sql = "select " & vUsu.Codigo & ",2, mid(cc.codmacta,5,6) codsocio, ss.nomsocio, mid(rr.hidrante,1,2) seccion, rr.codtipom, rr.numfactu, rr.fecfactu, sum(rr.totalfact)"
    Sql = Sql & " from " & ctabla1
    Sql = Sql & " " & cwhere1
    Sql = Sql & " group by 1,2,3,4,5,6,7,8 "
    
    conn.Execute SqlInsert & Sql

    CargarTemporalRecibosPdtes = True
    Exit Function

eCargarTemporal:
    MuestraError Err.Description, "Cargar Temporal", Err.Description
End Function


Private Function CargarTemporalRecibosConsumoPdtes(ctabla1 As String, cwhere1 As String, NConta As Integer) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Sql2 As String
Dim SqlInsert As String


    On Error GoTo eCargarTemporal
    
    CargarTemporalRecibosConsumoPdtes = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    SqlInsert = "insert into tmpinformes (codusu, campo1, codigo1, nombre1,importeb1, nombre3, importeb2, fecha1, importe5) "
    
    Sql = "select " & vUsu.Codigo & ",2, mid(cc.codmacta,5,6) codsocio, ss.nomsocio, rr.hidrante, rr.codtipom, rr.numfactu, rr.fecfactu, sum(rr.totalfact)"
    Sql = Sql & " from " & ctabla1
    Sql = Sql & " " & cwhere1
    Sql = Sql & " group by 1,2,3,4,5,6,7,8 "
    
    conn.Execute SqlInsert & Sql

    '[Monica]13/01/2015: cargamos el nro de reclamaciones que han hecho
    Sql = "update tmpinformes tt, usuarios.stipom ss "
    Sql = Sql & " set tt.importe1 = (select count(*) from conta" & NConta & ".shcocob aa where tt.importeb2 = aa.codfaccl "
    Sql = Sql & " and tt.fecha1 = aa.fecfaccl and ss.letraser = aa.numserie) "
    Sql = Sql & " where tt.codusu = " & vUsu.Codigo
    Sql = Sql & " and tt.nombre3 = ss.codtipom "
    
    conn.Execute Sql



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
        tabla = "rrecibpozos"
    End If
    
    If Not DatosOK Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
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
    cDesde = Trim(txtCodigo(40).Text)
    cHasta = Trim(txtCodigo(41).Text)
    nDesde = txtNombre(40).Text
    nHasta = txtNombre(41).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(42).Text)
    cHasta = Trim(txtCodigo(43).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    '[Monica]26/08/2011: añadido el nro de factura
    'D/H Nro Factura
    cDesde = Trim(txtCodigo(49).Text)
    cHasta = Trim(txtCodigo(50).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFactura= """) Then Exit Sub
    End If
    
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        tabla = tabla & " INNER JOIN rsocios ON rrecibpozos.codsocio = rsocios.codsocio "
    
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
    
    If HayRegistros(tabla, cadSelect) Then
        indRPT = 48
        ConSubInforme = False
        cadTitulo = "Facturas por Hidrante"
        
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
          
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
    
    If Not DatosOK Then Exit Sub
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    ' opcionlistado = 3 --> generacion de recibos de consumo
        
    '======== FORMULA  ====================================
    'D/H Hidrante
    cDesde = Trim(txtCodigo(11).Text)
    cHasta = Trim(txtCodigo(12).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.hidrante}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHHidrante=""") Then Exit Sub
    End If
    
    'D/H fecha
    cDesde = Trim(txtCodigo(13).Text)
    cHasta = Trim(txtCodigo(15).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.fech_act}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
'08/09/2010 : no va a ser el que tenga lectura a cero sino el que no tenga fecha de baja
'    If Not AnyadirAFormula(cadSelect, "{rpozos.lect_act} > 0") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{rpozos.fechabaja} is null") Then Exit Sub
    
    tabla = tabla & " INNER JOIN rsocios ON rpozos.codsocio = rsocios.codsocio "
    
    Select Case vParamAplic.Cooperativa
        '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
        Case 8, 10 ' UTXERA
            '[Monica]24/10/2011: dejamos que entren las facturas con lectura 0 solo para actualizar contadores
            '                    ponemos ({rpozos.consumo}) >= 0 antes ({rpozos.consumo}) > 0
            cadSelect = cadSelect & " and {rpozos.fech_act} is not null and {rpozos.lect_act} is not null and ({rpozos.consumo}) >= 0 "
        
            '[Monica]27/08/2012: en escalona dejamos unicamente los socios no bloqueados
            If vParamAplic.Cooperativa = 10 Then
                tabla = "(" & tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
            
                If Not AnyadirAFormula(cadSelect, "{rsituacion.bloqueo} = 0") Then Exit Sub
            End If
        
            ' un recibo por hidrante
            ProcesoFacturacionConsumoUTXERA tabla, cadSelect
    
        Case 7 ' Quatretonda
            '[Monica] 11/07/2011: tiene hidrantes que no son contadores y a los que solo se les facturan las acciones
            '                   : por lo tanto quito la condicion de la fecha de lectura actual
            'cadSelect = cadSelect & " and {rpozos.fech_act} is not null and {rpozos.lect_act} is not null "
        
            ProcesoFacturacionConsumo tabla, cadSelect, txtCodigo(14).Text, 0, False
    
        Case Else ' MALLAES
            cadSelect = cadSelect & " and {rpozos.fech_act} is not null and {rpozos.lect_act} is not null "
            '[Monica]07/03/2014: nuevo campo de si se cobra la cuota
            cadSelect = cadSelect & " and {rpozos.cobrarcuota} = 1 "
        
            ProcesoFacturacionConsumo tabla, cadSelect, txtCodigo(14).Text, 0, False
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

    If Not DatosOK Then Exit Sub


    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtCodigo(23).Text)
    cHasta = Trim(txtCodigo(24).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
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
    
    '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        tabla = "rsocios inner join rsituacion On rsituacion.bloqueo = 0"
        
        If HayRegParaInforme(tabla, cadSelect) Then
        
            If txtCodigo(23).Text <> txtCodigo(24).Text Or txtCodigo(23).Text = "" Or txtCodigo(24).Text = "" Then
                Set frmMen = New frmMensajes
                frmMen.cadWHERE = cadSelect
                frmMen.OpcionMensaje = 9 'Socios
                frmMen.Show vbModal
                Set frmMen = Nothing
                If cadSelect = "" Then Exit Sub
            End If
        End If
        
    Else
        tabla = "(rsocios_pozos INNER JOIN rsocios ON rsocios_pozos.codsocio = rsocios.codsocio) INNER JOIN rpozos ON rsocios.codsocio = rpozos.codsocio "
    End If
    
    ProcesoFacturacionContadores tabla, cadSelect

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

    If Not DatosOK Then Exit Sub


    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    '======== FORMULA  ====================================
    'D/H Socio
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
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
    If txtCodigo(6).Text <> "" Then vSQL = vSQL & " and rpozos.codsocio >= " & DBSet(txtCodigo(6).Text, "N")
    If txtCodigo(7).Text <> "" Then vSQL = vSQL & " and rpozos.codsocio <= " & DBSet(txtCodigo(7).Text, "N")

    'D/H hidrante
    cDesde = Trim(txtCodigo(62).Text)
    cHasta = Trim(txtCodigo(63).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.hidrante}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHHidrante=""") Then Exit Sub
    End If

    'D/H Poligono
    cDesde = Trim(txtCodigo(57).Text)
    cHasta = Trim(txtCodigo(58).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.poligono}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPoligono=""") Then Exit Sub
    End If

    'D/H Parcela
    cDesde = Trim(txtCodigo(59).Text)
    cHasta = Trim(txtCodigo(60).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rpozos.parcelas}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHParcela=""") Then Exit Sub
    End If

    'D/H fecha
    cDesde = Trim(txtCodigo(64).Text)
    cHasta = Trim(txtCodigo(65).Text)
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
        tabla = "rpozos INNER JOIN rsocios ON rpozos.codsocio = rsocios.codsocio "
        
        If vParamAplic.Cooperativa = 10 Then
            tabla = "(" & tabla & ") INNER JOIN rsituacion ON rsocios.codsitua = rsituacion.codsitua "
    
            If Not AnyadirAFormula(cadSelect, "{rsituacion.bloqueo} = 0") Then Exit Sub
            
        End If
    Else
        tabla = "rsocios_pozos INNER JOIN rsocios ON rsocios_pozos.codsocio = rsocios.codsocio "
        tabla = "(" & tabla & ") INNER JOIN rpozos ON rsocios_pozos.codsocio = rpozos.codsocio "
    End If

    '[Monica]08/05/2012: solo para Utxera y Escalona pq en turis se va a rsocios_pozos
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        If vSQL <> "" And txtCodigo(6).Text = txtCodigo(7).Text Then
            Set frmMens = New frmMensajes
        
            frmMens.OpcionMensaje = 37
            frmMens.cadWHERE = vSQL
            frmMens.Show vbModal
        
            Set frmMens = Nothing
        End If
    End If

    Select Case vParamAplic.Cooperativa
        '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
        Case 8 ' UTXERA
            ProcesoFacturacionMantenimientoUTXERA tabla, cadSelect
    
        Case 10 ' ESCALONA
            ProcesoFacturacionMantenimientoESCALONA tabla, cadSelect
    
        Case Else
            ProcesoFacturacionMantenimiento tabla, cadSelect
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
        tabla = "rrecibpozos"
    End If
    
    If Not DatosOK Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
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
    cDesde = Trim(txtCodigo(34).Text)
    cHasta = Trim(txtCodigo(35).Text)
    nDesde = txtNombre(34).Text
    nHasta = txtNombre(35).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Nro de Factura
    cDesde = Trim(txtCodigo(38).Text)
    cHasta = Trim(txtCodigo(39).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rrecibpozos.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(36).Text)
    cHasta = Trim(txtCodigo(37).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        tabla = tabla & " INNER JOIN rsocios ON rrecibpozos.codsocio = rsocios.codsocio "
    
    
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
    
    If HayRegistros(tabla, cadSelect) Then
        Select Case CodTipom
            Case "RCP", "FIN"
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
            Case "RMT"
                indRPT = 47
                ConSubInforme = True
                cadTitulo = "Reimpresión de Recibos Consumo Manta"
            
            '[Monica]14/01/2016: las rectificativas
            Case "RRC"
                indRPT = 46 ' impresion de recibos de consumo
                ConSubInforme = True
                cadTitulo = "Reimpresión de Recibos Rect.Consumo"
            Case "RRM"
                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                ConSubInforme = True
                cadTitulo = "Reimpresión de Recibos Rect.Mantenimiento"
            Case "RRV"
                indRPT = 47 'Impresion de recibos de contadores pozos
                ConSubInforme = True
                cadTitulo = "Reimpresión de Recibos Rect.Contadores"
            Case "RTA"
                indRPT = 47 'Impresion de recibos de talla
                ConSubInforme = True
                cadTitulo = "Reimpresión de Recibos Rect.Talla"
            Case "RRT"
                indRPT = 47 'Impresion de recibos de consumo a manta
                ConSubInforme = True
                cadTitulo = "Reimpresión de Recibos Rect.Consumo Manta"
        End Select
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
          
        If CodTipom = "TAL" Then nomDocu = Replace(nomDocu, "Mto.", "Tal.")
        If CodTipom = "RVP" Then nomDocu = Replace(nomDocu, "Mto.", "Cont.")
        If CodTipom = "RMT" Then nomDocu = Replace(nomDocu, "Mto.", "Manta.")
  
        '[Monica]14/01/2016: las rectificativas
        If CodTipom = "RTA" Then nomDocu = Replace(nomDocu, "Mto.", "Tal.")
        If CodTipom = "RRV" Then nomDocu = Replace(nomDocu, "Mto.", "Cont.")
        If CodTipom = "RRM" Then nomDocu = Replace(nomDocu, "Mto.", "Manta.")
  
  
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
          
        'Nombre fichero .rpt a Imprimir
        
        
        If vParamAplic.Cooperativa = 10 Then
            If CargarTemporalFrasPozos(tabla, cadSelect) Then
            
                Set frmMens = New frmMensajes
                
                frmMens.OpcionMensaje = 61
                frmMens.Show vbModal
                
                Set frmMens = Nothing
        
                If HayRegistros(tabla, cadSelect & " and rrecibpozos.imprimir=" & DBSet(vUsu.PC, "T")) Then
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim CadInsert As String
Dim CadValues As String
Dim numserie As String

    On Error GoTo eCargarTemporalFrasPozos

    CargarTemporalFrasPozos = False

    Sql = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql

    ' desmarcamos todas las facturas que vamos a imprimir
    Sql = "update rrecibpozos, rsocios set imprimir = null "
    Sql = Sql & " where rrecibpozos.codsocio = rsocios.codsocio "
    If cSelect <> "" Then Sql = Sql & " and " & cSelect
    
    conn.Execute Sql
    

    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If

    Sql = "select rrecibpozos.codtipom,rrecibpozos.numfactu,rrecibpozos.fecfactu,rrecibpozos.codsocio,rsocios.nomsocio, rrecibpozos.totalfact "
    Sql = Sql & " from rrecibpozos inner join rsocios on rrecibpozos.codsocio = rsocios.codsocio "
    If cSelect <> "" Then Sql = Sql & " where " & cSelect
    
    CadInsert = "insert into tmpinformes (codusu,nombre1,importe1,fecha1,codigo1,nombre2,importe2,campo1) VALUES "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    numserie = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(Mid(Me.Combo1(0).Text, 1, 3), "T"))
    
    Label4(52).Caption = ""
    
    While Not Rs.EOF
    
        Label4(52).Caption = "Comprobando factura: " & Format(DBLet(Rs!numfactu), "0000000")
        DoEvents
    
        CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs!CodTipom, "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & ","
        CadValues = CadValues & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!TotalFact, "N") & ","
    
        Sql = "select sum(coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) from scobro where numserie = " & DBSet(numserie, "T")
        Sql = Sql & " and codfaccl = " & DBSet(Rs!numfactu, "N")
        Sql = Sql & " and fecfaccl = " & DBSet(Rs!fecfactu, "F")

        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
        conn.Execute CadInsert & CadValues
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
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim codzonas As Integer

    On Error GoTo eErrores

    conn.BeginTrans


    Sql = "select * from rcampos order by codcampo"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
            Sql2 = "update rcampos set codzonas = " & DBSet(codzonas, "N") & "  where codcampo = " & DBSet(Rs!codcampo, "N")
        
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
                PonerFoco txtCodigo(0)
            Case 2  ' Listado de comprobacion de lecturas
                '[Monica]01/07/2013: introducimos los valores por defecto
                txtCodigo(16).Text = Format(Now, "dd/mm/yyyy")
                txtCodigo(17).Text = Format(Now, "dd/mm/yyyy")
            
                PonerFoco txtCodigo(18)
            
            Case 3 ' generacion de facturas de consumo
                '[Monica]01/07/2013: introducimos los valores por defecto
                txtCodigo(13).Text = Format(Now, "dd/mm/yyyy")
                txtCodigo(15).Text = Format(Now, "dd/mm/yyyy")
               
                PonerFoco txtCodigo(11)
            Case 4 ' generacion de facturas de mantenimiento
                PonerFoco txtCodigo(6)
            Case 5 ' generacion de facturas de contadores
                PonerFoco txtCodigo(23)
            Case 6 ' reimpresion de recibos
                '[Monica]01/07/2013: introducimos los valores por defecto
                txtCodigo(36).Text = Format(Now, "dd/mm/yyyy")
                txtCodigo(37).Text = Format(Now, "dd/mm/yyyy")
               
                
                PonerFoco txtCodigo(38)
                
                Option1(4).Value = True
                
            Case 7 ' informe de facturas por hidrante
                PonerFoco txtCodigo(40)
            
                Option1(7).Value = True
            
            Case 8 ' etiquetas contadores
                PonerFoco txtCodigo(45)
                
                txtCodigo(45).Text = "AGUA CON CUPO XXXM3/HG/MES"
                txtCodigo(46).Text = "DIA:"
                txtCodigo(47).Text = "LECTURA:"
                
                Nregs = DevuelveValor("select count(*) from rpozos")
                txtCodigo(44).Text = Format(Nregs, "###,###,##0")
                
            Case 9 ' rectificacion de facturas
                txtCodigo(54).Text = Format(Now, "dd/mm/yyyy")
                PonerFoco txtCodigo(52)
                
            Case 10 ' informe de tallas (recibos de mantenimiento de Escalona)
                PonerFoco txtCodigo(67)
                
                '[Monica]01/07/2013: introducimos los valores por defecto
                txtCodigo(69).Text = Format(Now, "dd/mm/yyyy")
                
                txtCodigo(88).Text = "29 de gener"
                txtCodigo(89).Text = "1 de març"
                txtCodigo(90).Text = "25 de febrer"
                txtCodigo(91).Text = "de l'1 SETEMBRE"
                txtCodigo(92).Text = "Març 2%"
                txtCodigo(93).Text = "Abril-Maig"
                txtCodigo(94).Text = "Juny fins Desembre 20%"
                                
                Label2(92).Caption = "Precios " & vParamAplic.NomZonaPOZ
                imgBuscar(19).ToolTipText = "Ver precios " & vParamAplic.NomZonaPOZ
                
                
            Case 11, 12 'recibos y bonificacion de talla
                PonerFoco txtCodigo(74)
                If OpcionListado = 11 Then txtCodigo(76).TabIndex = 236
                If OpcionListado = 12 Then ConexionConta

                txtCodigo(73).Text = Format(Now, "dd/mm/yyyy")
                
            Case 14
                PonerFoco txtCodigo(95)
                
            Case 15 ' listado de comprobacion de pozos
                Me.Option4(1).Value = True
                PonerFoco txtCodigo(98)
                
            Case 16 ' listado de comprobacion de cuentas bancarias de socios
                PonerFoco txtCodigo(100)
                
        
            Case 17 ' generacion de recibos a manta
                txtCodigo(113).Text = "TICKET RIEGO A MANTA"
                PonerFoco txtCodigo(115)
        
            Case 18 ' informe de recibos pendientes de cobro
                PonerFoco txtCodigo(104)
            
            Case 19 ' informe de recibos por fecha de riego
                PonerFoco txtCodigo(116)
                
                Option1(13).Value = True
            
            Case 20 ' informe de recibos de consumo pendientes de cobro
                PonerFoco txtCodigo(124)
            
                Option14.Value = True
                
            Case 21 ' importacion de datos
                PonerFoco txtCodigo(129)
                
                txtCodigo(128).Text = vParamAplic.PathEntradas & "\LECTHP.TXT"
                
                txtCodigo(131).Text = Format(Now, "dd/mm/yyyy")
                
                txtCodigo(129).Text = 7
                txtCodigo(130).Text = 7
                
                txtCodigo_LostFocus 129
                txtCodigo_LostFocus 130
                
            Case 22 ' exportacion de datos
                txtCodigo(121).Text = Format(Now, "dd/mm/yyyy")
                txtCodigo(134).Text = vParamAplic.PathEntradas & "\Exportacion" & Format(txtCodigo(121).Text, "yyyymmdd") & ".TXT"
                
                PonerFoco txtCodigo(121)
        
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim I As Integer
Dim Sql As String
Dim Rs As ADODB.Recordset

    'Icono del formulario
    Me.Icon = frmPpal.Icon


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
    
    
    For H = 0 To 34
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
    Me.FrameImporLecturas.visible = False
    Me.FrameExporLecturas.visible = False
    
    '[Monica]07/06/2013: Zona / Braçal
    Me.Label2(81).Caption = "Precios " & vParamAplic.NomZonaPOZ
    
    
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
        'LISTADOS DE MANTENIMIENTOS BASICOS
        '---------------------
        Case 1 ' Informe de Toma de Lectura
            FrameTomaLecturaVisible True, H, W
            indFrame = 0
            tabla = "rpozos"
            Me.Option1(0).Value = True
            
        Case 2 ' Informe de Comprobacion de lecturas
            FrameComprobacionVisible True, H, W
            indFrame = 0
            tabla = "rpozos"
            Label7.Caption = "Informe de Comprobación de Lecturas"
            Me.Pb1.visible = False
            
            If vParamAplic.Cooperativa = 17 Then Label2(27).Caption = "Contador"
        
        Case 3 ' generacion de recibos de consumo
            FrameReciboConsumoVisible True, H, W
            indFrame = 0
            tabla = "rpozos"
            txtCodigo(14).Text = Format(Now, "dd/mm/yyyy")
            
            Frame6.Enabled = (vParamAplic.Cooperativa = 7 Or vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            Frame6.visible = (vParamAplic.Cooperativa = 7 Or vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            Frame5.Enabled = Not (vParamAplic.Cooperativa = 7)
            Frame5.visible = Not (vParamAplic.Cooperativa = 7)
            
            If vParamAplic.Cooperativa = 7 Then
                Frame6.Top = 510
            End If
            
            
            '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                txtCodigo(2).Text = Format(DevuelveValor("select hastametcub1 from rtipopozos where codpozo = 1"), "0000000")
                txtCodigo(3).Text = Format(DevuelveValor("select hastametcub2 from rtipopozos where codpozo = 1"), "0000000")
                txtCodigo(4).Text = Format(DevuelveValor("select precio1 from rtipopozos where codpozo = 1"), "###,##0.0000")
                txtCodigo(5).Text = Format(DevuelveValor("select precio2 from rtipopozos where codpozo = 1"), "###,##0.0000")
                
                '[Monica]29/01/2014: por la insercion en tesoreria
                txtCodigo(48).MaxLength = 15
            Else
                txtCodigo(2).Text = Format(vParamAplic.Consumo1POZ, "0000000")
                txtCodigo(3).Text = Format(vParamAplic.Consumo2POZ, "0000000")
                txtCodigo(4).Text = Format(vParamAplic.Precio1POZ, "###,##0.00")
                txtCodigo(5).Text = Format(vParamAplic.Precio2POZ, "###,##0.00")
            End If
            
            Me.Pb1.visible = False
        
        Case 4 ' Generacion de recibos de mantenimiento
            FrameReciboMantenimientoVisible True, H, W
            indFrame = 0
            tabla = "rsocios_pozos"
            txtCodigo(10).Text = Format(Now, "dd/mm/yyyy")
            Me.Pb2.visible = False
            
            'Si es Escalona el concepto tiene que caber en textcsb33(40 posiciones)
            If vParamAplic.Cooperativa = 10 Then
                '[Monica]29/01/2014: limitacion del concepto al arimoney
                txtCodigo(9).MaxLength = 20 '40
                txtCodigo(8).Text = Format(DevuelveValor("select imporcuotahda from rtipopozos where codpozo = 1"), "###,##0.0000")
                Label2(6).Caption = "Euros/Hanegada"
                Check1(0).Value = 1
                Check1(1).Value = 1
            Else
                If vParamAplic.Cooperativa = 8 Then
                    txtCodigo(9).MaxLength = 20
                End If
            End If
            
        Case 5 ' Generacion de recibos de contadores
            FrameReciboContadorVisible True, H, W
            indFrame = 0
            tabla = "rsocios_pozos"
            txtCodigo(22).Text = Format(Now, "dd/mm/yyyy")
            Me.Pb3.visible = False
            
            '[Monica]27/06/2013: solo dejamos meter 40 caracteres para utxera y escalona pq se tiene que imprimir en la scobro
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                '[Monica]29/01/2014: solo dejamos meter 3 conceptos
                txtCodigo(20).MaxLength = 30 '40
                txtCodigo(25).MaxLength = 30 '40
                txtCodigo(27).MaxLength = 29 '40
                txtCodigo(29).MaxLength = 40
                txtCodigo(31).MaxLength = 40
            End If
        
        Case 6 ' Reimpresion de recibos de pozos
            FrameReimpresionVisible True, H, W
            tabla = "rrecibpozos"
            CargaCombo
            Combo1(0).ListIndex = 0
            
            '[Monica]11/03/2013: solo en el caso de escalona y utxera pedimos el tipo de pago
            Me.FrameTipoPago.Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            Me.FrameTipoPago.visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            
        Case 7 ' Informe de recibos por hidrante
            FrameFacturasHidranteVisible True, H, W
            tabla = "rrecibpozos"
            CargaCombo
            Combo1(1).ListIndex = 0
        
            FrameTipoPago2.visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            FrameTipoPago2.Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            
        Case 8 ' Etiquetas contadores
            FrameEtiquetasContadoresVisible True, H, W
            tabla = "tmpinformes"
        
        Case 9 ' Rectificacion de lecturas
            FrameRectificacionVisible True, H, W
            tabla = "rrecibpozos"
            CargaCombo
            Combo1(2).ListIndex = 0
        
        Case 10 ' Informe de Tallas recibos de Mto (solo visible para Escalona)
            FrameCartaTallasVisible True, H, W
            indFrame = 0
            tabla = "rrecibpozos"
        
        Case 11 ' Generacion de recibos de talla (solo para Escalona)
            FrameReciboTallaVisible True, H, W
            indFrame = 0
            tabla = "rrecibpozos"
            Me.Pb4.visible = False
            Check1(6).Value = 1
            Check1(7).Value = 1
        
            '[Monica]29/01/2014: longitud maxima del concepto
            txtCodigo(76).MaxLength = 15
        
        
            For I = 79 To 87
                txtCodigo(I).Text = ""
            Next I
            
            txtCodigo(72).Text = ""
            txtCodigo(66).Text = ""
            txtNombre(0).Text = ""
            txtNombre(2).Text = ""
            txtNombre(4).Text = ""
            txtNombre(8).Text = ""
        
            I = 0
            Sql = "select rpretallapoz.codzonas, rzonas.nomzonas, rpretallapoz.precio1, rpretallapoz.precio2 "
            Sql = Sql & " from rpretallapoz left join rzonas on rpretallapoz.codzonas = rzonas.codzonas "
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                If DBLet(Rs!codzonas) = 0 Then
                    txtCodigo(72).Text = DBLet(Rs!Precio1, "N")
                    txtCodigo(66).Text = DBLet(Rs!Precio2, "N")
                    PonerFormatoDecimal txtCodigo(72), 7
                    PonerFormatoDecimal txtCodigo(66), 7
                Else
                    txtCodigo(79 + I).Text = DBLet(Rs!codzonas, "N")
                    PonerFormatoEntero txtCodigo(79 + I)
                    txtCodigo(80 + I).Text = DBLet(Rs!Precio1, "N")
                    txtCodigo(81 + I).Text = DBLet(Rs!Precio2, "N")
                    txtNombre(79 + I).Text = DBLet(Rs!nomzonas)
                    PonerFormatoDecimal txtCodigo(80 + I), 7
                    PonerFormatoDecimal txtCodigo(81 + I), 7
                
                    I = I + 3
                End If
                Rs.MoveNext
            Wend
            RealizarCalculos
        
        Case 12 ' Calculo de bonificacion de recibos de talla
            FrameReciboTallaVisible True, H, W
            indFrame = 0
            tabla = "rrecibpozos"
            Me.Pb4.visible = False
            Check1(6).Value = 1
            Check1(7).Value = 1
        
            txtCodigo(78).TabIndex = 236
            txtCodigo(77).TabIndex = 237
            
        
        Case 13
            Me.Frame10.visible = True
            Me.Frame10.Height = 3469
            Me.Frame10.Width = 7335
            W = Me.Frame10.Width
            H = Me.Frame10.Height
            
        Case 14 ' asignacion de precios de talla
            FrameAsignacionPreciosVisible True, H, W
        
            indFrame = 0
            tabla = "rzonas"
            
        Case 15 ' informe de diferencias
            FrameComprobacionDatosVisible True, H, W
            indFrame = 0
            tabla = "rpozos"
            Me.Pb6.visible = False
        
        Case 16 ' informe de comprobacion de cuentas bancarias de socios
            FrameComprobacionCCCVisible True, H, W
            indFrame = 0
            tabla = "rsocios"
            
        Case 17
            FrameReciboConsumoMantaVisible True, H, W
            indFrame = 0
            tabla = "rcampos"
            txtCodigo(114).Text = Format(Now, "dd/mm/yyyy")
            Me.pb7.visible = False
            
            'Si es Escalona el concepto tiene que caber en textcsb33(40 posiciones)
            If vParamAplic.Cooperativa = 10 Then
                '[Monica]29/01/2014: limitacion del concepto al arimoney
                txtCodigo(113).MaxLength = 20 '40
                txtCodigo(112).Text = Format(DevuelveValor("select imporcuotahda from rtipopozos where codpozo = 1"), "###,##0.0000")
                Label2(109).Caption = "Euros/Hanegada"
                Check1(9).Value = 1
                Check1(10).Value = 1
            Else
                If vParamAplic.Cooperativa = 8 Then
                    txtCodigo(113).MaxLength = 20
                End If
            End If
            
        Case 18 ' informe de recibos pendientes de cobro
            FrameRecPdtesCobroVisible True, H, W
            tabla = "rrecibpozos"
            
        Case 19 ' informe de recibos por fecha de riego
            FrameInfMantaFechaRiegoVisible True, H, W
            tabla = "rpozticketsmanta"
    
        Case 20 ' informe de recibos de consumo pendientes de consumo
            FrameRecConsPdtesCobroVisible True, H, W
            tabla = "rpozticketsmanta"
    
        Case 21 ' importacion
            FrameImporLecturasVisible True, H, W
            tabla = "rpozos"
            
        Case 22 ' exportacion
            FrameExporLecturasVisible True, H, W
            tabla = "rpozos"
    
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
    txtCodigo(CByte(imgFecha(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCodigo(130).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNombre(130).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
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


Private Sub frmMens3_datoseleccionado(CadenaSeleccion As String)
    Continuar = (CadenaSeleccion = "1")
End Sub

Private Sub frmMens4_DatoSeleccionado(CadenaSeleccion As String)
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



Private Sub frmMensSoc_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        cadSelect = cadSelect & " and rsocios.codsocio in (" & CadenaSeleccion & ") "
    Else
        cadSelect = cadSelect & " and rsocios.codsocio is null"
    End If
End Sub

Private Sub frmPoz_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
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
            
        Case 25, 26 ' braçal
            AbrirFrmZonas (Index + 83)
            
        Case 27, 28 ' socios de recibos de manta por fecha de riego
            AbrirFrmSocios (Index + 89)
            
        Case 29 ' comunidad
            AbrirFrmComunidades (Index + 100)
        Case 30 ' concepto
            AbrirFrmConceptos (Index + 130)
            
        Case 31, 32 ' socios de recibos de consumo pendientes de cobro
            AbrirFrmSocios (Index + 93)
            
        Case 33 ' path del fichero
            AbrirDialogo (0)
            
        Case 34 ' path del fichero de exportacion
            AbrirDialogo (1)
    End Select
    
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub AbrirDialogo(Opcion As Byte)

    On Error GoTo EA
    
    With Me.cd1
        Select Case Opcion
        Case 0, 2
            .DialogTitle = "Archivo origen de datos"
        Case 1
            .DialogTitle = "Archivo destino de datos"
        End Select
        .Filter = "TEXT (*.txt)|*.txt"
        .InitDir = vParamAplic.PathEntradas
        .CancelError = True
        
        .ShowOpen
        If Opcion = 0 Then
            txtCodigo(128).Text = .FileName
        Else
            txtCodigo(121).Text = .FileName
        End If
    End With
    
EA:
End Sub




Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object
    Dim Indice As Integer

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
            Indice = 14
        Case 1
            Indice = 10
        Case 0, 2, 3
            Indice = Index + 14
        Case 4
            Indice = 22
        Case 5
            Indice = 13
        Case 6
            Indice = 15
        Case 7, 8
            Indice = Index + 29
        Case 9, 10
            Indice = Index + 33
        Case 11
            Indice = Index + 43
        Case 12, 13
            Indice = Index + 52
        Case 14
            Indice = 73
        Case 15
            Indice = 69
        Case 19
            Indice = 118
        Case 20
            Indice = 119
        '[Monica]25/09/2014: añadimos desde/hasta fecha de pago
        Case 21, 22
            Indice = Index + 89
        Case 23, 24
            Indice = Index + 99
        Case 25
            Indice = 121
        Case 26
            Indice = Index + 105
    
    End Select

    imgFecha(0).Tag = Indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Indice).Text <> "" Then frmC.NovaData = txtCodigo(Indice).Text
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
            '[Monica]25/09/2014: añadimos la fecha de pago
            Case 110: KEYFecha KeyAscii, 21 'fecha de pago
            Case 111: KEYFecha KeyAscii, 22 'fecha de pago
        
            '[Monica]30/12/2014: informe de recibos de consumo pendientes de cobro
            Case 124: KEYBusqueda KeyAscii, 31 ' socio desde
            Case 125: KEYBusqueda KeyAscii, 32 ' socio hasta
            Case 122: KEYFecha KeyAscii, 24 'fecha de factura
            Case 123: KEYFecha KeyAscii, 25 'fecha de factura
        
        
            '[Monica]11/09/2017: importacion de datos de monasterios
            Case 129: KEYBusqueda KeyAscii, 29 ' comunidad
            Case 130: KEYBusqueda KeyAscii, 30 ' concepto
            Case 131: KEYFecha KeyAscii, 26 'fecha de factura
        
            Case 121: KEYFecha KeyAscii, 25 'fecha de factura
        
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
    imgFecha_Click (Indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Precio As Currency

    'Quitar espacios en blanco por los lados
    '[Monica]29/07/2013: excepto en el caso de cooperativa = 8 or 9 los conceptos de recibo si me ponen un blanco lo dejamos
    If (Index = 9 Or Index = 48 Or Index = 76) And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) Then
        If txtCodigo(Index).Text <> " " Then txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    Else
        txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    End If
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1, 18, 19, 126, 127 ' Nro.hidrantes
    
        Case 10, 13, 14, 15, 16, 17, 22, 36, 37, 42, 43, 54, 64, 65, 69, 73, 114, 106, 107, 118, 119, 110, 111, 122, 123, 131, 121 'FECHAS
            If txtCodigo(Index).Text <> "" Then
                If PonerFormatoFecha(txtCodigo(Index)) Then
                End If
            End If
            
        Case 129 ' comunidad
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rtipopozos", "nompozo", "codpozo", "N")
            
        Case 130 ' concepto
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rriego", "nomriego", "codriego", "N")
        
        Case 2, 3 ' rangos de consumo
            PonerFormatoEntero txtCodigo(Index)
            
        Case 4, 5 'precios para los rangos de consumo
            PonerFormatoDecimal txtCodigo(Index), 7

        Case 6, 7, 23, 24, 34, 35, 40, 41, 56, 67, 68, 74, 75, 100, 101, 115, 116, 104, 105, 116, 117, 124, 125 'socios
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
        
            '[Monica]01/07/2013: si me dan el socio desde introducir el mismo socio hasta
            If Index = 23 And txtCodigo(Index).Text <> "" Then
                txtCodigo(24).Text = txtCodigo(23).Text
                txtNombre(24).Text = txtNombre(23).Text
            End If
        
        Case 8 ' euros/accion
            '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                PonerFormatoDecimal txtCodigo(Index), 7
            Else
                PonerFormatoDecimal txtCodigo(Index), 3
            End If
            
        Case 112 ' euros/accion
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                PonerFormatoDecimal txtCodigo(Index), 7
            Else
                PonerFormatoDecimal txtCodigo(Index), 3
            End If

        Case 70, 71 ' cuota amortizacion y de talla ordinaria
'            PonerFormatoDecimal txtcodigo(Index), 3
'            Precio = Round2((CCur(ImporteSinFormato(ComprobarCero(txtcodigo(70).Text))) + CCur(ImporteSinFormato(ComprobarCero(txtcodigo(71).Text)))) / 200, 4)
'            txtNombre(1).Text = Format(Precio, "##,##0.0000")

            If PonerFormatoDecimal(txtCodigo(70), 7) Then
                If PonerFormatoDecimal(txtCodigo(71), 7) Then
                    txtNombre(1).Text = CCur(ComprobarCero(txtCodigo(70).Text)) + CCur(ComprobarCero(txtCodigo(71).Text))
                    If CCur(txtNombre(1).Text) = 0 Then txtNombre(1).Text = ""
                    PonerFormatoDecimal txtNombre(1), 7
                Else
                    txtCodigo(71).Text = "0"
                End If
            Else
                txtCodigo(70).Text = "0"
            End If

        Case 120 ' precio de riego a manta
            PonerFormatoDecimal txtCodigo(Index), 7

        Case 72, 66 ' cuota amortizacion y de talla ordinaria
            PonerFormatoDecimal txtCodigo(Index), 7
            
            RealizarCalculos
            
        Case 53 ' bonificacion
            PonerFormatoDecimal txtCodigo(Index), 4
            If ComprobarCero(txtCodigo(53).Text) = 0 Then
                'el recargo es el siguiente campo
                PonerFoco txtCodigo(61)
            Else
                'el concepto es el siguiente campo
                PonerFoco txtCodigo(9)
            End If

        Case 61 ' recargo
            PonerFormatoDecimal txtCodigo(Index), 4

        Case 21, 26, 28, 30, 32 ' Importes de recibo de contadores
            PonerFormatoDecimal txtCodigo(Index), 3
            CalcularTotales
        
        Case 44 ' numero de etiquetas
            PonerFormatoEntero txtCodigo(Index)

        Case 51 ' lectura
            PonerFormatoEntero txtCodigo(Index)
            
        Case 52 ' nro de factura
            PonerFormatoEntero txtCodigo(Index)
            
        Case 79, 82, 85, 95, 96, 108, 109 ' zonas
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rzonas", "nomzonas", "codzonas", "N")
            
        Case 80, 81, 83, 84, 86, 87 'precios para los rangos de consumo
            PonerFormatoDecimal txtCodigo(Index), 7
            
            RealizarCalculos
    End Select
End Sub

Private Sub RealizarCalculos()
'hacemos las sumas de lo que hemos descargado
    
    ' Para las zonas en general
    If CCur(ComprobarCero(txtCodigo(72).Text)) + CCur(ComprobarCero(txtCodigo(66).Text)) <> 0 Then
        txtNombre(0).Text = CCur(ComprobarCero(txtCodigo(72).Text)) + CCur(ComprobarCero(txtCodigo(66).Text))
        PonerFormatoDecimal txtNombre(0), 7
    Else
        txtNombre(0).Text = ""
    End If
    
    If CCur(ComprobarCero(txtCodigo(80).Text)) + CCur(ComprobarCero(txtCodigo(81).Text)) <> 0 Then
        txtNombre(2).Text = CCur(ComprobarCero(txtCodigo(80).Text)) + CCur(ComprobarCero(txtCodigo(81).Text))
        PonerFormatoDecimal txtNombre(2), 7
    Else
        txtNombre(2).Text = ""
    End If
    
    If CCur(ComprobarCero(txtCodigo(83).Text)) + CCur(ComprobarCero(txtCodigo(84).Text)) <> 0 Then
        txtNombre(4).Text = CCur(ComprobarCero(txtCodigo(83).Text)) + CCur(ComprobarCero(txtCodigo(84).Text))
        PonerFormatoDecimal txtNombre(4), 7
    Else
        txtNombre(4).Text = ""
    End If
    
    If CCur(ComprobarCero(txtCodigo(86).Text)) + CCur(ComprobarCero(txtCodigo(87).Text)) <> 0 Then
        txtNombre(8).Text = CCur(ComprobarCero(txtCodigo(86).Text)) + CCur(ComprobarCero(txtCodigo(87).Text))
        PonerFormatoDecimal txtNombre(8), 7
    Else
        txtNombre(8).Text = ""
    End If

End Sub

Private Sub AbrirFrmSocios(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmZonas(Indice As Integer)
    indCodigo = Indice

    Set frmZon = New frmManZonas
    If Indice = 8 Then
        frmZon.DeConsulta = False
        frmZon.DatosADevolverBusqueda = ""
        '[Monica]07/06/2013: zonas/braçal
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            frmZon.Caption = "Braçals"
        End If
        frmZon.DeInformes = True
    Else
        frmZon.DeConsulta = True
        frmZon.DatosADevolverBusqueda = "0|1|"
    End If
    frmZon.Show vbModal
    Set frmZon = Nothing

End Sub


Private Sub AbrirFrmComunidades(Indice As Integer)
    indCodigo = Indice

    Set frmPoz = New frmPOZPozos
    frmPoz.DeConsulta = True
    frmPoz.DatosADevolverBusqueda = "0|1|"
    frmPoz.Show vbModal
    
    Set frmPoz = Nothing

End Sub


Private Sub AbrirFrmConceptos(Indice As Integer)
    
    indCodigo = Indice + 100
            
    Set frmCon = New frmBasico
    
    AyudaConceptos frmCon
    
    Set frmCon = Nothing


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
        Me.FrameComprobacionDatos.Width = 8400
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
        Me.FrameCartaTallas.Width = 8250
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
        Me.FrameReciboTalla.Width = 7530
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


Private Sub FrameAsignacionPreciosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameAsignacionPrecios.visible = visible
    If visible = True Then
        Me.FrameAsignacionPrecios.Top = -90
        Me.FrameAsignacionPrecios.Left = 0
        Me.FrameAsignacionPrecios.Height = 5025 ' 4545 '5925
        Me.FrameAsignacionPrecios.Width = 7395 '6645
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
        Me.FrameReciboMantenimiento.Height = 7505
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


Private Sub FrameImporLecturasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameImporLecturas.visible = visible
    If visible = True Then
        Me.FrameImporLecturas.Top = -90
        Me.FrameImporLecturas.Left = 0
        Me.FrameImporLecturas.Height = 4725
        Me.FrameImporLecturas.Width = 6675
        W = Me.FrameImporLecturas.Width
        H = Me.FrameImporLecturas.Height
    End If
End Sub


Private Sub FrameExporLecturasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameExporLecturas.visible = visible
    If visible = True Then
        Me.FrameExporLecturas.Top = -90
        Me.FrameExporLecturas.Left = 0
        Me.FrameExporLecturas.Height = 3735
        Me.FrameExporLecturas.Width = 7350
        W = Me.FrameExporLecturas.Width
        H = Me.FrameExporLecturas.Height
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
        Me.FrameRectificacion.Width = 7035
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

Dim B As Boolean
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
                Mens = "Proceso Facturación Consumo: " & vbCrLf & vbCrLf
                If vParamAplic.Cooperativa = 7 Then ' QUATRETONDA
                    B = FacturacionConsumoQUATRETONDA(nTabla, cadSelect, FecFac, Me.Pb1, Mens, Consumo, EsRectificativa)
                Else ' MALLAES
                    B = FacturacionConsumo(nTabla, cadSelect, FecFac, Me.Pb1, Mens)
                End If
                If B Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                   
                    'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                    If Me.Check1(2).Value Then
                        cadFormula = ""
                        CadParam = CadParam & "pFecFac= """ & txtCodigo(14).Text & """|"
                        numParam = numParam + 1
                        CadParam = CadParam & "pTitulo= ""Resumen Facturación de Contadores""|"
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
                        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
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


Private Function FacturacionConsumo(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rpozos.codsocio,hidrante,nroorden,codparti,poligono,parcelas,hanegada,lect_ant,lect_act,fech_ant,fech_act"
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by rpozos.codsocio, rpozos.hidrante "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    B = True
    
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
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
    
    
    While Not Rs.EOF And B
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
        
            If ConsumoHan < CLng(txtCodigo(3).Text) Then
                If ConsumoHan < CLng(txtCodigo(2).Text) Then
                    Consumo1 = DBLet(Rs2!Consumo, "N")
                    Consumo2 = 0
                Else
                    Consumo1 = CLng(txtCodigo(2).Text) * (Acciones / 30 * DBLet(Rs2!Dias, "N"))
                    Consumo2 = DBLet(Rs2!Consumo) - Consumo1
                End If
            End If
            
            Set Rs2 = Nothing
            
            '[Monica]28/10/2011: añadido el recalculo de tramos de los contadores de la factura
            Sql = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
            
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
            
                TotalFac = Round2(vConsumo1 * CCur(ImporteSinFormato(txtCodigo(4).Text)), 2) + _
                           Round2(vConsumo2 * CCur(ImporteSinFormato(txtCodigo(5).Text)), 2) + _
                           vParamAplic.CuotaPOZ
            
                Sql = "update rrecibpozos set consumo1 = " & DBSet(vConsumo1, "N") & ", consumo2 = " & DBSet(vConsumo2, "N")
                Sql = Sql & ", baseimpo = " & DBSet(TotalFac, "N") & ", totalfact = " & DBSet(TotalFac, "N")
                Sql = Sql & " where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
                Sql = Sql & " and numlinea = " & DBSet(RsFacturas!numlinea, "N")
                
                conn.Execute Sql
            
                RsFacturas.MoveNext
            Wend
            
            Set RsFacturas = Nothing
            
            
            Sql = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
            
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
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
           
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
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
        
        TotalFac = Round2(ConsTra1 * CCur(ImporteSinFormato(txtCodigo(4).Text)), 2) + _
                   Round2(ConsTra2 * CCur(ImporteSinFormato(txtCodigo(5).Text)), 2) + _
                   vParamAplic.CuotaPOZ
    
        IncrementarProgresNew Pb1, 1
        
        NumLin = NumLin + 1
        
        DiferenciaDias = DBLet(Rs!fech_act, "F") - DBLet(Rs!fech_ant, "F")
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, numlinea, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, difdias) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(NumLin, "N") & "," & DBSet(ActSocio, "N") & ","
        Sql = Sql & DBSet(Rs!Hidrante, "T") & "," & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(ConsumoHidrante, "N") & "," & DBSet(vParamAplic.CuotaPOZ, "N") & ","
        Sql = Sql & DBSet(Rs!lect_ant, "N") & "," & DBSet(Rs!fech_ant, "F") & ","
        Sql = Sql & DBSet(Rs!lect_act, "N") & "," & DBSet(Rs!fech_act, "F") & ","
        Sql = Sql & DBSet(ConsTra1, "N") & "," & DBSet(ImporteSinFormato(txtCodigo(4).Text), "N") & ","
        Sql = Sql & DBSet(ConsTra2, "N") & "," & DBSet(ImporteSinFormato(txtCodigo(5).Text), "N") & ","
        Sql = Sql & "'Recibo de Consumo',0,"
        Sql = Sql & DBSet(DiferenciaDias, "N") & ")"
        
        conn.Execute Sql
        
        '
        '[Monica]21/10/2011: insertamos las distintas fases(acciones) del socio en la facturacion
        '
        Sql = "insert into rrecibpozos_acc(codtipom,numfactu,fecfactu,numlinea,numfases,acciones,observac) "
        Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
        Sql = Sql & DBSet(NumLin, "N") & ", numfases, acciones, observac from rsocios_pozos where codsocio = " & DBSet(ActSocio, "N")
        
        conn.Execute Sql
            
            
        ' actualizar en los acumulados de hidrantes
        Sql = "update rpozos set acumconsumo = acumconsumo + " & DBSet(ConsumoHidrante, "N")
        Sql = Sql & ", acumcuota = acumcuota + " & DBSet(vParamAplic.CuotaPOZ, "N")
        
'        Sql = Sql & ", lect_ant = lect_act "
'        Sql = Sql & ", fech_ant = fech_act "
'        Sql = Sql & ", consumo = 0 "
        
        
        Sql = Sql & " WHERE hidrante = " & DBSet(Rs!Hidrante, "T")
        
        conn.Execute Sql
            
            
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
    
        If ConsumoHan < CLng(txtCodigo(3).Text) Then
            If ConsumoHan < CLng(txtCodigo(2).Text) Then
                Consumo1 = DBLet(Rs2!Consumo, "N")
                Consumo2 = 0
            Else
                Consumo1 = CLng(txtCodigo(2).Text) * (Acciones / 30 * DBLet(Rs2!Dias, "N"))
                Consumo2 = DBLet(Rs2!Consumo) - Consumo1
            End If
        End If
        
        Set Rs2 = Nothing
    
    
        '[Monica]28/10/2011: añadido el recalculo de tramos de los contadores de la factura
        Sql = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
        
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
        
            TotalFac = Round2(vConsumo1 * CCur(ImporteSinFormato(txtCodigo(4).Text)), 2) + _
                       Round2(vConsumo2 * CCur(ImporteSinFormato(txtCodigo(5).Text)), 2) + _
                       vParamAplic.CuotaPOZ
        
            Sql = "update rrecibpozos set consumo1 = " & DBSet(vConsumo1, "N") & ", consumo2 = " & DBSet(vConsumo2, "N")
            Sql = Sql & ", baseimpo = " & DBSet(TotalFac, "N") & ", totalfact = " & DBSet(TotalFac, "N")
            Sql = Sql & " where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
            Sql = Sql & " and numlinea = " & DBSet(RsFacturas!numlinea, "N")
            
            conn.Execute Sql
        
            RsFacturas.MoveNext
        Wend
        
        Set RsFacturas = Nothing
        
        
        Sql = "select * from rrecibpozos where codtipom = 'RCP' and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
        
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

        B = InsertResumen(tipoMov, CStr(numfactu))
        If B And HayReg Then B = vTipoMov.IncrementarContador(tipoMov)
    
    End If
    
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionConsumo = False
    Else
        conn.CommitTrans
        FacturacionConsumo = True
    End If
End Function

Private Function FacturacionConsumoQUATRETONDA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String, ConsumoRectif As Long, EsRectificativa As Boolean) As Boolean
Dim Sql As String
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
Dim B As Boolean
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

Dim vSocio As cSocio


    On Error GoTo eFacturacion

    FacturacionConsumoQUATRETONDA = False
    
    tipoMov = "RCP"
    
    conn.BeginTrans
    
    If EsRectificativa Then
        Sql = "update rpozos set consumo = " & DBSet(ConsumoRectif, "N")
        Sql = Sql & ", lect_act = " & DBSet(txtCodigo(51).Text, "N")
        Sql = Sql & ", fech_act = " & DBSet(txtCodigo(54).Text, "F")
        Sql = Sql & " where " & cWhere

        conn.Execute Sql
    End If
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rpozos.codsocio,hidrante,nroorden,codparti,poligono,parcelas,hanegada,lect_ant,lect_act,fech_ant,fech_act,nroacciones,codpozo,consumo,calibre "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by rpozos.codsocio, rpozos.hidrante "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    B = True
    
    
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
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))

        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        
'[Monica]01/02/2016: Introducimos las facturas internas
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Rs!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.SeccionPOZOS) Then
                vPorcIva = ""
                If vSocio.EsFactADVInt Then
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSeccion.TipIvaExento, "N")
                    tipoMov = "FIN"
                    If vPorcIva = "" Then
                        MsgBox "No se ha encontrado el tipo de Iva " & vSeccion.TipIvaExento & ". Revise.", vbExclamation
                        conn.RollbackTrans
                        Exit Function
                    End If
                Else
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
                    tipoMov = "RCP"
                    If vPorcIva = "" Then
                        MsgBox "No se ha encontrado el tipo de Iva " & vParamAplic.CodIvaPOZ & ". Revise.", vbExclamation
                        conn.RollbackTrans
                        Exit Function
                    End If
                End If
            End If
            PorcIva = CCur(ImporteSinFormato(vPorcIva))
        End If
'hasta aquí

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
    
    
    While Not Rs.EOF And B
        HayReg = True
        
        ActSocio = Rs!Codsocio
        
        If ActSocio <> AntSocio Then
            
            AntSocio = ActSocio
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
           
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
        '[Monica]01/02/2016: Introducimos las facturas internas
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(Rs!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.SeccionPOZOS) Then
                    vPorcIva = ""
                    If vSocio.EsFactADVInt Then
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSeccion.TipIvaExento, "N")
                        tipoMov = "FIN"
                        If vPorcIva = "" Then
                            MsgBox "No se ha encontrado el tipo de Iva " & vSeccion.TipIvaExento & ". Revise.", vbExclamation
                            conn.RollbackTrans
                            Exit Function
                        End If
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
                        tipoMov = "RCP"
                        If vPorcIva = "" Then
                            MsgBox "No se ha encontrado el tipo de Iva " & vParamAplic.CodIvaPOZ & ". Revise.", vbExclamation
                            conn.RollbackTrans
                            Exit Function
                        End If
                    End If
                End If
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
            End If
        'hasta aquí
            
            
            
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
            
        Sql2 = "select precio1, imporcuota, imporcuotahda from rtipopozos where codpozo = " & DBSet(Rs!codpozo, "N")
        
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
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, numlinea, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, "
        '[Monica]28/02/2012: introducimos los nuevos campos
        Sql = Sql & "codparti, calibre, codpozo) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(NumLin, "N") & "," & DBSet(ActSocio, "N") & ","
        Sql = Sql & DBSet(Rs!Hidrante, "T") & "," & DBSet(baseimpo, "N") & ","
        '[Monica]01/02/2016: Introducimos las facturas internas
        If Not vSocio.EsFactADVInt Then
            Sql = Sql & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(vPorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Else
            Sql = Sql & DBSet(vParamAplic.CodIvaExeADV, "N") & "," & DBSet(vPorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        End If
        Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!Consumo, "N") & "," & DBSet(ImpCuota, "N") & ","
        Sql = Sql & DBSet(Rs!lect_ant, "N") & "," & DBSet(Rs!fech_ant, "F") & ","
        Sql = Sql & DBSet(Rs!lect_act, "N") & "," & DBSet(Rs!fech_act, "F") & ","
        Sql = Sql & DBSet(Rs!Consumo, "N") & "," & DBSet(Precio1, "N") & "," ' consumo
        Sql = Sql & DBSet(Acciones, "N") & "," & DBSet(CuotaHda, "N") & ","  ' mantenimiento
        Sql = Sql & DBSet(txtCodigo(48).Text, "T") & ",0,"
        '[Monica]28/02/2012: introducimos los nuevos campos: partida,calibre y codpozo
        Sql = Sql & DBSet(Rs!codparti, "N") & "," & DBSet(Rs!Calibre, "N") & "," & DBSet(Rs!codpozo, "N") & ")"
        
        conn.Execute Sql
            
        ' actualizar en los acumulados de hidrantes
        Sql = "update rpozos set acumconsumo = acumconsumo + " & DBSet(Rs!Consumo, "N")
        Sql = Sql & ", acumcuota = acumcuota + " & DBSet(ImpCuota, "N")
        
        Sql = Sql & ", lect_ant = lect_act "
        Sql = Sql & ", fech_ant = fech_act "
        Sql = Sql & ", lect_act = null "
        Sql = Sql & ", fech_act = null "
        Sql = Sql & ", consumo = 0 "
        
        Sql = Sql & " WHERE hidrante = " & DBSet(Rs!Hidrante, "T")
        
        conn.Execute Sql
        
        Rs.MoveNext
    Wend
    
    If HayReg Then B = InsertResumen(tipoMov, CStr(numfactu))
    If B And HayReg Then B = vTipoMov.IncrementarContador(tipoMov)
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionConsumoQUATRETONDA = False
    Else
        conn.CommitTrans
        FacturacionConsumoQUATRETONDA = True
    End If
End Function



Private Function TotalFacturasSocios(cTabla As String, cWhere As String) As Long
Dim Sql As String

    TotalFacturasSocios = 0
    
    Sql = "SELECT  count(distinct rpozos.codsocio) "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If

    TotalFacturasSocios = TotalRegistros(Sql)

End Function

Private Function TotalFacturasHidrante(cTabla As String, cWhere As String) As Long
Dim Sql As String

    TotalFacturasHidrante = 0
    
    Sql = "SELECT  count(distinct rpozos.hidrante) "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If

    TotalFacturasHidrante = TotalRegistros(Sql)

End Function



Private Function DatosOK() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean
Dim Sql As String
Dim FecFac As Date
Dim FecUlt As Date
Dim vSeccion As CSeccion

    On Error GoTo EDatosOK

    DatosOK = False
    B = True
    Select Case OpcionListado
        Case 3 ' generacion de recibos de consumo
            If txtCodigo(14).Text = "" Then
                MsgBox "Debe introducir un valor para la Fecha de Recibo.", vbExclamation
                PonerFoco txtCodigo(14)
                B = False
            End If
            If B Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
                    If vSeccion.AbrirConta Then
                        '[Monica]20/06/2017: control de fechas que antes no estaba
                        ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(14)))
                        If ResultadoFechaContaOK > 0 Then
                            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                            B = False
                        End If
                    End If
                End If
            
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If
            
            
            If B Then
                If txtCodigo(2).Text = "" Or txtCodigo(3).Text = "" Or txtCodigo(4).Text = "" Or txtCodigo(5).Text = "" Then
                    MsgBox "Debe introducir valores en rangos y precios de los tramos.", vbExclamation
                    PonerFoco txtCodigo(2)
                    B = False
                End If
            End If
            '[Monica]29/05/2013: Solo para escalona y utxera obligamos a escribir el concepto o poner un blanco.
            If B Then
                If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                    If Len(txtCodigo(48).Text) = 0 Then
                        MsgBox "Debe introducir un valor en el concepto.", vbExclamation
                        PonerFoco txtCodigo(48)
                        B = False
                    End If
                End If
            End If
        Case 4 ' generacion de recibos de mantenimiento
            If txtCodigo(10).Text = "" Then
                MsgBox "Debe introducir un valor para la Fecha de Recibo.", vbExclamation
                PonerFoco txtCodigo(10)
                B = False
            End If
            If B Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
                    If vSeccion.AbrirConta Then
                        '[Monica]20/06/2017: control de fechas que antes no estaba
                        ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(10)))
                        If ResultadoFechaContaOK > 0 Then
                            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                            B = False
                        End If
                    End If
                End If
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If
            
            If B Then
                If txtCodigo(8).Text = "" Then
                    MsgBox "Debe introducir un valor en Euros/Acción.", vbExclamation
                    PonerFoco txtCodigo(8)
                    B = False
                End If
            End If
            If B Then
                If (txtCodigo(9).Text = "" And vParamAplic.Cooperativa <> 8 And vParamAplic.Cooperativa <> 10) Or (Len(txtCodigo(9).Text) = 0 And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)) Then
                    MsgBox "Debe introducir un valor en el concepto", vbExclamation
                    PonerFoco txtCodigo(9)
                    B = False
                End If
            End If
            
            'o metemos una bonificacion o un recargo o nada, pero no ambos a la vez
            If B Then
                If ComprobarCero(txtCodigo(53).Text) <> 0 And ComprobarCero(txtCodigo(61).Text) <> 0 Then
                    MsgBox "No se permite introducir a la vez una Bonificacion y un Recargo. Revise.", vbExclamation
                    PonerFoco txtCodigo(53)
                    B = False
                End If
            End If
            
        Case 5 ' generacion de recibos de contadores
            If txtCodigo(22).Text = "" Then
                MsgBox "Debe introducir un valor para la Fecha de Recibo.", vbExclamation
                PonerFoco txtCodigo(22)
                B = False
            End If
            If B Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
                    If vSeccion.AbrirConta Then
                        '[Monica]20/06/2017: control de fechas que antes no estaba
                        ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(22)))
                        If ResultadoFechaContaOK > 0 Then
                            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                            B = False
                        End If
                    End If
                End If
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If
            
            
            If B Then
                If txtCodigo(21).Text <> "" And txtCodigo(20).Text = "" Then
                    MsgBox "Si introduce un Importe para Mano de Obra, debe introducir un Concepto.", vbExclamation
                    PonerFoco txtCodigo(20)
                    B = False
                End If
            End If
            If B Then
                If txtCodigo(26).Text <> "" And txtCodigo(25).Text = "" Then
                    MsgBox "Si introduce un Importe para el Artículo 1, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtCodigo(25)
                    B = False
                End If
            End If
            If B Then
                If txtCodigo(28).Text <> "" And txtCodigo(27).Text = "" Then
                    MsgBox "Si introduce un Importe para el Artículo 2, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtCodigo(27)
                    B = False
                End If
            End If
            If B Then
                If txtCodigo(30).Text <> "" And txtCodigo(29).Text = "" Then
                    MsgBox "Si introduce un Importe para el Artículo 3, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtCodigo(29)
                    B = False
                End If
            End If
            If B Then
                If txtCodigo(32).Text <> "" And txtCodigo(31).Text = "" Then
                    MsgBox "Si introduce un Importe para el Artículo 4, debe introducir un Concepto correspondiente.", vbExclamation
                    PonerFoco txtCodigo(31)
                    B = False
                End If
            End If
            
            If B Then
                If txtCodigo(33).Text = "" Then
                    MsgBox "El Recibo debe de ser de un valor distinto de cero. Revise."
                    PonerFoco txtCodigo(20)
                    B = False
                End If
            End If
    
        Case 8 ' etiquetas de contadores
            If txtCodigo(44).Text = 0 Then
                MsgBox "El número de etiquetas debe ser superior a 0. Revise."
                PonerFoco txtCodigo(44)
                B = False
            End If
        
            If B Then
                If Trim(txtCodigo(45).Text) = "" And Trim(txtCodigo(46).Text) = "" And Trim(txtCodigo(47).Text) = "" Then
                    MsgBox "Debe haber algún valor en alguna de las Líneas. Revise."
                    PonerFoco txtCodigo(45)
                    B = False
                End If
            End If
            
        Case 9 ' Rectificacion de Lecturas
            If txtCodigo(52).Text = "" Then
                MsgBox "Debe introducir un Nº de Factura. Revise", vbExclamation
                PonerFoco txtCodigo(52)
                B = False
            End If
            If B And txtCodigo(56).Text = "" Then
                MsgBox "Debe introducir el Socio de la Factura. Revise", vbExclamation
                PonerFoco txtCodigo(56)
                B = False
            End If
            If B And txtCodigo(55).Text = "" Then
                MsgBox "Debe introducir el Hidrante de la Factura. Revise", vbExclamation
                PonerFoco txtCodigo(55)
                B = False
            End If
            If B And txtCodigo(54).Text = "" Then
                MsgBox "Debe introducir la Fecha de la Factura. Revise", vbExclamation
                PonerFoco txtCodigo(54)
                B = False
            End If
            If B Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
                    If vSeccion.AbrirConta Then
                        '[Monica]20/06/2017: control de fechas que antes no estaba
                        ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(54)))
                        If ResultadoFechaContaOK > 0 Then
                            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                            B = False
                        End If
                    End If
                End If
            
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If
            
            If B And txtCodigo(51).Text = "" Then
                MsgBox "Debe introducir cual es la lectura actual. Revise", vbExclamation
                PonerFoco txtCodigo(51)
                B = False
            End If
            If B Then
                Sql = "select count(*) from rrecibpozos where codtipom = " & DBSet(Mid(Combo1(2).Text, 1, 3), "T")
                Sql = Sql & " and numfactu = " & DBSet(txtCodigo(52).Text, "N")
                Sql = Sql & " and codsocio = " & DBSet(txtCodigo(56).Text, "N")
                Sql = Sql & " and hidrante = " & DBSet(txtCodigo(55).Text, "T")
                If TotalRegistros(Sql) = 0 Then
                    MsgBox "No existe ninguna factura con estos datos para rectificar. Revise.", vbExclamation
                    PonerFoco txtCodigo(52)
                    B = False
                Else
                    ' miramos si es la ultima factura de ese hidrante
                    ' en este caso no debemos hacer la rectificativa porque dejariamos el hidrante con las
                    ' lecturas incorrectas
                    Sql = "select max(fecfactu) from rrecibpozos where codtipom = " & DBSet(Mid(Combo1(2).Text, 1, 3), "T")
                    Sql = Sql & " and hidrante = " & DBSet(txtCodigo(55).Text, "T")
                    FecUlt = DevuelveValor(Sql)
                    
                    Sql = "select fecfactu from rrecibpozos where codtipom = " & DBSet(Mid(Combo1(2).Text, 1, 3), "T")
                    Sql = Sql & " and numfactu= " & DBSet(txtCodigo(52).Text, "N")
                    Sql = Sql & " and hidrante = " & DBSet(txtCodigo(55).Text, "T")
                    FecFac = DevuelveValor(Sql)
                    
                    If FecUlt > FecFac Then
                        MsgBox "Existe un factura de fecha posterior sobre este hidrante, no se permite el proceso. Revise.", vbExclamation
                        PonerFoco txtCodigo(52)
                        B = False
                    End If
                    
                    If B Then
                        If CDate(txtCodigo(54).Text) < FecUlt Then
                            MsgBox "la fecha de la factura rectificativa es inferior a la que rectifica. Revise.", vbExclamation
                            PonerFoco txtCodigo(52)
                            B = False
                        End If
                    End If
                    ' comprobaciones contables
                End If
            End If
            
            
        Case 10 ' Carta de Tallas a socios
            If txtCodigo(69).Text = "" Then
                MsgBox "Debe introducir la fecha de recibo. Revise.", vbExclamation
                PonerFoco txtCodigo(69)
                B = False
            End If
            
            If B Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
                    If vSeccion.AbrirConta Then
                        '[Monica]20/06/2017: control de fechas que antes no estaba
                        ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(69)))
                        If ResultadoFechaContaOK > 0 Then
                            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                            B = False
                        End If
                    End If
                End If
            
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If
        Case 11, 12 ' generacion y a actualizacion de recibos de talla para Escalona
            If txtCodigo(73).Text = "" Then
                MsgBox "Debe introducir la fecha de recibo. Revise.", vbExclamation
                PonerFoco txtCodigo(73)
                B = False
            End If
            
            If B Then
                If OpcionListado = 11 Then
                    Set vSeccion = New CSeccion
                    If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
                        If vSeccion.AbrirConta Then
                            '[Monica]20/06/2017: control de fechas que antes no estaba
                            ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(73)))
                            If ResultadoFechaContaOK > 0 Then
                                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                                B = False
                            End If
                        End If
                    End If
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                Else
                    '[Monica]20/06/2017: control de fechas que antes no estaba
                    ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(73)))
                    If ResultadoFechaContaOK > 0 Then
                        If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                        B = False
                    End If
                End If
            End If
            
            
            
            '[Monica]29/05/2013: Solo para escalona y utxera obligamos a escribir el concepto o poner un blanco.
            If B Then
                '[Monica]13/03/2014: añadimos la condicion de opcionlistado = 11 pq sino pedia un concepto
                '                    en la bonificacion no pedimos concepto
                If (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) And OpcionListado = 11 Then
                    If Len(txtCodigo(76).Text) = 0 Then
                        MsgBox "Debe introducir un valor en el concepto.", vbExclamation
                        PonerFoco txtCodigo(76)
                        B = False
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
                If B Then
                    If ComprobarCero(txtCodigo(78).Text) <> 0 And ComprobarCero(txtCodigo(77).Text) <> 0 Then
                        MsgBox "No se permite introducir a la vez una Bonificacion y un Recargo. Revise.", vbExclamation
                        PonerFoco txtCodigo(78)
                        B = False
                    End If
                    If B And ComprobarCero(txtCodigo(78).Text) = 0 And ComprobarCero(txtCodigo(77).Text) = 0 Then
                        MsgBox "Debe introducir un porcentaje de Bonificacion o de Recargo. Revise.", vbExclamation
                        PonerFoco txtCodigo(78)
                        B = False
                    
                    End If
            
                End If
            End If
    
        Case 17 ' generacion de recibos de consumo a manta
            If txtCodigo(115).Text = "" Then
                MsgBox "Debe introducir un valor para el Socio. Reintroduzca.", vbExclamation
                PonerFoco txtCodigo(115)
                B = False
            End If
            
            If txtCodigo(114).Text = "" Then
                MsgBox "Debe introducir un valor para la Fecha del Ticket.", vbExclamation
                PonerFoco txtCodigo(114)
                B = False
            End If
            If B Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
                    If vSeccion.AbrirConta Then
                        '[Monica]20/06/2017: control de fechas que antes no estaba
                        ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(114)))
                        If ResultadoFechaContaOK > 0 Then
                            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                            B = False
                        End If
                    End If
                End If
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If
            
            
            
            If B Then
                If txtCodigo(112).Text = "" Then
                    MsgBox "Debe introducir un valor en Euros/Acción.", vbExclamation
                    PonerFoco txtCodigo(112)
                    B = False
                End If
            End If
            If B Then
                If (txtCodigo(113).Text = "" And vParamAplic.Cooperativa <> 8 And vParamAplic.Cooperativa <> 10) Or (Len(txtCodigo(113).Text) = 0 And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)) Then
                    MsgBox "Debe introducir un valor en el concepto", vbExclamation
                    PonerFoco txtCodigo(113)
                    B = False
                End If
            End If
            
        Case 18 ' informe de recibos pendientes de cobro por braçal
            If B Then ' socio
                If txtCodigo(104).Text <> "" And txtCodigo(105).Text <> "" Then
                    If CLng(txtCodigo(104).Text) > CLng(txtCodigo(105).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtCodigo(104)
                        B = False
                    End If
                End If
            End If
            If B Then 'fecha
                If txtCodigo(106).Text <> "" And txtCodigo(107).Text <> "" Then
                    If CDate(txtCodigo(106).Text) > CDate(txtCodigo(107).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtCodigo(106)
                        B = False
                    End If
                End If
            End If
            If B Then 'zona
                If txtCodigo(108).Text <> "" And txtCodigo(109).Text <> "" Then
                    If CLng(txtCodigo(108).Text) > CLng(txtCodigo(109).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtCodigo(108)
                        B = False
                    End If
                End If
            End If
            If B Then 'sector
                If txtCodigo(102).Text <> "" And txtCodigo(103).Text <> "" Then
                    If CLng(txtCodigo(102).Text) > CLng(txtCodigo(103).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtCodigo(102)
                        B = False
                    End If
                End If
            End If
            
        Case 20 ' informe de recibos de consumo pendientes de cobro
            If B Then ' socio
                If txtCodigo(124).Text <> "" And txtCodigo(125).Text <> "" Then
                    If CLng(txtCodigo(124).Text) > CLng(txtCodigo(125).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtCodigo(124)
                        B = False
                    End If
                End If
            End If
            If B Then 'fecha
                If txtCodigo(122).Text <> "" And txtCodigo(123).Text <> "" Then
                    If CDate(txtCodigo(122).Text) > CDate(txtCodigo(123).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtCodigo(122)
                        B = False
                    End If
                End If
            End If
            If B Then 'hidrante
                If txtCodigo(126).Text <> "" And txtCodigo(127).Text <> "" Then
                    If CLng(txtCodigo(126).Text) > CLng(txtCodigo(127).Text) Then
                        MsgBox "El campo Desde no puede ser superior al Hasta", vbExclamation
                        PonerFoco txtCodigo(126)
                        B = False
                    End If
                End If
            End If
            
        Case 21 ' importacion de datos de monasterios
            ' fecha
            If B Then
                If txtCodigo(131).Text = "" Then
                    MsgBox "Debe introducir una fecha. Revise.", vbExclamation
                    PonerFoco txtCodigo(131)
                    B = False
                End If
            End If
            ' comunidad
            If B Then
                If txtNombre(129).Text = "" Then
                    MsgBox "Debe introducir una comunidad. Revise.", vbExclamation
                    PonerFoco txtCodigo(129)
                    B = False
                End If
            End If
            ' concepto
            If B Then
                If txtNombre(130).Text = "" Then
                    MsgBox "Debe introducir un concepto. Revise.", vbExclamation
                    PonerFoco txtCodigo(130)
                    B = False
                End If
            End If
            
            '[Monica]19/09/2017: añadimos esta condicion
            ' comprobamos que no haya fech_ant con fecha superior a la introducida
            If B Then
                If CDate(FechaAnteriorMaxima) > CDate(txtCodigo(131).Text) Then
                    If MsgBox("Hay lecturas posteriores a la fecha introducida. " & vbCrLf & vbCrLf & " ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        B = True
                    Else
                        B = False
                        PonerFoco txtCodigo(131)
                    End If
                End If
            End If
            
            
            ' comprobamos que no hayan contadores con la fecha introducida
            If B Then
                Sql = "select count(*) from rpozos where not fech_act is null "
                If TotalRegistros(Sql) <> 0 Then
                    If MsgBox("Hay lecturas en curso. " & vbCrLf & vbCrLf & " ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        B = True
                    Else
                        B = False
                        PonerFoco txtCodigo(131)
                    End If
                End If
            End If
            
        Case 22 ' exportacion de datos de monasterios
            ' fecha
            If B Then
                If txtCodigo(121).Text = "" Then
                    MsgBox "Debe introducir una fecha. Revise.", vbExclamation
                    PonerFoco txtCodigo(121)
                    B = False
                End If
            End If
            
            ' comprobamos que no hayan facturas con la fecha de factura
            If B Then
                Sql = "select count(*) from rrecibpozos where fecfactu = " & DBSet(txtCodigo(121).Text, "F") & " and codtipom = 'RCP'"
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Hay recibos que tienen la misma fecha de factura. Revise.", vbExclamation
                    PonerFoco txtCodigo(121)
                    B = False
                End If
            End If
            
    End Select
    
    DatosOK = B
    
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function FechaAnteriorMaxima() As String
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Sql = "select max(fech_ant) from rpozos where not fech_ant is null "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    FechaAnteriorMaxima = "01/01/1900"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0).Value) Then FechaAnteriorMaxima = DBLet(Rs.Fields(0).Value, "F") 'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        FechaAnteriorMaxima = "01/01/1900"
        Err.Clear
    End If

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

Dim B As Boolean
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
        
        Mens = "Proceso Facturación Mantenimiento: " & vbCrLf & vbCrLf
        B = FacturacionMantenimiento(nTabla, cadSelect, txtCodigo(10).Text, Me.Pb2, Mens)
        If B Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtCodigo(10).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturación Mantenimiento""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtCodigo(10).Text)
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
                FecFac = CDate(txtCodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
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
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim B As Boolean
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
        
        Mens = "Proceso Facturación Mantenimiento: " & vbCrLf & vbCrLf
        B = FacturacionMantenimientoUTXERA(nTabla, cadSelect, txtCodigo(10).Text, Me.Pb2, Mens)
        If B Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtCodigo(10).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturación Mantenimiento""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtCodigo(10).Text)
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
'                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(1) & "])"
'                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
'                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
'                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                '[Monica]06/03/2013: solo lo facturado
                cadAux = "{rrecibpozos.codtipom} = 'RMP'"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub


                'Fecha de Factura
                FecFac = CDate(txtCodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
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
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim B As Boolean
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
        
        Mens = "Proceso Facturación Mantenimiento: " & vbCrLf & vbCrLf
        B = FacturacionMantenimientoESCALONA(nTabla, cadSelect, txtCodigo(10).Text, Me.Pb2, Mens)
        If B Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(1).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtCodigo(10).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturación Mantenimiento""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtCodigo(10).Text)
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
'                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(1) & "])"
'                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
'                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
'                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                '[Monica]06/03/2013: solo lo facturado
                cadAux = "{rrecibpozos.codtipom} = 'RMP'"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                'Fecha de Factura
                FecFac = CDate(txtCodigo(10).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
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
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim B As Boolean
Dim Sql2 As String
Dim cadena As String

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
        Me.Pb4.visible = True
        Me.Pb4.Max = Nregs
        Me.Pb4.Value = 0
        Me.Label2(78).visible = True
        Me.Refresh
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        If OpcionListado = 11 Then
            LOG.Insertar 8, vUsu, "Facturacion Talla Recibos Pozos: " & vbCrLf & nTabla & vbCrLf & cadSelect
        Else
            If CCur(ComprobarCero(txtCodigo(78).Text)) <> 0 Then
                cadena = "Bonificacion: " & CCur(ImporteSinFormato(txtCodigo(78).Text)) & "%"
            Else
                cadena = "Recargo: " & CCur(ImporteSinFormato(txtCodigo(77).Text)) & "%"
            End If
        
            LOG.Insertar 8, vUsu, "Actualización Recibos Talla Pozos: " & vbCrLf & cadena & vbCrLf & cadSelect
        End If
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        Mens = "Proceso Facturación Talla: " & vbCrLf & vbCrLf
        If OpcionListado = 11 Then
            B = FacturacionTallaESCALONA(nTabla, cadSelect, txtCodigo(73).Text, Me.Pb4, Mens)
        Else
            Me.Label2(78).Caption = "Comprobando recibos ..."
            Me.Refresh
            If Not HayFactContabilizadas(nTabla, cadSelect) Then
                Me.Label2(78).Caption = "Actualizando recibos ..."
                Me.Refresh
                B = ActualizacionTallaESCALONA(nTabla, cadSelect, txtCodigo(73).Text, Me.Pb4, Mens)
            Else
                Me.Pb4.visible = False
                Me.Label2(78).visible = False
                DoEvents
                Exit Sub
            End If
        End If
        If B Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(7).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtCodigo(73).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturación Talla""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtCodigo(73).Text)
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
'                'Nº Factura
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
'                cadTitulo = "Reimpresión de Facturas de Talla"
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

Dim B As Boolean
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
        Me.pb7.visible = True
        Me.pb7.Max = Nregs
        Me.pb7.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturación Mantenimiento: " & vbCrLf & vbCrLf
        B = FacturacionConsumoMantaESCALONA(nTabla, cadSelect, txtCodigo(114).Text, Me.pb7, Mens)
        If B Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
            If Me.Check1(9).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtCodigo(114).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturación Consumo a Manta""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtCodigo(114).Text)
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
                'Nº Factura
                cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(4) & "])"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                '[Monica]06/03/2013: solo lo facturado
                cadAux = "{rrecibpozos.codtipom} = 'RMT'"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                'Fecha de Factura
                FecFac = CDate(txtCodigo(114).Text)
                cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = Replace(nomDocu, "Mto.", "Manta.")
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresión Facturas Consumo a Manta"
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

Dim B As Boolean
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
        Me.pb7.visible = True
        Me.pb7.Max = Nregs
        Me.pb7.Value = 0
        Me.Refresh
        
        Mens = "Proceso Facturación Mantenimiento: " & vbCrLf & vbCrLf
        B = FacturacionConsumoMantaESCALONANew(nTabla, cadSelect, txtCodigo(114).Text, Me.pb7, Mens)
        If B Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
            If Me.Check1(10).Value Then
                cadFormula = ""
                cadSelect = ""
                'Nº Factura
                cadAux = "({rpozticketsmanta.numalbar} IN [" & FacturasGeneradasPOZOS(5) & "])"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                'Fecha de Ticket
                FecFac = CDate(txtCodigo(114).Text)
                cadAux = "{rpozticketsmanta.fecalbar}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rpozticketsmanta.fecalbar}='" & Format(FecFac, FormatoFecha) & "'"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                
                '[Monica]02/09/2014: antes escontado
                If EsSocioContadoPOZOS(txtCodigo(115).Text) Then
                    'Nombre fichero .rpt a Imprimir
                    cadNombreRPT = Replace(nomDocu, "Mto.", "TicketMantaCont.")
                Else
                    'Nombre fichero .rpt a Imprimir
                    cadNombreRPT = Replace(nomDocu, "Mto.", "TicketManta.")
                End If
                'Nombre fichero .rpt a Imprimir
                cadTitulo = "Reimpresión Tickets Consumo a Manta"
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



Private Function HayFactContabilizadas(tabla As String, cSelect As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Contabiliz As Boolean
Dim LEtra As String
Dim EstaEnTesoreria As String
Dim numasien As String

    On Error GoTo eHayFactContabilizadas

    Screen.MousePointer = vbHourglass

    Sql = "SELECT rrecibpozos.* "
    Sql = Sql & " FROM  " & tabla

    If cSelect <> "" Then
        cSelect = QuitarCaracterACadena(cSelect, "{")
        cSelect = QuitarCaracterACadena(cSelect, "}")
        cSelect = QuitarCaracterACadena(cSelect, "_1")
        Sql = Sql & " WHERE " & cSelect
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
                If numasien <> "" Then LEtra = LEtra & vbCrLf & "Nº asiento: " & numasien
                
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
Dim Sql As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rsocios_pozos.codsocio, sum(acciones) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1 having sum(acciones) <> 0"
    
    TotalRegFacturasMto = TotalRegistrosConsulta(Sql)
    
End Function


Public Function TotalRegFacturasMtoUTXERA(cTabla As String, cWhere As String) As Long
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rpozos.codsocio, sum(hanegada) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1 having sum(hanegada) <> 0"
    
    TotalRegFacturasMtoUTXERA = TotalRegistrosConsulta(Sql)
    
End Function


Public Function TotalRegFacturasMantaESCALONA(cTabla As String, cWhere As String) As Long
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rcampos.codsocio, sum(round(supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) hanegada FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1 having sum(round(supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) <> 0"
    
    TotalRegFacturasMantaESCALONA = TotalRegistrosConsulta(Sql)
    
End Function



Public Function TotalRegFacturasTallaUTXERA(cTabla As String, cWhere As String) As Long
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    '[Monica]19/09/2012: ahora se factura al propietario del campo no al socio | 13/03/2014: ahora se factura al socio no al propietario
    Sql = "Select rcampos.codsocio codsocio, rcampos.codzonas, sum(round(supcoope * " & DBSet(vParamAplic.Faneca, "N") & ",2)) hanegada  FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1, 2  having sum(round(supcoope * " & DBSet(vParamAplic.Faneca, "N") & ",2)) <> 0"
    
    TotalRegFacturasTallaUTXERA = TotalRegistrosConsulta(Sql)
    
End Function




Private Function FacturacionMantenimiento(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rsocios_pozos.codsocio, sum(acciones) acciones "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
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
    
    B = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And B
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
        
        TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtCodigo(8).Text)), 2)
    
        IncrementarProgresNew Pb2, 1
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        Sql = Sql & "concepto, contabilizado) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
        Sql = Sql & ValorNulo & "," & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(txtCodigo(9).Text, "T") & ",0)"
        
        conn.Execute Sql
            
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        Rs.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionMantenimiento = False
    Else
        conn.CommitTrans
        FacturacionMantenimiento = True
    End If
End Function


Private Function FacturacionMantenimientoUTXERA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rpozos.codsocio, rpozos.hidrante, rpozos.hanegada  "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
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
    
    B = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And B
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
        TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtCodigo(8).Text)), 2)
    
        
        '[Monica]14/05/2012: tambien añadimos el poder poner una bonificacion o recargo (como en escalona)
        ' si hay bonificacion la calculamos
        If ComprobarCero(txtCodigo(53).Text) <> "0" Then
            PorcDto = CCur(ImporteSinFormato(txtCodigo(53).Text))
            Descuento = Round2(TotalFac * PorcDto / 100, 2)
            
            TotalFac = TotalFac + Descuento
        End If
    
    
        '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
        baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - baseimpo
    
        IncrementarProgresNew Pb2, 1
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        Sql = Sql & "concepto, contabilizado, porcdto, impdto, precio, escontado) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
        Sql = Sql & DBSet(Rs!Hidrante, "T") & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(txtCodigo(9).Text, "T") & ",0,"
        Sql = Sql & DBSet(PorcDto, "N") & ","
        Sql = Sql & DBSet(Descuento, "N") & ","
        Sql = Sql & DBSet(CCur(ImporteSinFormato(txtCodigo(8).Text)), "N") '& ")"
        
        '[Monica]02/09/2014: CONTADOSSSS
        If EsSocioContadoPOZOS(CStr(Rs!Codsocio)) Then
            Sql = Sql & ",1)"
        Else
            Sql = Sql & ",0)"
        End If
        
        conn.Execute Sql
            
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        cadMen = ""
        If B Then B = RepartoCoopropietarios(tipoMov, CStr(numfactu), CStr(FecFac), cadMen, False)
        cadMen = "Reparto Coopropietarios: " & cadMen
        
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        Rs.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionMantenimientoUTXERA = False
    Else
        conn.CommitTrans
        FacturacionMantenimientoUTXERA = True
    End If
End Function



Private Function FacturacionMantenimientoESCALONA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rpozos.codsocio, sum(rpozos.hanegada) hanegada, count(*) nrohidrante  "
'    Sql = "SELECT rpozos.codsocio, round(sum(rcampos.supcoope) * 12.03, 2) hanegada, count(*) nrohidrante  "
    Sql = Sql & " FROM  " & cTabla ' & ") INNER JOIN rcampos On rpozos.codcampo = rcampos.codcampo "

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
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
    
    B = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And B
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

        TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtCodigo(8).Text)), 2)
        
'        ' si lo que hacemos una factura de un importe no multimplicamos por nada
'        If Check1(6).Value Then TotalFac = Round2(DBLet(Rs!nrohidrante, "N") * CCur(ImporteSinFormato(txtCodigo(8).Text)), 2)
    
        ' si hay bonificacion la calculamos
        If ComprobarCero(txtCodigo(53).Text) <> "0" Then
            PorcDto = CCur(ImporteSinFormato(txtCodigo(53).Text)) * (-1)
            Descuento = Round2(TotalFac * PorcDto / 100, 2)
            
            TotalFac = TotalFac + Descuento
        Else
            If ComprobarCero(txtCodigo(61).Text) <> 0 Then
                PorcDto = CCur(ImporteSinFormato(txtCodigo(61).Text))
                Descuento = Round2(TotalFac * PorcDto / 100, 2)
                
                TotalFac = TotalFac + Descuento
            End If
        End If
    
        '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
        baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - baseimpo
    
        IncrementarProgresNew Pb2, 1
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        Sql = Sql & "concepto, contabilizado, porcdto, impdto, precio, escontado) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
        Sql = Sql & ValorNulo & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(txtCodigo(9).Text, "T") & ",0,"
        Sql = Sql & DBSet(PorcDto, "N") & ","
        Sql = Sql & DBSet(Descuento, "N") & ","
        Sql = Sql & DBSet(CCur(ImporteSinFormato(txtCodigo(8).Text)), "N") '& ")"
        
        '[Monica]02/09/2014: CONTADOSSSS
        If EsSocioContadoPOZOS(CStr(Rs!Codsocio)) Then
            Sql = Sql & ",1)"
        Else
            Sql = Sql & ",0)"
        End If
        
        conn.Execute Sql
            
            
        ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
'        Sql = "SELECT hidrante, round(rcampos.supcoope * 12.03, 2) hanegada "
        Sql = "SELECT hidrante, hanegada, nroorden "
        Sql = Sql & " FROM  " & cTabla '& ") INNER JOIN rcampos ON rpozos.codcampo = rcampos.codcampo"
'        Sql = Sql & " FROM  (" & cTabla & ") INNER JOIN rcampos ON rpozos.codcampo = rcampos.codcampo"
        If cWhere <> "" Then
            Sql = Sql & " WHERE " & cWhere
            Sql = Sql & " and rpozos.codsocio = " & DBSet(Rs!Codsocio, "N")
        Else
            Sql = Sql & " where rpozos.codsocio = " & DBSet(Rs!Codsocio, "N")
        End If
            
        Set Rs8 = New ADODB.Recordset
        Rs8.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = "insert into rrecibpozos_hid (codtipom, numfactu, fecfactu, hidrante, hanegada, nroorden) values  "
        CadValues = ""
        While Not Rs8.EOF
            CadValues = CadValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
            CadValues = CadValues & DBSet(Rs8!Hidrante, "T") & "," & DBSet(Rs8!hanegada, "N") & "," & DBSet(Rs8!nroorden, "N") & "),"
            Rs8.MoveNext
        Wend
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute Sql & CadValues
        End If
        Set Rs8 = Nothing
            
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
'[Monica]10/05/2012: no hay reparto de coopropietarios pq ese reparto va por hidrante, ya lo veremos
'        CadMen = ""
'        If b Then b = RepartoCoopropietarios(tipoMov, CStr(NumFactu), CStr(FecFac), CadMen, False)
'        CadMen = "Reparto Coopropietarios: " & CadMen
'
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        Rs.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionMantenimientoESCALONA = False
    Else
        conn.CommitTrans
        FacturacionMantenimientoESCALONA = True
    End If
End Function


Private Function FacturacionTallaESCALONA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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
    B = True
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    '[Monica]13/03/2014: ahora se factura al socio no al propietario
    Sql = "SELECT rcampos.codsocio codsocio, rcampos.codzonas, round(sum(rcampos.supcoope) / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    
    Sql = Sql & " group by 1, 2 having hanegada <> 0  "
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by codsocio, codzonas "
    
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
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        SocioAnt = DBLet(Rs!Codsocio, "N")
        
    End If
    
    While Not Rs.EOF And B
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
            Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
            Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
            Sql = Sql & "concepto, contabilizado, porcdto, impdto, precio,escontado) "
            Sql = Sql & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(SocioAnt, "N") & ","
            Sql = Sql & ValorNulo & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(txtCodigo(76).Text, "T") & ",0,"
            Sql = Sql & DBSet(PorcDto, "N") & ","
            Sql = Sql & DBSet(Descuento, "N") & ","
            Sql = Sql & DBSet(PrecioBrz, "N") '& ")"
            
            '[Monica]02/09/2014: CONTADOSSSS
            If EsSocioContadoPOZOS(CStr(SocioAnt)) Then
                Sql = Sql & ",1)"
            Else
                Sql = Sql & ",0)"
            End If
            
            conn.Execute Sql
            
            ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
            Sql = "SELECT rcampos.codcampo, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
            Sql = Sql & " FROM  " & cTabla
            If cWhere <> "" Then
                Sql = Sql & " WHERE " & cWhere
                '[Monica]13/03/2014: se factura al socio no al propietario
                Sql = Sql & " and rcampos.codsocio = " & DBSet(SocioAnt, "N")
            Else
                '[Monica]13/03/2014: se factura al socio no al propietario
                Sql = Sql & " where rcampos.codsocio = " & DBSet(SocioAnt, "N")
            End If
                
            Set Rs8 = New ADODB.Recordset
            Rs8.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Sql = "insert into rrecibpozos_cam (codtipom, numfactu, fecfactu, codcampo, hanegada, precio1, precio2, codzonas, poligono, parcela, subparce) values  "
            CadValues = ""
            While Not Rs8.EOF
                Precio = DevuelvePrecio(DBLet(Rs8!codzonas, "N"))
                
                CadValues = CadValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                CadValues = CadValues & DBSet(Rs8!codcampo, "N") & "," & DBSet(Rs8!hanegada, "N") & ","
                CadValues = CadValues & DBSet(PrecioTalla1, "N") & "," & DBSet(PrecioTalla2, "N") & "," & DBSet(ZonaTalla, "N") & ","
                CadValues = CadValues & DBSet(Rs8!Poligono, "N") & "," & DBSet(Rs8!Parcela, "N") & "," & DBSet(Rs8!SubParce, "T") & "),"
                
                Rs8.MoveNext
            Wend
            
            If CadValues <> "" Then
                CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
                conn.Execute Sql & CadValues
            End If
            Set Rs8 = Nothing
                
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            
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
    
    If HayReg And B Then
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
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        Sql = Sql & "concepto, contabilizado, porcdto, impdto, precio, escontado) "
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(SocioAnt, "N") & ","
        Sql = Sql & ValorNulo & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(txtCodigo(76).Text, "T") & ",0,"
        Sql = Sql & DBSet(PorcDto, "N") & ","
        Sql = Sql & DBSet(Descuento, "N") & ","
        Sql = Sql & DBSet(PrecioBrz, "N") '& ")"
        
        '[Monica]02/09/2014: CONTADOSSSS
        If EsSocioContadoPOZOS(CStr(SocioAnt)) Then
            Sql = Sql & ",1)"
        Else
            Sql = Sql & ",0)"
        End If
            
        conn.Execute Sql
        
        ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
        Sql = "SELECT rcampos.codcampo, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
        Sql = Sql & " FROM  " & cTabla
        If cWhere <> "" Then
            Sql = Sql & " WHERE " & cWhere
            '[Monica]13/03/2014: se factura al socio no al propietario
            Sql = Sql & " and rcampos.codsocio = " & DBSet(SocioAnt, "N")
        Else
            '[Monica]13/03/2014: se factura al socio no al propietario
            Sql = Sql & " where rcampos.codsocio = " & DBSet(SocioAnt, "N")
        End If
            
        Set Rs8 = New ADODB.Recordset
        Rs8.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = "insert into rrecibpozos_cam (codtipom, numfactu, fecfactu, codcampo, hanegada, precio1, precio2, codzonas, poligono, parcela, subparce) values  "
        CadValues = ""
        While Not Rs8.EOF
            Precio = DevuelvePrecio(DBLet(Rs8!codzonas))
        
            CadValues = CadValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
            CadValues = CadValues & DBSet(Rs8!codcampo, "N") & "," & DBSet(Rs8!hanegada, "N") & ","
            CadValues = CadValues & DBSet(PrecioTalla1, "N") & "," & DBSet(PrecioTalla2, "N") & "," & DBSet(ZonaTalla, "N") & ","
            CadValues = CadValues & DBSet(Rs8!Poligono, "N") & "," & DBSet(Rs8!Parcela, "N") & "," & DBSet(Rs8!SubParce, "T") & "),"
            
            Rs8.MoveNext
        Wend
        
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute Sql & CadValues
        End If
        Set Rs8 = Nothing
            
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
    
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionTallaESCALONA = False
    Else
        conn.CommitTrans
        FacturacionTallaESCALONA = True
    End If
End Function


Private Function ActualizacionTallaESCALONA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rrecibpozos.* "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    
    ' ordenado por socio
    Sql = Sql & " order by rrecibpozos.codsocio, rrecibpozos.numfactu "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    B = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And B
        HayReg = True
        
        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        TotalFac = DBLet(Rs!TotalFact, "N")
        
        ' si hay bonificacion la calculamos
        If CCur(ComprobarCero(txtCodigo(78).Text)) <> 0 Then
            PorcDto = CCur(ImporteSinFormato(txtCodigo(78).Text)) * (-1)
            Descuento = Round2(TotalFac * PorcDto / 100, 2)
            
            TotalFac = TotalFac + Descuento
        Else
            If CCur(ComprobarCero(txtCodigo(77).Text)) <> 0 Then
                PorcDto = CCur(ImporteSinFormato(txtCodigo(77).Text))
                Descuento = Round2(TotalFac * PorcDto / 100, 2)
                
                TotalFac = TotalFac + Descuento
            End If
        End If
    
        '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
        baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - baseimpo
    
        IncrementarProgresNew Pb1, 1
        
        'modificamos la tabla de recibos de pozos
        Sql = "update rrecibpozos set baseimpo = " & DBSet(baseimpo, "N")
        Sql = Sql & ", tipoiva = " & DBSet(vParamAplic.CodIvaPOZ, "N")
        Sql = Sql & ", porc_iva = " & DBSet(PorcIva, "N")
        Sql = Sql & ", imporiva = " & DBSet(ImpoIva, "N")
        Sql = Sql & ", totalfact = " & DBSet(TotalFac, "N")
        Sql = Sql & ", porcdto = " & DBSet(PorcDto, "N")
        Sql = Sql & ", impdto = " & DBSet(Descuento, "N")
        Sql = Sql & " where codtipom = 'TAL'"
        Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
        Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        Sql = Sql & " and codsocio = " & DBSet(Rs!Codsocio, "N")
        
        conn.Execute Sql
            
        ' Si el recibo está contabilizado actualizaremos el arimoney
        LetraSerie = DevuelveValor("select letraser from usuarios.stipom where codtipom = 'TAL'")

        Sql = "update scobro set impvenci = " & DBSet(TotalFac, "N")
        Sql = Sql & " where numserie = " & DBSet(LetraSerie, "T")
        Sql = Sql & " and codfaccl = " & DBSet(Rs!numfactu, "N")
        Sql = Sql & " and fecfaccl = " & DBSet(Rs!fecfactu, "F")
        Sql = Sql & " and numorden = 1 "
            
        ConnConta.Execute Sql
        
        '[Monica]19/09/2012: al enlazar por el propietario y campos me salian todos los campos de ese propietario,
        '                    si el nro de factura, tipo ya existe no lo volvemos a insertar en el resumen
        '                    He añadido: and totalregistros...
        If B And TotalRegistros("select count(*) from tmpinformes where codusu = " & vUsu.Codigo & " and nombre1 = 'TAL' and importe1 = " & DBSet(Rs!numfactu, "N")) = 0 Then B = InsertResumen("TAL", CStr(Rs!numfactu))
        
        Rs.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
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
    Sql = Sql & " and nombre1  "
    Select Case Tipo
        Case 0 ' recibos de consumo de pozos
            Sql = Sql & " in ('RCP','FIN')"
        Case 1 ' recibos de mantenimiento de pozos
            Sql = Sql & "='RMP'"
        Case 2 ' recibos de contadores de pozos
            Sql = Sql & "='RVP'"
        Case 3
            Sql = Sql & "='TAL'"
        Case 4
            Sql = Sql & "='RMT'"
        Case 5
            Sql = Sql & "='ALV'"
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
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long
Dim FecFac As Date
Dim Mens As String

Dim B As Boolean
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
        
        Mens = "Proceso Facturación Contadores: " & vbCrLf & vbCrLf
        B = FacturacionContadores(nTabla, cadSelect, txtCodigo(22).Text, Me.Pb3, Mens)
        If B Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
                           
            'IMPRESION DEL RESUMEN DE LA FACTURACION de recibos de contadores
            If Me.Check1(5).Value Then
                cadFormula = ""
                CadParam = CadParam & "pFecFac= """ & txtCodigo(22).Text & """|"
                numParam = numParam + 1
                CadParam = CadParam & "pTitulo= ""Resumen Facturación Contadores""|"
                numParam = numParam + 1
                
                FecFac = CDate(txtCodigo(22).Text)
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
                FecFac = CDate(txtCodigo(22).Text)
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

Private Function FacturacionContadores(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rsocios.codsocio "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
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
    
    B = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And B
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
        
        TotalFac = CCur(ImporteSinFormato(ComprobarCero(txtCodigo(33).Text)))
        '[Monica]08/06/2016: antes en estas facturas no grababamos el importe de iva y en baseimpo poniamos el totalfac
        baseimpo = Round2(TotalFac / (1 + (vPorcIva / 100)), 2)
        ImpoIva = TotalFac - baseimpo
        
    
        IncrementarProgresNew Pb3, 1
        
        'insertar en la tabla de recibos de pozos
        Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
        Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
        Sql = Sql & "concepto, contabilizado, conceptomo, importemo, conceptoar1, importear1, conceptoar2, importear2, conceptoar3, "
        Sql = Sql & "importear3, conceptoar4, importear4"
        '[Monica]02/09/2014: CONTADOSSSS
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            Sql = Sql & ",escontado) "
        Else
            Sql = Sql & ") "
        End If
        Sql = Sql & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
        Sql = Sql & ValorNulo & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & ValorNulo & ",0,"
        Sql = Sql & DBSet(txtCodigo(20).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtCodigo(21).Text))), "N", "S") & "," ' mano de obra
        Sql = Sql & DBSet(txtCodigo(25).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtCodigo(26).Text))), "N", "S") & "," ' articulo 1
        Sql = Sql & DBSet(txtCodigo(27).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtCodigo(28).Text))), "N", "S") & "," ' articulo 2
        Sql = Sql & DBSet(txtCodigo(29).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtCodigo(30).Text))), "N", "S") & "," ' articulo 3
        Sql = Sql & DBSet(txtCodigo(31).Text, "T") & "," & DBSet(CCur(ImporteSinFormato(ComprobarCero(txtCodigo(32).Text))), "N", "S") '& ")" ' articulo 4
        
        '[Monica]02/09/2014: CONTADOSSSS
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            If EsSocioContadoPOZOS(CStr(Rs!Codsocio)) Then
                Sql = Sql & ",1)"
            Else
                Sql = Sql & ",0)"
            End If
        Else
            Sql = Sql & ")"
        End If
        
        
        conn.Execute Sql
            
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        Rs.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionContadores = False
    Else
        conn.CommitTrans
        FacturacionContadores = True
    End If
End Function


Private Sub CalcularTotales()
Dim Total As Currency

    Total = 0
    
    If txtCodigo(21).Text <> "" Then Total = Total + CCur(ImporteSinFormato(txtCodigo(21).Text))
    If txtCodigo(26).Text <> "" Then Total = Total + CCur(ImporteSinFormato(txtCodigo(26).Text))
    If txtCodigo(28).Text <> "" Then Total = Total + CCur(ImporteSinFormato(txtCodigo(28).Text))
    If txtCodigo(30).Text <> "" Then Total = Total + CCur(ImporteSinFormato(txtCodigo(30).Text))
    If txtCodigo(32).Text <> "" Then Total = Total + CCur(ImporteSinFormato(txtCodigo(32).Text))

    txtCodigo(33).Text = Total
    PonerFormatoDecimal txtCodigo(33), 3

End Sub


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
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
        Combo1(0).AddItem "RMT-Consumo Manta"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 4
        '[Monica]14/01/2016: rectificativas
        Combo1(0).AddItem "RRC-Rect.Consumo"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 5
        Combo1(0).AddItem "RRM-Rect.Mantenimiento"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 6
        Combo1(0).AddItem "RRV-Rect.Contadores"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 7
        Combo1(0).AddItem "RTA-Rect.Talla"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 8
'        Combo1(0).AddItem "RRT-Rect.Consumo Manta"
'        Combo1(0).ItemData(Combo1(0).NewIndex) = 9
    End If
    If vParamAplic.Cooperativa = 7 Then
        Combo1(0).AddItem "FIN-Interna"
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
        Combo1(1).AddItem "RMT-Consumo Manta"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 5
        '[Monica]14/01/2016: rectificativas
        Combo1(1).AddItem "RRC-Rect.Consumo"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 6
        Combo1(1).AddItem "RRM-Rect.Mantenimiento"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 7
        Combo1(1).AddItem "RRV-Rect.Contadores"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 8
        Combo1(1).AddItem "RTA-Rect.Talla"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 9
'        Combo1(1).AddItem "RRT-Rect.Consumo Manta"
'        Combo1(1).ItemData(Combo1(1).NewIndex) = 10
    End If
    If vParamAplic.Cooperativa = 7 Then
        Combo1(1).AddItem "FIN-Interna"
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
        '[Monica]14/01/2016: rectificativas
        Combo1(2).AddItem "RRC-Rect.Consumo"
        Combo1(2).ItemData(Combo1(2).NewIndex) = 5
        Combo1(2).AddItem "RRM-Rect.Mantenimiento"
        Combo1(2).ItemData(Combo1(2).NewIndex) = 6
        Combo1(2).AddItem "RRV-Rect.Contadores"
        Combo1(2).ItemData(Combo1(2).NewIndex) = 7
        Combo1(2).AddItem "RTA-Rect.Talla"
        Combo1(2).ItemData(Combo1(2).NewIndex) = 8
'        Combo1(2).AddItem "RRT-Rect.Consumo Manta"
'        Combo1(2).ItemData(Combo1(2).NewIndex) = 9
    End If
    If vParamAplic.Cooperativa = 7 Then
        Combo1(2).AddItem "FIN-Interna"
        Combo1(2).ItemData(Combo1(2).NewIndex) = 3
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

Dim B As Boolean
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
                Mens = "Proceso Facturación Consumo: " & vbCrLf & vbCrLf
                B = FacturacionConsumoUTXERA(nTabla, cadSelect, txtCodigo(14).Text, Me.Pb1, Mens)
                If B Then
                    If Not HayFacturas Then
                        MsgBox "No se han generado de facturas de consumo.", vbExclamation
                        cmdCancel_Click (0)
                        Exit Sub
                    End If
                                
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                   
                    'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                    If Me.Check1(2).Value Then
                        cadFormula = ""
                        CadParam = CadParam & "pFecFac= """ & txtCodigo(14).Text & """|"
                        numParam = numParam + 1
                        CadParam = CadParam & "pTitulo= ""Resumen Facturación de Contadores""|"
                        numParam = numParam + 1
                        
                        FecFac = CDate(txtCodigo(14).Text)
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
'                        cadAux = "({rrecibpozos.numfactu} IN [" & FacturasGeneradasPOZOS(0) & "])"
'                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
'                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
'                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        '[Monica]06/03/2013: solo lo facturado
                        cadAux = "{rrecibpozos.codtipom} = 'RCP'"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        'Fecha de Factura
                        FecFac = CDate(txtCodigo(14).Text)
                        cadAux = "{rrecibpozos.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = "{rrecibpozos.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        indRPT = 46 'Impresion de recibos de consumo de contadores de pozos
                        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
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


Private Function FacturacionConsumoUTXERA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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


    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT rpozos.codsocio,hidrante,nroorden,codparti,poligono,parcelas,hanegada,lect_ant,lect_act,fech_ant,fech_act,codpozo,consumo "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If

    ' ordenado por socio, hidrante
    Sql = Sql & " order by rpozos.codsocio, rpozos.hidrante "

    Set vSeccion = New CSeccion

    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If

    B = True

    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    If vPorcIva = "" Then vPorcIva = "0"
    PorcIva = CCur(ImporteSinFormato(vPorcIva))

    Set vTipoMov = New CTiposMov

    HayReg = False

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    baseimpo = 0
    ImpoIva = 0
    TotalFac = 0


    While Not Rs.EOF And B
        HayReg = True
            
            
        IncrementarProgresNew Pb1, 1

'If RS!CodSocio = 168 Then
'    MsgBox "168"
'End If



        '[Monica]17/05/2013: añadida la condicion de que el consumo ha de ser superior o igual al mínimo
        '[Monica]24/10/2011: añadida esta condicion para que si no hay consumo se actualicen fechas
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
            
            baseimpo = Round2(ConsTra1 * CCur(ImporteSinFormato(txtCodigo(4).Text)), 2) + _
                       Round2(ConsTra1 * CCur(ImporteSinFormato(txtCodigo(5).Text)), 2)
    
    
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
            TotalFac = baseimpo + Round2(baseimpo * PorcIva / 100, 2)
    
    
            'insertar en la tabla de recibos de pozos
            Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, numlinea, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
            Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, codparti, parcelas, poligono, nroorden, escontado) "
            Sql = Sql & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ",1,"
            Sql = Sql & DBSet(Rs!Hidrante, "T") & "," & DBSet(baseimpo, "N") & "," & vParamAplic.CodIvaPOZ & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(ConsumoHidrante, "N") & "," & DBSet(0, "N") & ","
            Sql = Sql & DBSet(Rs!lect_ant, "N") & "," & DBSet(Rs!fech_ant, "F") & ","
            Sql = Sql & DBSet(Rs!lect_act, "N") & "," & DBSet(Rs!fech_act, "F") & ","
            Sql = Sql & DBSet(ConsTra1, "N") & "," & DBSet(ImporteSinFormato(txtCodigo(4).Text), "N") & ","
            Sql = Sql & DBSet(ConsTra1, "N") & "," & DBSet(ImporteSinFormato(txtCodigo(5).Text), "N") & ","
            
            '[Monica]22/10/2012: si nos han puesto un concepto guardammos el concepto
            ' antes :     Sql = Sql & "'Recibo de Consumo',0,"
            If txtCodigo(48).Text <> "" Then
                Sql = Sql & DBSet(txtCodigo(48).Text, "T") & ",0,"
            Else
                Sql = Sql & DBSet(vTipoMov.NombreMovimiento, "T") & ",0,"
            End If
            
            '[Monica]22/10/2012: guardamos tambien la partida [Monica]03/05/2013: ahora tb el poligono [Monica]22/07/2013: metemos el nro de orden
            Sql = Sql & DBSet(Rs!codparti, "N") & "," & DBSet(Rs!parcelas, "T") & "," & DBSet(Rs!Poligono, "T") & "," & DBSet(Rs!nroorden, "N") '& ")"
    
            '[Monica]02/09/2014: CONTADOSSSS
            If EsSocioContadoPOZOS(CStr(Rs!Codsocio)) Then
                Sql = Sql & ",1)"
            Else
                Sql = Sql & ",0)"
            End If
    
            conn.Execute Sql
        
            If B Then B = RepartoCoopropietarios(tipoMov, CStr(numfactu), CStr(FecFac), cadMen)
            cadMen = "Reparto Coopropietarios: " & cadMen
        
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        End If

        If DBLet(Rs!fech_act, "F") <> "" Then
            '[Monica]11/06/2013: añadida la condicion de que el consumo sea inferior o igual al consumo maximo de parametros
            If DBLet(ConsumoHidrante, "N") >= vParamAplic.ConsumoMinPOZ And DBLet(ConsumoHidrante, "N") <= vParamAplic.ConsumoMaxPOZ Then
            
                HayFacturas = True
            
                ' actualizar en los acumulados de hidrantes
                Sql = "update rpozos set acumconsumo = acumconsumo + " & DBSet(ConsumoHidrante, "N")
                Sql = Sql & ", lect_ant = lect_act "
                Sql = Sql & ", fech_ant = fech_act "
        '        sql = sql & ", lect_act = null "
                Sql = Sql & ", fech_act = null "
                Sql = Sql & ", consumo = 0 "
                Sql = Sql & " WHERE hidrante = " & DBSet(Rs!Hidrante, "T")
            Else
                '[Monica]17/05/2013: en el caso de que el consumo no supere el mínimo
                '                    dejamos la lectura actual = a la que tenia la lectura anterior
                '                    la fecha anterior no se actualiza
                '                    y la fecha actual se deja a null
                Sql = "update rpozos set lect_act = lect_ant "
                Sql = Sql & ", fech_act = null "
                Sql = Sql & ", consumo = 0 "
                Sql = Sql & " WHERE hidrante = " & DBSet(Rs!Hidrante, "T")
            End If
            
            conn.Execute Sql
        End If
        
        Rs.MoveNext
    Wend

    vSeccion.CerrarConta
    Set vSeccion = Nothing

eFacturacion:
    If Err.Number <> 0 Or Not B Then
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
Dim B As Boolean
Dim devuelve As String
Dim Existe As Boolean
Dim Mens As String
Dim NroPropiedad As String
Dim NumReg As Long
Dim campo As Long
Dim Porcentaje As Single
Dim numFac As Long
Dim vPorcen As String
Dim vPorcIva As Currency

    On Error GoTo eRepartoCoopropietarios

    RepartoCoopropietarios = False
    
    cadErr = ""
    
    B = True
    
    Sql = "select * from rrecibpozos where codtipom  = " & DBSet(tipoMov, "T")
    Sql = Sql & " and numfactu = " & DBSet(Factura, "N")
    Sql = Sql & " and fecfactu = " & DBSet(Fecha, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        
        If TieneCopropietariosPOZOS(CStr(Rs!Hidrante), CStr(Rs!Codsocio)) Then
            CodTipoMov = tipoMov
        
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(CodTipoMov) Then

    
                '[Monica]08/06/2016: sacamos el porcentaje de iva grabado
                vPorcIva = DBLet(Rs!porc_iva, "N")


                tBaseImpo = DBLet(Rs!baseimpo, "N")
                tImporIva = DBLet(Rs!ImporIva, "N")
                tTotalFact = DBLet(Rs!TotalFact, "N")

                Sql2 = "select * from rpozos_cooprop where hidrante = " & DBSet(Rs!Hidrante, "T")
                Sql2 = Sql2 & " and rpozos_cooprop.codsocio <> " & DBSet(Rs!Codsocio, "N")
                Sql2 = Sql2 & " order by numlinea "
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs2.EOF And B
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
                    
'                    vBaseImpo = Round2(DBLet(Rs!baseimpo, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
'                    vImporIva = Round2(DBLet(Rs!ImporIva, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vTotalFact = Round2(DBLet(Rs!TotalFact, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    '[Monica]08/06/2016: cambio de calculo de base imponible
                    vBaseImpo = Round2(vTotalFact / (1 + (vPorcIva / 100)), 2)
                    vImporIva = vTotalFact - vBaseImpo
                    
                    
                    tBaseImpo = tBaseImpo - vBaseImpo
                    tImporIva = tImporIva - vImporIva
                    tTotalFact = tTotalFact - vTotalFact
                    
                    'insertar en la tabla de recibos de pozos
                    Sql4 = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, numlinea, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
                    Sql4 = Sql4 & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, concepto, contabilizado, "
                    '[Monica]29/04/2014: añadidos campos que faltaban
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
                    
                    '[Monica]29/04/2014: añadidos los campos que faltaban
                    Sql4 = Sql4 & DBSet(Rs!conceptomo, "T") & "," & DBSet(Rs!importemo, "N") & "," & DBSet(Rs!Conceptoar1, "T") & "," & DBSet(Rs!importear1, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!Conceptoar2, "T") & "," & DBSet(Rs!importear2, "N") & "," & DBSet(Rs!conceptoar3, "T") & "," & DBSet(Rs!importear3, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!conceptoar4, "T") & "," & DBSet(Rs!importear4, "N") & "," & DBSet(Rs!difdias, "N") & "," & DBSet(Rs!Calibre, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!codpozo, "N") & "," & DBSet(Rs!PorcDto, "N") & "," & DBSet(Rs!ImpDto, "N") & "," & DBSet(Rs!Precio, "N") & "," & DBSet(Rs!pasaridoc, "N") & ","
                    
                    '[Monica]22/10/2012: guardamos tambien la partida [Monica]03/05/2013: ahora tb el poligono [Monica]22/07/2013: ahora tb metemos el nro de orden
                    Sql4 = Sql4 & DBSet(Rs!codparti, "N") & "," & DBSet(Rs!parcelas, "T") & "," & DBSet(Rs!Poligono, "T") & "," & DBSet(Rs!nroorden, "N") '& ")"

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
                    
                    If B Then B = InsertResumen(tipoMov, CStr(numFac))
                    
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                If B Then
                    '[Monica]08/06/2016: recalculo de base imponivle y de importe de iva
                    vBaseImpo = Round2(tTotalFact / (1 + (vPorcIva / 100)), 2)
                    vImporIva = tTotalFact - vBaseImpo
                
                
                    vPorcen = DevuelveValor("select porcentaje from rpozos_cooprop where codsocio = " & DBSet(Rs!Codsocio, "N") & " and hidrante = " & DBSet(Rs!Hidrante, "T"))
                    vPorcen = Format(vPorcen, "##0.00") & "%"
                    vPorcen = " " & vPorcen
                
                
                
                    ' ultimo registro la diferencia ( se updatean las tablas del registro de rrecibpozos origen )
                    Sql4 = "update rrecibpozos set baseimpo = " & DBSet(vBaseImpo, "N") & "," ' antes tBaseImpo
                    
                    If SinIva Then
                    
                    Else
                        Sql4 = Sql4 & "imporiva = " & DBSet(vImporIva, "N") & "," ' antes tImporIva
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
                B = False
            End If
        
        End If
    
    End If
    
    Set Rs = Nothing

eRepartoCoopropietarios:
    If Err.Number <> 0 Or Not B Then
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
Dim Sql As String
Dim Inicio As Long
Dim Fin As Long
Dim NroDig As Integer
Dim Limite As Long


    On Error GoTo eCalculoConsumoHidrante


    CalculoConsumoHidrante = False
    
    Sql = "select * from rpozos where hidrante = " & DBSet(Hidrante, "T")
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
       Inicio = 0
       Fin = 0
       NroDig = DBLet(Rs!Digcontrol, "N")
       Limite = 10 ^ NroDig
       
       Inicio = DBLet(Rs!lect_ant, "N")
       Fin = CLng(txtCodigo(51).Text)
    
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

Private Function DevuelvePrecio(Zona As Integer) As Currency
Dim Sql As String
Dim Precio As Currency
Dim Rs As ADODB.Recordset
Dim Prec1Zona0 As Currency
Dim Prec2Zona0 As Currency
    
    PrecioTalla1 = 0
    PrecioTalla2 = 0
    ZonaTalla = 0
    Prec1Zona0 = 0
    Prec2Zona0 = 0
    
    Sql = "select precio1, precio2 from rzonas where codzonas = 0"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Prec1Zona0 = DBLet(Rs.Fields(0).Value, "N")
        Prec2Zona0 = DBLet(Rs.Fields(1).Value, "N")
    End If
    
    Set Rs = Nothing
    
    Sql = "select precio1, precio2 from rzonas where codzonas = " & DBSet(Zona, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
Dim Sql As String
Dim SqlValues As String

    On Error GoTo eCargarTablaPrecios

    CargarTablaPrecios = False

    Sql = "delete from rpretallapoz "
    conn.Execute Sql
    
    SqlValues = ""
    
    Sql = "insert ignore into rpretallapoz (codzonas, precio1, precio2) values "
    
    SqlValues = SqlValues & "(0," & DBSet(txtCodigo(72).Text, "N") & "," & DBSet(txtCodigo(66).Text, "N") & "),"
    
    If txtCodigo(79).Text <> "" Then
        SqlValues = SqlValues & "(" & DBSet(txtCodigo(79).Text, "N") & "," & DBSet(txtCodigo(80).Text, "N") & "," & DBSet(txtCodigo(81).Text, "N") & "),"
    End If
    
    If txtCodigo(82).Text <> "" Then
        SqlValues = SqlValues & "(" & DBSet(txtCodigo(82).Text, "N") & "," & DBSet(txtCodigo(83).Text, "N") & "," & DBSet(txtCodigo(84).Text, "N") & "),"
    End If
    
    If txtCodigo(85).Text <> "" Then
        SqlValues = SqlValues & "(" & DBSet(txtCodigo(85).Text, "N") & "," & DBSet(txtCodigo(86).Text, "N") & "," & DBSet(txtCodigo(87).Text, "N") & "),"
    End If

    If SqlValues <> "" Then
        conn.Execute Sql & Mid(SqlValues, 1, Len(SqlValues) - 1)
    End If
    
    CargarTablaPrecios = True
    Exit Function

eCargarTablaPrecios:
    MuestraError Err.Number, "Cargar Tabla Precios", Err.Description

End Function

Private Sub EnviarEMailMulti(cadWHERE As String, cadTit As String, cadRpt As String, cadTabla As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer

On Error GoTo EEnviar

    Screen.MousePointer = vbHourglass
    
    Sql = "SELECT distinct rsocios.codsocio,nomsocio,maisocio "
    Sql = Sql & "FROM " & cadTabla
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ' Primero la borro por si acaso
    Sql = " DROP TABLE IF EXISTS tmpMail;"
    conn.Execute Sql
    
    'creamos una temporal donde guardamos para cada proveedor que SI tiene
    'e-mail, el mail1 o el mail2 al que vamos a enviar
    Sql = "CREATE TEMPORARY TABLE tmpMail ( "
    Sql = Sql & "codusu INT(7) UNSIGNED  DEFAULT '0' NOT NULL, "
    Sql = Sql & "codprove INT(6) UNSIGNED  DEFAULT '0' NOT NULL, "
    Sql = Sql & "nomprove varchar(40)  DEFAULT '' NOT NULL, "
    Sql = Sql & "email varchar(40)  DEFAULT '' NOT NULL) "
    conn.Execute Sql
    
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
                
                Sql = "{rsocios.codsocio}=" & Rs.Fields(0)

                .Opcion = 86
                .FormulaSeleccion = Sql
                .EnvioEMail = True
                CadenaDesdeOtroForm = "GENERANDO"
                .Titulo = "Cartas Talla"
                .NombreRPT = cadRpt
                .ConSubInforme = True
                .Show vbModal

                If CadenaDesdeOtroForm = "" Then
                'si se ha generado el .pdf para enviar
                    Sql = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                    Sql = Sql & " VALUES (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "T") & "," & DBSet(Cad1, "T") & ")"
                    conn.Execute Sql
            
                    Me.Refresh
                    espera 0.4
                    cont = cont + 1
                    'Se ha generado bien el documento
                    'Lo copiamos sobre app.path & \temp
                    Sql = Rs.Fields(0) & ".pdf"
                    FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Sql
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
        Sql = "Carta de Talla" & "|"
       
       
        frmEMail.Opcion = 2
        frmEMail.DatosEnvio = Sql
        frmEMail.CodCryst = IndRptReport
        frmEMail.Ficheros = ""
        frmEMail.EsCartaTalla = True
        frmEMail.Show vbModal

        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute Sql
        
        'Borrar la carpeta con temporales
        Kill App.Path & "\temp\*.pdf"
    End If
    
    Screen.MousePointer = vbDefault
   
EEnviar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Enviando Carta de Talla por e-mail", Err.Description
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute Sql
    End If
End Sub

Private Function FacturacionTallaPreviaESCALONA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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
    
    B = True
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    Sql = "delete from tmpinformes2 where codusu = " & vUsu.Codigo
    conn.Execute Sql

    '[Monica]13/03/2014: se factura al socio no al propietario antes era codpropiet
    Sql = "SELECT rcampos.codsocio codsocio, rcampos.codzonas, round(sum(rcampos.supcoope) / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    
    Sql = Sql & " group by 1, 2 having hanegada <> 0  "
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by codsocio, codzonas "
    
    Me.pb5.visible = True
    Nregs = TotalRegistrosConsulta(Sql)
    CargarProgresNew pb5, Nregs
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
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        baseimpo = 0
        ImpoIva = 0
        TotalFac = 0
        
        SocioAnt = DBLet(Rs!Codsocio, "N")
        numfactu = 0
    End If
    
    While Not Rs.EOF And B
        HayReg = True
        
        If SocioAnt <> DBLet(Rs!Codsocio, "N") Then
        
            numfactu = numfactu + 1
        
            baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
            ImpoIva = TotalFac - baseimpo
        
            'insertar en la tabla de recibos de pozos tmpinformes
            '                               codusu, numfactu,fecfactu,codsocio,baseimpo,codivapoz,porciva,imporiva, totalfac, concepto
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, codigo1, importe2, campo1, porcen1, importe3, importe4, nombre1) "
            Sql = Sql & " values (" & vUsu.Codigo & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(SocioAnt, "N") & ","
            Sql = Sql & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql = Sql & DBSet(TotalFac, "N") & ","
            Sql = Sql & DBSet(txtCodigo(97).Text, "T") & ")"
            
            conn.Execute Sql
            
            ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
            Sql = "SELECT rcampos.codcampo, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
            Sql = Sql & " FROM  " & cTabla
            If cWhere <> "" Then
                Sql = Sql & " WHERE " & cWhere
                '[Monica]13/03/2014: hidrantes del socio, antes eran hidrantes del propietario
                Sql = Sql & " and rcampos.codsocio = " & DBSet(SocioAnt, "N")
            Else
                '[Monica]13/03/2014: hidrantes del socio, antes eran hidrantes del propietario
                Sql = Sql & " where rcampos.codsocio = " & DBSet(SocioAnt, "N")
            End If
                
            Set Rs8 = New ADODB.Recordset
            Rs8.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            '                                       numfactu, fecfactu,codcampo,hanegadas,precio1, precio2, codzona
            Sql = "insert into tmpinformes2 (codusu, importe1, fecha1, importe2, importe3, precio1, precio2, campo1) values  "
            CadValues = ""
            While Not Rs8.EOF
                Precio = DevuelvePrecio(DBLet(Rs8!codzonas, "N"))
                
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                CadValues = CadValues & DBSet(Rs8!codcampo, "N") & "," & DBSet(Rs8!hanegada, "N") & ","
                CadValues = CadValues & DBSet(PrecioTalla1, "N") & "," & DBSet(PrecioTalla2, "N") & "," & DBSet(ZonaTalla, "N") & "),"
                
                Rs8.MoveNext
            Wend
            
            If CadValues <> "" Then
                CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
                conn.Execute Sql & CadValues
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
        
        IncrementarProgresNew pb5, 1
        
'        Label2(78).Caption = "Socio: " & Format(Rs!Codsocio, "000000")
        DoEvents
        
        Rs.MoveNext
    Wend
    
    If HayReg And B Then
        numfactu = numfactu + 1
            
        baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
        ImpoIva = TotalFac - baseimpo
    
        'insertar en la tabla de recibos de pozos (intermedia)
        Sql = "insert into tmpinformes (codusu, importe1, fecha1, codigo1, importe2, campo1, porcen1, importe3, importe4, nombre1) "
        Sql = Sql & " values (" & vUsu.Codigo & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(SocioAnt, "N") & ","
        Sql = Sql & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql = Sql & DBSet(TotalFac, "N") & ","
        Sql = Sql & DBSet(txtCodigo(97).Text, "T") & ")"
        
        conn.Execute Sql
        
        ' Introducimos en la tabla de lineas que hidrantes intervienen en la factura para la impresion
        Sql = "SELECT rcampos.codcampo, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
        Sql = Sql & " FROM  " & cTabla
        If cWhere <> "" Then
            Sql = Sql & " WHERE " & cWhere
            '[Monica]13/03/2014: hidrantes del socio, antes eran hidrantes del propietario
            Sql = Sql & " and rcampos.codsocio = " & DBSet(SocioAnt, "N")
        Else
            Sql = Sql & " where rcampos.codsocio = " & DBSet(SocioAnt, "N")
        End If
            
        Set Rs8 = New ADODB.Recordset
        Rs8.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        '                                       numfactu, fecfactu,codcampo,hanegadas,precio1, precio2, codzona
        Sql = "insert into tmpinformes2 (codusu, importe1, fecha1, importe2, importe3, precio1, precio2, campo1) values  "
        CadValues = ""
        While Not Rs8.EOF
            Precio = DevuelvePrecio(DBLet(Rs8!codzonas, "N"))
            
            CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
            CadValues = CadValues & DBSet(Rs8!codcampo, "N") & "," & DBSet(Rs8!hanegada, "N") & ","
            CadValues = CadValues & DBSet(PrecioTalla1, "N") & "," & DBSet(PrecioTalla2, "N") & "," & DBSet(ZonaTalla, "N") & "),"
            
            Rs8.MoveNext
        Wend
        
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute Sql & CadValues
        End If
        Set Rs8 = Nothing
    
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        Mens = Mens & " " & Err.Description
        FacturacionTallaPreviaESCALONA = False
    Else
        FacturacionTallaPreviaESCALONA = True
    End If
    Me.pb5.visible = False
End Function

Private Sub MostrarContadoresANoFacturar(cTabla As String, cSelect As String)
Dim Sql As String


    Sql = "select rpozos.hidrante from " & cTabla & " where (rpozos.consumo < " & DBSet(vParamAplic.ConsumoMinPOZ, "N") & " or rpozos.consumo > " & DBSet(vParamAplic.ConsumoMaxPOZ, "N") & ") "
    If cSelect <> "" Then Sql = Sql & " and " & cSelect
    
    If TotalRegistrosConsulta(Sql) <> 0 Then
        
        Set frmMens3 = New frmMensajes
        
        frmMens3.OpcionMensaje = 50
        frmMens3.cadWHERE = " and rpozos.hidrante in (" & Sql & ")"
        frmMens3.Show vbModal
    
        Set frmMens3 = Nothing
        
    Else
    
        Continuar = True
    
    End If
    
End Sub



Private Sub InsertarTemporal(cadWHERE As String, cadSelect As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer
Dim Sql2 As String

    
    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    'seleccionamos todos los socios a los que queremos enviar e-mail
    Sql = "SELECT distinct " & vUsu.Codigo & ", rsocios.codsocio, rrecibpozos.codtipom, rrecibpozos.numfactu, rrecibpozos.fecfactu  from rsocios, rrecibpozos where rrecibpozos.codsocio in (" & cadWHERE & ")"
    Sql = Sql & " and rsocios.codsocio = rrecibpozos.codsocio "
    Sql = Sql & " and " & cadSelect
    
    Sql2 = "insert into tmpinformes (codusu, codigo1, nombre1, importe1, fecha1) " & Sql
    conn.Execute Sql2

End Sub



Private Function TotalSocios(cTabla As String, cWhere As String) As Long
Dim Sql As String

    TotalSocios = 0
    
    Sql = "SELECT  count(distinct rsocios.codsocio) "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If

    TotalSocios = TotalRegistros(Sql)

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
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Cad2 As String
Dim Cad3 As String
Dim CadValues As String
Dim CadInsert As String
Dim Contador As String
Dim Nregs As Integer
Dim Fecha As Date
Dim DDCC As Integer
Dim CC As String
Dim Ent As String ' Entidad
Dim Suc As String ' Oficina
Dim DC As String ' Digitos de control
Dim I, i2, i3, i4 As Integer
Dim NumCC As String ' Número de cuenta propiamente dicho
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
    
    
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Label2(102).visible = True
    DoEvents
    
    CadInsert = "insert into tmpinformes (codusu, codigo1, nombre1, nombre2)  VALUES "
    
    Sql = "select codsocio,codbanco,codsucur,digcontr,cuentaba, iban from rsocios "
    Sql = Sql & "where cuentaba <> '8888888888' "
    If cWhere <> "" Then Sql = Sql & " and  " & cWhere
    Sql = Sql & " order by codsocio "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

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
                
                '-- Calculamos el primer dígito de control
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
                
                '-- Calculamos el segundo dígito de control
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
        conn.Execute CadInsert & CadValues
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
    MuestraError Err.Description, "Cargar Temporal CCC Erróneas", Err.Description
End Function


Private Function FacturacionConsumoMantaESCALONA(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    'hacemos una factura por socio campo
    Sql = "SELECT rcampos.codsocio, rcampos.codcampo, rpozauxmanta.nroimpresion, sum(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) hanegada  "
    Sql = Sql & " FROM  (" & cTabla & ") INNER JOIN rpozauxmanta On rcampos.codcampo = rpozauxmanta.codcampo "

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    
    Sql = Sql & " group by 1, 2, 3 having sum(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2)) <> 0 "
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by codsocio, codcampo "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    B = True
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And B
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
    
            TotalFac = Round2(Acciones * CCur(ImporteSinFormato(txtCodigo(112).Text)), 2)
            
        
            '[Monica]14/03/2012, descomponemos el total que lleva el iva incluido
            baseimpo = Round2(TotalFac / (1 + (PorcIva / 100)), 2)
            ImpoIva = TotalFac - baseimpo
        
            IncrementarProgresNew pb7, 1
            
            'insertar en la tabla de recibos de pozos
            Sql = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
            Sql = Sql & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
            Sql = Sql & "concepto, contabilizado, porcdto, impdto, precio) "
            Sql = Sql & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
            Sql = Sql & ValorNulo & "," & DBSet(baseimpo, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql = Sql & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(txtCodigo(113).Text, "T") & ",0,"
            Sql = Sql & DBSet(0, "N") & ","
            Sql = Sql & DBSet(0, "N") & ","
            Sql = Sql & DBSet(CCur(ImporteSinFormato(txtCodigo(112).Text)), "N") & ")"
            
            conn.Execute Sql
                
                
            ' Introducimos en la tabla de lineas de campos que intervienen en la factura para la impresion
            ' SOLO HABRA UN CAMPO
            Sql = "SELECT rcampos.codcampo, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ", 2) hanegada "
            Sql = Sql & " FROM  " & cTabla '& ") INNER JOIN rcampos ON rpozos.codcampo = rcampos.codcampo"
            If cWhere <> "" Then
                Sql = Sql & " WHERE " & cWhere
                Sql = Sql & " and rcampos.codsocio = " & DBSet(Rs!Codsocio, "N")
                Sql = Sql & " and rcampos.codcampo = " & DBSet(Rs!codcampo, "N")
            Else
                Sql = Sql & " where rcampos.codsocio = " & DBSet(Rs!Codsocio, "N")
                Sql = Sql & " and rcampos.codcampo = " & DBSet(Rs!codcampo, "N")
            End If
                
            Set Rs8 = New ADODB.Recordset
            Rs8.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Sql = "insert into rrecibpozos_cam (codtipom, numfactu, fecfactu, codcampo, hanegada, precio1, codzonas, poligono, parcela, subparce) values  "
            CadValues = ""
            While Not Rs8.EOF
                CadValues = CadValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                CadValues = CadValues & DBSet(Rs8!codcampo, "N") & "," & DBSet(Rs8!hanegada, "N") & "," & DBSet(txtCodigo(112).Text, "N") & ","
                CadValues = CadValues & DBSet(Rs8!codzonas, "N") & "," & DBSet(Rs8!Poligono, "N") & "," & DBSet(Rs8!Parcela, "N") & "," & DBSet(Rs8!SubParce, "T")
                CadValues = CadValues & "),"
                Rs8.MoveNext
            Wend
            If CadValues <> "" Then
                CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
                conn.Execute Sql & CadValues
            End If
            Set Rs8 = Nothing
                
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        Next I
        Rs.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionConsumoMantaESCALONA = False
    Else
        conn.CommitTrans
        FacturacionConsumoMantaESCALONA = True
    End If
End Function


Private Function FacturacionConsumoMantaESCALONANew(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Mens As String) As Boolean
Dim Sql As String
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
Dim B As Boolean
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
    
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
'[Monica]20/04/2015: ya no actualizamos nada ¿?
'    ' actualizamos el precio de recibo a manta
'    Sql = "update rtipopozos set imporcuotahda = " & DBSet(CCur(ImporteSinFormato(txtCodigo(112).Text)), "N")
'    Sql = Sql & " where codpozo = 1"
'    conn.Execute Sql


    'hacemos una factura por socio campo
    Sql = "SELECT rcampos.codsocio, rcampos.codcampo, rpozauxmanta.nroimpresion, rpozauxmanta.hanegadas, rcampos.codzonas, rcampos.poligono, rcampos.parcela, rcampos.subparce, rzonas.preciomanta  "
    Sql = Sql & " FROM  (" & cTabla & ") INNER JOIN rpozauxmanta On rcampos.codcampo = rpozauxmanta.codcampo "

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    
    Sql = Sql & " group by 1, 2, 3 having rpozauxmanta.hanegadas <> 0 "
    
    ' ordenado por socio, hidrante
    Sql = Sql & " order by codsocio, codcampo "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    B = True
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And B
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
        
            IncrementarProgresNew pb7, 1
            
            'insertar en la tabla de tickets de pozos
            Sql = "insert into rpozticketsmanta (numalbar,fecalbar,codsocio,codcampo,hanegada,precio1,importe,codzonas,poligono,parcela,subparce,fecriego,fecpago,concepto) "
            Sql = Sql & " values (" & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
            Sql = Sql & DBSet(Rs!codcampo, "N") & "," & DBSet(Rs!Hanegadas, "N") & "," & DBSet(Precio, "N") & "," ' DBSet(CCur(ImporteSinFormato(txtCodigo(112).Text)), "N") & ","
            Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!codzonas, "N") & "," & DBSet(Rs!Poligono, "N") & "," & DBSet(Rs!Parcela, "N") & "," & DBSet(Rs!SubParce, "T")
            Sql = Sql & "," & ValorNulo & "," & ValorNulo & "," & DBSet(txtCodigo(113).Text, "T") & ")"
            
            conn.Execute Sql
                
            Sql = "insert into tmpinformes (codusu, nombre1, importe1) values ( " & vUsu.Codigo & ",'ALV'," & DBSet(numfactu, "N") & ")"
            conn.Execute Sql
                
                
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        Next I
        Rs.MoveNext
    Wend
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        Mens = Mens & " " & Err.Description
        conn.RollbackTrans
        FacturacionConsumoMantaESCALONANew = False
    Else
        conn.CommitTrans
        FacturacionConsumoMantaESCALONANew = True
    End If
End Function

Private Function RellenaABlancos(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Space(longitud)
    If PorLaDerecha Then
        cad = cadena & cad
        RellenaABlancos = Left(cad, longitud)
    Else
        cad = cad & cadena
        RellenaABlancos = Right(cad, longitud)
    End If
    
End Function


