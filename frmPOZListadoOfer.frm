VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmPOZListadoOfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   10785
   Icon            =   "frmPOZListadoOfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameEtiqProv 
      Height          =   6975
      Left            =   0
      TabIndex        =   14
      Top             =   30
      Width           =   7035
      Begin VB.Frame Frame13 
         Height          =   900
         Left            =   270
         TabIndex        =   71
         Top             =   4830
         Width           =   6405
         Begin VB.OptionButton OptMail 
            Caption         =   "Imprimir Todos"
            Height          =   255
            Index           =   1
            Left            =   4680
            TabIndex        =   73
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptMail 
            Caption         =   "Enviar por e-mail e imprimir a los socios sin correo"
            Height          =   255
            Index           =   0
            Left            =   270
            TabIndex        =   72
            Top             =   360
            Width           =   3885
         End
      End
      Begin VB.Frame FrameTipoSocio 
         Caption         =   "Tipo Pago"
         ForeColor       =   &H00972E0B&
         Height          =   1215
         Left            =   4050
         TabIndex        =   10
         Top             =   1650
         Width           =   2145
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   2
            Left            =   420
            TabIndex        =   69
            Top             =   840
            Width           =   885
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Efecto"
            Height          =   225
            Index           =   1
            Left            =   420
            TabIndex        =   68
            Top             =   570
            Width           =   1005
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Contado"
            Height          =   225
            Index           =   0
            Left            =   420
            TabIndex        =   67
            Top             =   300
            Width           =   1005
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1140
         Width           =   2070
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2505
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2160
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1140
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   1
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1470
         Width           =   830
      End
      Begin VB.CheckBox chkMail 
         Caption         =   "Solo con marca de Correo"
         Height          =   345
         Index           =   2
         Left            =   1650
         TabIndex        =   6
         Top             =   5460
         Width           =   2505
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   5850
         Top             =   5550
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5490
         TabIndex        =   13
         Top             =   6420
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarCartaRec 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4410
         TabIndex        =   12
         Top             =   6420
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   270
         TabIndex        =   21
         Top             =   3990
         Width           =   6255
         Begin VB.CheckBox chkMail 
            Caption         =   "Enviar por e-mail"
            Height          =   345
            Index           =   0
            Left            =   4260
            TabIndex        =   11
            Top             =   1050
            Width           =   1935
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1470
            MaxLength       =   10
            TabIndex        =   8
            Top             =   450
            Width           =   1005
         End
         Begin VB.Frame Frame3 
            Caption         =   "e-Mail"
            Enabled         =   0   'False
            Height          =   780
            Left            =   180
            TabIndex        =   24
            Top             =   870
            Visible         =   0   'False
            Width           =   1695
            Begin VB.OptionButton OptMailAdm 
               Caption         =   "Administración"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   26
               Top             =   210
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton OptMailCom 
               Caption         =   "Compras"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   460
               Width           =   1335
            End
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   63
            Left            =   2550
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "Text5"
            Top             =   90
            Width           =   3735
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   63
            Left            =   1470
            MaxLength       =   6
            TabIndex        =   7
            Top             =   105
            Width           =   1005
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1170
            Picture         =   "frmPOZListadoOfer.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label9 
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
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   465
            Width           =   435
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   39
            Left            =   1170
            ToolTipText     =   "Buscar carta"
            Top             =   105
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Carta"
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
            Left            =   240
            TabIndex        =   23
            Top             =   120
            Width           =   405
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   60
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   3105
         Width           =   3735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   4
         Top             =   3105
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   61
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   3450
         Width           =   3735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3450
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   66
         Top             =   840
         Width           =   1815
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   1500
         Picture         =   "frmPOZListadoOfer.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   2490
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   840
         TabIndex        =   65
         Top             =   2490
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1500
         Picture         =   "frmPOZListadoOfer.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   840
         TabIndex        =   64
         Top             =   2160
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
         Index           =   16
         Left            =   480
         TabIndex        =   63
         Top             =   1860
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
         Left            =   510
         TabIndex        =   62
         Top             =   900
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   915
         TabIndex        =   61
         Top             =   1170
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   915
         TabIndex        =   60
         Top             =   1470
         Width           =   420
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
         Left            =   600
         TabIndex        =   59
         Top             =   6180
         Width           =   3705
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   4170
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   5460
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Index           =   5
         Left            =   480
         TabIndex        =   20
         Top             =   2820
         Width           =   375
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   37
         Left            =   1440
         ToolTipText     =   "Buscar socio"
         Top             =   3105
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   19
         Top             =   3105
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   38
         Left            =   1440
         ToolTipText     =   "Buscar socio"
         Top             =   3450
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   840
         TabIndex        =   18
         Top             =   3450
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "Carta de Reclamación"
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
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Top             =   270
         Width           =   5340
      End
   End
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   30
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Preparando datos envio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   360
         TabIndex        =   29
         Top             =   840
         Width           =   5805
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6435
      Top             =   5985
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameEnvioFacMail 
      Height          =   6015
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   10215
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   107
         Left            =   3840
         MaxLength       =   7
         TabIndex        =   39
         Tag             =   "Nº Factura|N|S|||rfactsoc|numfactu|0000000|S|"
         Top             =   3660
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   106
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   38
         Tag             =   "Nº Factura|N|S|||rfactsoc|numfactu|0000000|S|"
         Text            =   "wwwwwww"
         Top             =   3660
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   108
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   36
         Top             =   2778
         Width           =   1080
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   109
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   37
         Top             =   2778
         Width           =   1080
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   9000
         TabIndex        =   47
         Top             =   5370
         Width           =   975
      End
      Begin VB.CheckBox chkMail 
         Caption         =   "Copia remitente"
         Height          =   255
         Index           =   3
         Left            =   5610
         TabIndex        =   41
         Top             =   1830
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   0
         Left            =   5640
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   2760
         Width           =   4335
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   110
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text5"
         Top             =   1185
         Width           =   3015
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   110
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   33
         Top             =   1185
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   111
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   111
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   35
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   1665
         Index           =   1
         Left            =   5640
         MultiLine       =   -1  'True
         TabIndex        =   43
         Text            =   "frmPOZListadoOfer.frx":01AD
         Top             =   3480
         Width           =   4335
      End
      Begin VB.CommandButton cmdEnvioMail 
         Caption         =   "&Enviar"
         Height          =   375
         Left            =   7920
         TabIndex        =   45
         Top             =   5370
         Width           =   975
      End
      Begin VB.ListBox ListTipoMov 
         Height          =   960
         Index           =   1000
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   40
         Top             =   4170
         Width           =   4095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   3360
         TabIndex        =   58
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   600
         TabIndex        =   57
         Top             =   3645
         Width           =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         Index           =   15
         Left            =   240
         TabIndex        =   56
         Top             =   3360
         Width           =   780
      End
      Begin VB.Label Label14 
         Caption         =   "Envio facturas por mail"
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
         Index           =   16
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   600
         TabIndex        =   54
         Top             =   2823
         Width           =   450
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fact."
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
         Index           =   1
         Left            =   240
         TabIndex        =   53
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   33
         Left            =   1080
         Picture         =   "frmPOZListadoOfer.frx":01B3
         Top             =   2800
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   34
         Left            =   3600
         Picture         =   "frmPOZListadoOfer.frx":023E
         Top             =   2800
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   3120
         TabIndex        =   52
         Top             =   2823
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Asunto"
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
         Left            =   5640
         TabIndex        =   51
         Top             =   2430
         Width           =   510
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   56
         Left            =   1080
         Top             =   1185
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Index           =   32
         Left            =   240
         TabIndex        =   50
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   600
         TabIndex        =   49
         Top             =   1185
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   57
         Left            =   1080
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   600
         TabIndex        =   48
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   20
         Left            =   240
         TabIndex        =   46
         Top             =   4110
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje"
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
         Left            =   5640
         TabIndex        =   44
         Top             =   3180
         Width           =   600
      End
   End
   Begin VB.CheckBox chkMail 
      Caption         =   "Enviar SMS"
      Height          =   345
      Index           =   1
      Left            =   4920
      TabIndex        =   70
      Top             =   6630
      Width           =   1935
   End
End
Attribute VB_Name = "frmPOZListadoOfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
    '(ver opciones en frmListado)
    ' 1 : carta de reclamacion
        
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta/pedido a imprimir

Public CodClien As String 'Para seleccionar inicialmente las ofertas del Cliente
                          'en el listado de Recordatorio de Ofertas y de Valoracion de Ofertas

Public FecEntre As String 'Para pasar inicialmente la fecha de entrega de la Oferta que se va a pasar a pedido
                          'como la fecha de entega del PEdido
                          
Private NomTabla As String
Private NomTablaLin As String

Private WithEvents frmCar As frmCartasSocio
Attribute frmCar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSeccion
Attribute frmSec.VB_VarHelpID = -1


'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
'Private WithEvents frmCP As frmCPostal 'codigo postal
Private WithEvents frmMen As frmMensajes  'Form Mensajes para mostrar las etiquetas a imprimir
Attribute frmMen.VB_VarHelpID = -1
Private WithEvents frmMen2 As frmMensajes  'Form Mensajes para mostrar las etiquetas a imprimir
Attribute frmMen2.VB_VarHelpID = -1
Private WithEvents frmMen3 As frmMensajes  'Ficheros que vamos a adjuntar
Attribute frmMen3.VB_VarHelpID = -1


'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'cadena con los parametros q se pasan a Crystal R.
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------

Dim indCodigo As Byte 'indice para txtCodigo
Dim Codigo As String 'Código para FormulaSelection de Crystal Report

Dim PrimeraVez As Boolean
Dim IndRptReport As Integer

Dim Documento As String

'Indicamos si ejecuta el rpt o envia unicamente un fichero
Dim ImpresionNormal As Boolean

Dim TipCod As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub chkmail_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub



'******************************************************
'
'[Monica]31/10/2013: cambio las opciones se envian si hay correo y se imprimen si no hay correo o solo se imprimen
'
' he guardado lo anterior en
'                         cmdAceptarCartaRecANTES_Click
'******************************************************

Private Sub cmdAceptarCartaRec_Click()
'1: Listado para cartas de reclamacion a proveedor
Dim campo As String
Dim Tabla As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

Dim OK As Boolean

Dim Situacion As String
Dim Tipos As String
Dim CodTipom As String
Dim cDesde As String
Dim cHasta As String
Dim nDesde As String
Dim nHasta As String

Dim cadSelect1 As String
Dim cadFormula1 As String
Dim cadNombreRPT As String

Dim ConSubInforme As String
Dim SQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    Tabla = "rrecibpozos"
    
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
    cDesde = Trim(txtcodigo(60).Text)
    cHasta = Trim(txtcodigo(61).Text)
    nDesde = txtNombre(60).Text
    nHasta = txtNombre(61).Text
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
    
    'si es listado de CARTAS/eMAIL a socios comprobar que se ha seleccionado
    'una carta para imprimir
    If txtcodigo(63).Text = "" Then
        MsgBox "Debe seleccionar una carta para imprimir.", vbInformation
        Exit Sub
    End If
    
    'Parametro cod. carta
    cadParam = "|pCodCarta= " & txtcodigo(63).Text & "|"
    numParam = numParam + 1
    
    'Parametro fecha
    cadParam = cadParam & "|pFecha= """ & txtcodigo(0).Text & """|"
    numParam = numParam + 1
    
    
    '[Monica]11/11/2011: añadimos los socios que esten dados de baja en todas las secciones
    ' solo se sacan los socios que no esten dados de baja
    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null ") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "isnull({rsocios.fechabaja})") Then Exit Sub

    '[Monica]23/11/2012: si es escalona o utxera seleccionamos que tipo de socio
    If Option1(0).Value Then    ' solo contado
        If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
    End If
    If Option1(1).Value Then    ' solo efecto
        If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba<>""8888888888""") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}<>""8888888888""") Then Exit Sub
    End If
'   no hacemos nada
'    If Option1(2).Value Then
'    End If
    
    '[Monica]08/11/2012: solo los socios que no tengan situacion de bloqueo
    Situacion = SituacionesBloqueo
    If Situacion <> "" Then
        Situacion = Mid(Situacion, 1, Len(Situacion) - 1)
        If Not AnyadirAFormula(cadFormula, "not ({rsocios.codsitua} in [" & Situacion & "])") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "not rsocios.codsitua in (" & Situacion & ")") Then Exit Sub
    End If
    
    
    'Nombre fichero .rpt a Imprimir
    nomRPT = "rSocioCarta.rpt" '"rComProveCarta.rpt"
    Titulo = "Cartas Reclamación a Socios"
    
    indRPT = 61 'Personalizacion de la carta a socios
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
     '[Monica]19/10/2012: nueva variable para indicar que se pasa por visreport o no ImpresionNormal
    ImpresionNormal = True
    Documento = DevuelveDesdeBDNew(cAgro, "scartas", "documrpt", "codcarta", txtcodigo(63).Text, "N")
    
    '[Monica]19/07/2013: dejo introducir una carta que hay creado el usuario
    If Documento <> "" Then
        nomDocu = Documento
    End If
      
    'Nombre fichero .rpt a Imprimir
    nomRPT = nomDocu
    
    conSubRPT = True
    
    Tabla = "rsocios inner join rrecibpozos on rsocios.codsocio = rrecibpozos.codsocio"
    
    'ver si hay registros seleccionados para mostrar en el informe
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
        
        
    Set frmMen = New frmMensajes
    frmMen.cadwhere = cadSelect
    frmMen.OpcionMensaje = 42 'Socios
    frmMen.Show vbModal
    Set frmMen = Nothing
    If cadSelect = "" Then Exit Sub
    
    '[Monica]31/10/2013: cambio las opciones se envian si hay correo y se imprimen si no hay correo o solo se imprimen
    '******************************************************
    
    ' si es un correo electronico miramos solo los que tienen mail
    If OptMail(0).Value Then
        
        cadSelect = QuitarCaracterACadena(cadSelect, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
        cadSelect = QuitarCaracterACadena(cadSelect, "_1")
        
        cadSelect1 = cadSelect
        cadFormula1 = cadFormula
        
        If Not AnyadirAFormula(cadSelect, "not rsocios.maisocio is null and rsocios.maisocio<>''") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "not isnull({rsocios.maisocio}) and {rsocios.maisocio}<>''") Then Exit Sub
    
        cadNombreRPT = nomDocu
        ConSubInforme = True
    
        SQL = "select count(*) from " & Tabla & " where " & cadSelect
    
        If TotalRegistros(SQL) <> 0 Then
            'Enviarlo por e-mail
            IndRptReport = indRPT
            EnviarEMailMulti cadSelect, Titulo, nomDocu, Tabla ' "rSocioCarta.rpt", Tabla  'email para socios
        Else
            If MsgBox("No hay socios a enviar carta por email. ¿ Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
        End If
    
        If Not AnyadirAFormula(cadSelect1, "(rsocios.maisocio is null or rsocios.maisocio='')") Then Exit Sub
        If Not AnyadirAFormula(cadFormula1, "(isnull({rsocios.maisocio}) or {rsocios.maisocio}='')") Then Exit Sub
    
        SQL = "select count(*) from " & Tabla & " where " & cadSelect1
        
        If TotalRegistros(SQL) <> 0 Then
            cadFormula = cadFormula1
            LlamarImprimir
        Else
            MsgBox "No hay Socios para imprimir cartas.", vbExclamation
        End If
    
    Else
        If HayRegParaInforme(Tabla, cadSelect) Then
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            LlamarImprimir
        End If
    End If
       
  '******************************************************
      
      
'    '[Monica]19/10/2012: nueva variable para indicar que se pasa por visreport o no ImpresionNormal
'    ImpresionNormal = True
'    Documento = DevuelveDesdeBDNew(cAgro, "scartas", "documrpt", "codcarta", txtcodigo(63).Text, "N")
'
'    '[Monica]19/07/2013: dejo introducir una carta que hay creado el usuario
'    If Documento <> "" Then
'        nomDocu = Documento
'    End If
'
'    'Nombre fichero .rpt a Imprimir
'    nomRPT = nomDocu
'
'    conSubRPT = True
'
'
'
'    Tabla = "rsocios inner join rrecibpozos on rsocios.codsocio = rrecibpozos.codsocio"
'
'    'ver si hay registros seleccionados para mostrar en el informe
'    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
'
'    Set frmMen = New frmMensajes
'    frmMen.cadwhere = cadSelect
'    frmMen.OpcionMensaje = 42 'Socios
'    frmMen.Show vbModal
'    Set frmMen = Nothing
'    If cadSelect = "" Then Exit Sub
'
'    If OpcionListado = 1 And Me.chkMail(0).Value = 1 Then
'        'Enviarlo por e-mail
'        IndRptReport = indRPT
'        EnviarEMailMulti cadSelect, Titulo, nomDocu, Tabla ' "rSocioCarta.rpt", Tabla  'email para socios
'        cmdCancel_Click (9)
'    Else
'        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
'        LlamarImprimir
'    End If

End Sub

Private Sub cmdEnvioMail_Click()
Dim RS As ADODB.Recordset

    'El proceso constara de varias fases.
    'Fase 1: Montar el select y ver si hay registros
    'Fase 2: Preparar carpetas para los pdf
    'Fase 3: Generar para cada factura (una a una) del select su pdf
    'Fase 4: Enviar por mail, adjuntando los archivos correspondientes
    If Text1(0).Text = "" Then
        MsgBox "Ponga el asunto", vbExclamation
        Exit Sub
    End If
    
    'AHora pongo los tipo de facturas
    cadFormula = ""
    cadSelect = ""  'ME dira si estan todas o no
    For indCodigo = 0 To Me.ListTipoMov(1000).ListCount - 1
        If Me.ListTipoMov(1000).Selected(indCodigo) Then
            'Esta checkeado
            cadFormula = cadFormula & " OR rfactsoc.codtipom = '" & Trim(Mid(ListTipoMov(1000).List(indCodigo), 1, 3)) & "'"
        Else
            cadSelect = "NO"
        End If
    Next indCodigo
    
    If cadFormula = "" Then
        MsgBox "Seleccione algun tipo de factura", vbExclamation
        Exit Sub
    Else
        cadFormula = Mid(cadFormula, 4)
    End If
    'En notabla tendre

    NomTabla = "(" & cadFormula & ")"
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    InicializarVbles
    cadFormula = ""
    cadSelect = ""
    
'    'Cadena para seleccion D/H Letra Serie
'    '--------------------------------------------
'    If txtCodigo(0).Text <> "" Or txtCodigo(1).Text <> "" Then
'        Codigo = "facturas.codtipom"
'        'Parametro Desde/Hasta Letra Serie
'        If Not PonerDesdeHasta(Codigo, "T", 0, 1, "") Then Exit Sub
'    End If
        
    'Cadena para seleccion D/H Factura
    '--------------------------------------------
    If txtcodigo(106).Text <> "" Or txtcodigo(107).Text <> "" Then
        Codigo = "rfactsoc.numfactu"
        If Not PonerDesdeHasta(Codigo, "N", 106, 107, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Fecha
    '--------------------------------------------
    If txtcodigo(108).Text <> "" Or txtcodigo(109).Text <> "" Then
        Codigo = "rfactsoc.fecfactu"
        If Not PonerDesdeHasta(Codigo, "F", 108, 109, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Socio
    '--------------------------------------------
    If txtcodigo(110).Text <> "" Or txtcodigo(111).Text <> "" Then
        Codigo = "rfactsoc.codsocio"
        If Not PonerDesdeHasta(Codigo, "N", 110, 111, "") Then Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Eliminamos temporales
    conn.Execute "DELETE from tmpinformes where codusu =" & vUsu.Codigo
    
    If cadSelect <> "" Then cadSelect = cadSelect & " AND "
    cadSelect = cadSelect & NomTabla
    cadSelect = " WHERE " & cadSelect

    
    Set RS = New ADODB.Recordset
    DoEvents
        
    'Ahora insertare en la tabla temporal tmpinformes las facturas que voy a generar pdf
    '                                         codsocio,numfactu,codtipom,fecfactu,totalfac,esliqcomplem
    Codigo = "insert into tmpinformes (codusu,codigo1,importe1, nombre1, fecha1, importe2,campo1) "
    Codigo = Codigo & " values ( " & vUsu.Codigo & ","
    
    If Not PrepararCarpetasEnvioMail Then Exit Sub
        
    Screen.MousePointer = vbHourglass

    'Vamos a meter todas las facturas en la tabla temporal para comprobar si tienen mail
    'los clientes
    
    NomTabla = "Select codtipom,numfactu,codsocio,fecfactu,totalfac,esliqcomplem from rfactsoc  " & cadSelect
    'El orden vamos a hacerlo por: Tipo documento
    NomTabla = NomTabla & " ORDER BY codtipom, numfactu, fecfactu "
    RS.Open NomTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not RS.EOF
        NomTabla = RS!Codsocio & "," & RS!numfactu & ",'" & Trim(RS!CodTipom) & "','" & Format(RS!fecfactu, FormatoFecha)
        NomTabla = NomTabla & "'," & TransformaComasPuntos(CStr(DBLet(RS!TotalFac, "N"))) & "," & DBLet(RS!Esliqcomplem, "N") & ")"
        conn.Execute Codigo & NomTabla
        NumRegElim = NumRegElim + 1
        RS.MoveNext
    Wend
    RS.Close
    
    
    If NumRegElim = 0 Then
        MsgBox "Ningun dato a enviar por mail", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Numero de registros
    NomTabla = NumRegElim
    
    cadSelect = "Select codsocio,maisocio "
    cadSelect = cadSelect & " as email from tmpinformes,rsocios where codusu = " & vUsu.Codigo & " and codsocio=codigo1"
    cadSelect = cadSelect & " group by codsocio having email is null"
    RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not RS.EOF
        NumRegElim = NumRegElim + 1
        RS.MoveNext
    Wend
    RS.Close
    
    If NumRegElim > 0 Then
        If MsgBox("Tiene socio sin mail. Continuar sin sus datos?", vbQuestion + vbYesNo) = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
            
        'Si no salimos borramos
        RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cadSelect = "DELETE from tmpinformes where codusu =" & vSesion.Codigo & " and codigo1 ="
        While Not RS.EOF
            conn.Execute cadSelect & RS!Codsocio
            RS.MoveNext
        Wend
        RS.Close
        
        
        cadSelect = "Select count(*) from tmpinformes where codusu =" & vSesion.Codigo
        RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then NumRegElim = DBLet(RS.Fields(0), "N")
            
        End If
        RS.Close
        
        If NumRegElim = 0 Then
            'NO hay datos para enviar
            
            Screen.MousePointer = vbDefault
            MsgBox "No hay datos para enviar por mail", vbExclamation
            Exit Sub
        Else
            cadSelect = "Hay " & NumRegElim & " facturas para enviar por mail." & vbCrLf & "¿Continuar?"
            If MsgBox(cadSelect, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
        End If
        If NumRegElim = 0 Then
            Set RS = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        NomTabla = NumRegElim
    
    End If
        
    PonerTamnyosMail True
    MDIppal.visible = False
    'Voy arriesgar.
    'Confio en que no envien por mail mas de 32000 facturas (un integer)
    Label14(22).Caption = "Preparando datos"
    Me.ProgressBar1.Max = CInt(NomTabla)
    Me.ProgressBar1.Value = 0
    
    
    
    NumRegElim = 0
    If GeneracionEnvioMail(RS) Then NumRegElim = 1
        
    
    'Si ha ido todo bien entonces numregelim=1
    If NumRegElim = 1 Then
        cadSelect = "Select nomsocio, maisocio"
        cadSelect = cadSelect & " as email,tmpinformes.* from tmpinformes,rsocios where codusu = " & vUsu.Codigo & " and codsocio=codigo1"
'        cadSelect = cadSelect & " group by codclien having email is null"

        
        frmEMail.DatosEnvio = Text1(0).Text & "|" & Text1(1).Text & "|" & Abs(0) & "|" & cadSelect & "|"
        frmEMail.Opcion = 4 'Multienvio de facturacion
        frmEMail.Show vbModal
        
        
        'Para tranquilizar las pantallas, borrar los ficheros generados
        'Confio en que no envien por mail mas de 32000 facturas (un integer)
        Label14(22).Caption = "Restaurando ...."
        Me.ProgressBar1.visible = False
        Me.Refresh
        DoEvents
        espera 1
        PrepararCarpetasEnvioMail
        Me.ProgressBar1.visible = True
        
        
    End If
    
    
    
    
    'Es para evitar la cantidad de pantallas abriendose y cerrandose
    Me.visible = False
    PonerTamnyosMail False
    espera 1
    Unload Me
    MDIppal.Show

    Screen.MousePointer = vbDefault

End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "El check de 'Solo con Marca de Correo' indica que se generará el " & vbCrLf & _
                      "documento únicamente a los socios que tengan marcada la casilla " & vbCrLf & _
                      "'Correo' de su ficha. " & vbCrLf
                      
            If OpcionListado = 306 Then
                      vCadena = vCadena & vbCrLf & "No tendrá ningún efecto en el caso de que estemos mandando un sms. " & vbCrLf
            End If
            
            vCadena = vCadena & "" & vbCrLf & vbCrLf
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub



Private Sub cmdAceptarCartaRec2_Click()
'1: Listado para cartas de reclamacion a proveedor
Dim campo As String
Dim Tabla As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

Dim OK As Boolean

Dim Situacion As String
Dim Tipos As String
Dim CodTipom As String
Dim cDesde As String
Dim cHasta As String
Dim nDesde As String
Dim nHasta As String


    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    Tabla = "rrecibpozos"
    
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
    cDesde = Trim(txtcodigo(60).Text)
    cHasta = Trim(txtcodigo(61).Text)
    nDesde = txtNombre(60).Text
    nHasta = txtNombre(61).Text
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
    
    'si es listado de CARTAS/eMAIL a socios comprobar que se ha seleccionado
    'una carta para imprimir
    If txtcodigo(63).Text = "" Then
        MsgBox "Debe seleccionar una carta para imprimir.", vbInformation
        Exit Sub
    End If
    
    'Parametro cod. carta
    cadParam = "|pCodCarta= " & txtcodigo(63).Text & "|"
    numParam = numParam + 1
    
    'Parametro fecha
    cadParam = cadParam & "|pFecha= """ & txtcodigo(0).Text & """|"
    numParam = numParam + 1
    
    'Nombre fichero .rpt a Imprimir
    nomRPT = "rSocioCarta.rpt" '"rComProveCarta.rpt"
    Titulo = "Cartas Reclamación a Socios"
    
    indRPT = 61 'Personalizacion de la carta a socios
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    '[Monica]19/10/2012: nueva variable para indicar que se pasa por visreport o no ImpresionNormal
    ImpresionNormal = True
    Documento = DevuelveDesdeBDNew(cAgro, "scartas", "documrpt", "codcarta", txtcodigo(63).Text, "N")
    
    '[Monica]19/07/2013: dejo introducir una carta que hay creado el usuario
    If Documento <> "" Then
        nomDocu = Documento
    End If
      
    'Nombre fichero .rpt a Imprimir
    nomRPT = nomDocu
    
    conSubRPT = True
    
        
    '[Monica]11/11/2011: añadimos los socios que esten dados de baja en todas las secciones
    ' solo se sacan los socios que no esten dados de baja
    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null ") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "isnull({rsocios.fechabaja})") Then Exit Sub

    '[Monica]23/11/2012: si es escalona o utxera seleccionamos que tipo de socio
    If Option1(0).Value Then    ' solo contado
        If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
    End If
    If Option1(1).Value Then    ' solo efecto
        If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba<>""8888888888""") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}<>""8888888888""") Then Exit Sub
    End If
'   no hacemos nada
'    If Option1(2).Value Then
'    End If
    
    '[Monica]08/11/2012: solo los socios que no tengan situacion de bloqueo
    Situacion = SituacionesBloqueo
    If Situacion <> "" Then
        Situacion = Mid(Situacion, 1, Len(Situacion) - 1)
        If Not AnyadirAFormula(cadFormula, "not {rsocios.codsitua} in [" & Situacion & "]") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "not rsocios.codsitua in (" & Situacion & ")") Then Exit Sub
    End If
    
    ' si es un correo electronico miramos solo los que tienen mail
    If chkMail(0).Value = 1 Then
        If Not AnyadirAFormula(cadSelect, "not {rsocios.maisocio} is null and {rsocios.maisocio}<>''") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "not isnull({rsocios.maisocio}) and {rsocios.maisocio}<>''") Then Exit Sub
    End If
    ' hasta aqui
    
    '=========================================================================================
    '[Monica]20/01/2012: Añadida la condicion de marca de correo del socio
    '                    Sólo se tendrá en cuenta en cartas y etiquetas. No en los sms
    '=========================================================================================
    If chkMail(0).Value = 0 Then
        If chkMail(2).Value = 1 Then
            If Not AnyadirAFormula(cadSelect, "{rsocios.correo} = 1") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.correo} = 1") Then Exit Sub
        End If
    End If
    
    Tabla = "rsocios inner join rrecibpozos on rsocios.codsocio = rrecibpozos.codsocio"
    
    'ver si hay registros seleccionados para mostrar en el informe
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    Set frmMen = New frmMensajes
    frmMen.cadwhere = cadSelect
    frmMen.OpcionMensaje = 42 'Socios
    frmMen.Show vbModal
    Set frmMen = Nothing
    If cadSelect = "" Then Exit Sub
    
    If OpcionListado = 1 And Me.chkMail(0).Value = 1 Then
        'Enviarlo por e-mail
        IndRptReport = indRPT
        EnviarEMailMulti cadSelect, Titulo, nomDocu, Tabla ' "rSocioCarta.rpt", Tabla  'email para socios
        cmdCancel_Click (9)
    Else
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        LlamarImprimir
    End If
End Sub


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 1 ' Cartas de reclamacion
                '[Monica]01/07/2013: introducimos los valores por defecto
                txtcodigo(36).Text = Format(Now, "dd/mm/yyyy")
                txtcodigo(37).Text = Format(Now, "dd/mm/yyyy")
            
                PonerFoco txtcodigo(58)
                
            Case 315 ' envio de facturas por email
                PonerFoco txtcodigo(110)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim h As Integer, W As Integer
Dim indFrame As Single
Dim devuelve As String
    
'    'Icono del formulario
'    Me.Icon = frmPpal.Icon
'
    PrimeraVez = True
    limpiar Me
    indCodigo = 0
    NomTabla = ""

    'Ocultar todos los Frames de Formulario
    Me.FrameEtiqProv.visible = False
    Me.FrameEnvioFacMail.visible = False
    CommitConexion
    
    CargarIconos
    
    CargarCombo
    
    Me.Option1(2).Value = True
    
    Select Case OpcionListado
        Case 1 ' 1: Cartas de reclamacion
            indFrame = 9
            h = 7155 '5325
            W = 7035
            PonerFrameVisible Me.FrameEtiqProv, True, h, W
            Me.Frame2.visible = True
            Me.Frame3.visible = (chkMail(0).Value = True)
            Me.Frame3.Enabled = (chkMail(0).Value = True)
            txtcodigo(0).Text = Format(Now, "dd/mm/yyyy")
            
            Combo1(0).ListIndex = 0
            
            txtcodigo(63).Text = vParamAplic.CartaPOZ
            txtNombre(63).Text = DevuelveDesdeBDNew(cAgro, "scartas", "descarta", "codcarta", txtcodigo(63).Text, "N")
            
        Case 315 'Envio masivo de Facturas
            indFrame = 18
            h = FrameEnvioFacMail.Height
            W = FrameEnvioFacMail.Width
            PonerFrameVisible FrameEnvioFacMail, True, h, W
        
    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = h + 350
    
End Sub

Private Sub frmCar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Cartas de Socio
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtcodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        InsertarTemporal CadenaSeleccion, cadSelect
        cadSelect = cadSelect & " and rrecibpozos.codsocio IN (" & CadenaSeleccion & ")"
    Else 'no seleccionamos ningun cliente
        cadFormula = ""
        cadSelect = ""
    End If
End Sub

Private Sub frmMen2_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        If OpcionListado = 306 Then 'Socios
            cadFormula = "{rcampos.codcampo} IN [" & CadenaSeleccion & "]"
            cadSelect = "rcampos.codcampo IN (" & CadenaSeleccion & ")"
        End If
    Else 'no seleccionamos ningun cliente
        cadFormula = ""
        cadSelect = ""
    End If
End Sub

Private Sub frmMen3_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        cadSelect = CadenaSeleccion
    Else 'no seleccionamos ningun archivo
        cadSelect = ""
    End If
End Sub



Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Secciones
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Socios
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
Dim SQL As String

    Select Case Index
        Case 39 'Cod. Carta
            SQL = DevuelveDesdeBDNew(cAgro, "scartas", "descarta", "codcarta", vParamAplic.CartaPOZ, "N")
            If SQL = "" Then
                MsgBox "No tiene en parámetros la carta de Reclamación. Revise.", vbExclamation
                Exit Sub
            End If
            indCodigo = 63
            Set frmCar = New frmCartasSocio
            frmCar.CodigoActual = txtcodigo(63).Text 'vParamAplic.CartaPOZ
            frmCar.DatosADevolverBusqueda = "0|1|"
            frmCar.Show vbModal
            Set frmCar = Nothing
            
            
        Case 37, 38 'Cod socio
            indCodigo = Index + 23

            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing

        Case 56, 57 'Cod socio
            indCodigo = Index + 54

            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing

        Case 35, 36 'cod. seccion
            indCodigo = Index + 23
            
            Set frmSec = New frmManSeccion
            frmSec.DatosADevolverBusqueda = "0|1|"
            frmSec.Show vbModal
            Set frmSec = Nothing
            
    End Select
    PonerFoco txtcodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal

   '++monica
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmF = New frmCal

    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top

    Set obj = imgFecha(Index).Container

    While imgFecha(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend

    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmF.Left = esq + imgFecha(Index).Parent.Left + 30
    frmF.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

   frmF.NovaData = Now

   Select Case Index
        Case 0
            indCodigo = 0
        Case 33, 34
            indCodigo = Index + 75
        Case 7, 8
            indCodigo = Index + 29
   End Select

   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)
End Sub


Private Sub chkMail_Click(Index As Integer)

    '[Monica]20/01/2012: el check de solo con marca de correo no hace nada
    '                    Solo tiene efecto para la condicion del where
    If Index = 2 Then Exit Sub


    If Index = 0 Then
        If chkMail(0).Value = 1 Then
            chkMail(1).Value = 0
            chkMail(1).Enabled = False
        Else
            chkMail(1).Value = 0
            chkMail(1).Enabled = True
        End If
    Else
        If chkMail(1).Value = 1 Then
            chkMail(0).Value = 0
            chkMail(0).Enabled = False
        Else
            chkMail(0).Value = 0
            chkMail(0).Enabled = True
        End If
    End If

    Frame3.visible = (chkMail(0).Value = 1)
    Frame3.Enabled = (chkMail(0).Value = 1)
    
    '[Monica]22/12/2011: si tenemos metido el numero de carta que me traiga el texto del sms
    If chkMail(1).Value = 1 And txtcodigo(63).Text <> "" Then
        txtcodigo(2).Text = DevuelveValor("select textosms from scartas where codcarta = " & DBSet(txtcodigo(63).Text, "N"))
        If txtcodigo(2).Text = "0" Then txtcodigo(2).Text = ""
    End If
    
    
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 33 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            ' Etiquetas y cartas socios
            Case 60: KEYBusqueda KeyAscii, 37 'socio desde
            Case 61: KEYBusqueda KeyAscii, 38 'socio hasta
            Case 63: KEYBusqueda KeyAscii, 39 'codigo de carta
            Case 0: KEYFecha KeyAscii, 0 'fecha envio
        
            ' envio de facturas por email
            Case 110: KEYBusqueda KeyAscii, 56 'socio desde
            Case 111: KEYBusqueda KeyAscii, 57 'socio hasta
            Case 108: KEYFecha KeyAscii, 33 'fecha desde
            Case 109: KEYFecha KeyAscii, 34 'fecha hasta
        
            Case 36: KEYFecha KeyAscii, 7 'fecha desde
            Case 37: KEYFecha KeyAscii, 8 'fecha hasta
        
        
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscarOfer_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFecha_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codcampo As String, nomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean


    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        Case 0, 108, 109, 36, 37 ' Fecha
            '[Monica]15/11/2013: He añadido el if del ponerformato y si es index = 36
            If PonerFormatoFecha(txtcodigo(Index), True) Then
                If Index = 36 Then txtcodigo(37).Text = txtcodigo(36).Text
            End If
        
        Case 1
            PonerFormatoHora txtcodigo(Index)
        
        Case 63, 64 'CARTA de la Oferta
            EsNomCod = True
            Tabla = "scartas"
            codcampo = "codcarta"
            nomCampo = "descarta"
            Formato = "000"
            Titulo = "cod. de Carta"
                    
        Case 60, 61, 110, 111 'Socio
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = txtcodigo(Index).Text
            
         Case 58, 59 'Cod. Seccion
            EsNomCod = True
            Tabla = "rseccion"
            codcampo = "codsecci"
            nomCampo = "nomsecci"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Sección"
            
        Case 38, 39 ' nro de factura
            PonerFormatoEntero txtcodigo(Index)
            
    End Select
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), Tabla, nomCampo, codcampo, TipCampo)
                If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, Formato)
                
                If Index = 63 And chkMail(1).Value = 1 Then
                    txtcodigo(2).Text = DevuelveValor("select textosms from scartas where codcarta = " & DBSet(txtcodigo(63).Text, "N"))
                    If txtcodigo(2).Text = "0" Then txtcodigo(2).Text = ""
                End If
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), Tabla, nomCampo, codcampo, TipCampo)
        End If
    End If
End Sub

'Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
'On Error Resume Next
'    If txtcodigo(indD).Text <> "" Then
'        cad = cad & "desde " & txtcodigo(indD).Text
'        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
'    End If
'    If txtcodigo(indH).Text <> "" Then
'        cad = cad & "  hasta " & txtcodigo(indH).Text
'        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
'    End If
'    AnyadirParametroDH = cad
'    If Err.Number <> 0 Then Err.Clear
'End Function

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
    
    Documento = ""
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



'Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
'Dim devuelve As String
'Dim cad As String
'
'    PonerDesdeHasta = False
'    devuelve = CadenaDesdeHasta(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
'    If devuelve = "Error" Then Exit Function
'    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
'
'    'para MySQL
'    If Tipo <> "F" Then
'        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
'    Else
'        'Fecha para la Base de Datos
'        cad = CadenaDesdeHastaBD(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
'        If Not AnyadirAFormula(cadSelect, cad) Then Exit Function
'    End If
'
'    If devuelve <> "" Then
'        If param <> "" Then
'            'Parametro Desde/Hasta
'            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
'            numParam = numParam + 1
'        End If
'        PonerDesdeHasta = True
'    End If
'End Function


Private Sub LlamarImprimir()
     With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = Titulo
        .NombreRPT = nomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub



'Private Sub EnviarSMS(cadWhere As String, cadTit As String, cadRpt As String, cadTabla As String, ByRef EstaOk As Boolean)
'Dim Sql As String
'Dim RS As ADODB.Recordset
'Dim Cad1 As String, cad2 As String, lista As String
'Dim cont As Integer
'Dim Direccion As String
'
'Dim NF As Integer
'Dim cad As String
'Dim b As Boolean
'
'On Error GoTo EEnviar
'
'
'    If vParamAplic.SMSclave = "" Or vParamAplic.SMSemail = "" Or vParamAplic.SMSremitente = "" Then
'        MsgBox "No tiene configurados los parámetros de Envio de SMS. Revise.", vbExclamation
'        Exit Sub
'    End If
'
'    Screen.MousePointer = vbHourglass
'
'    If cadTabla = "(rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio) inner join rcampos on rsocios.codsocio = rcampos.codsocio " Or _
'       cadTabla = "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio" Then
'        'seleccionamos todos los socios a los que queremos enviar un SMS
'        Sql = "SELECT distinct rsocios.codsocio,nomsocio,rsocios.movsocio "
'    End If
'    Sql = Sql & "FROM " & cadTabla
'    Sql = Sql & " WHERE " & cadWhere
'
'    Set RS = New ADODB.Recordset
'    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    cont = 0
'    lista = ""
'
'    b = True
'
'
'    While Not RS.EOF And b
'    'para cada socio enviamos un SMS
'        Cad1 = DBLet(RS.Fields(2), "T") 'movil socio
'
'        If Cad1 = "" Then 'no tiene movil
'            lista = lista & Format(RS.Fields(0), "000000") & " - " & RS.Fields(1) & vbCrLf
'            EstaOk = False
'        End If
'
'        If Cad1 <> "" Then 'HAY movil  --> ENVIAMOS el mensaje
'            Direccion = "http://www.afilnet.com/http/sms/?email=" & Trim(vParamAplic.SMSemail) & "&pass=" & Trim(vParamAplic.SMSclave)
'            Direccion = Direccion & "&mobile=" & Trim(Cad1) & "&id=" & Trim(vParamAplic.SMSremitente)
'            Direccion = Direccion & "&country=0034" & "&sms=" & txtcodigo(2).Text & "&now=" & Format(Check1.Value, "0")
'            Direccion = Direccion & "&date=" & Format(txtcodigo(0).Text, "yyyy/mm/dd") & " " & Format(txtcodigo(1).Text, "hh:mm")
'            Direccion = Direccion & "&type=" & Format(Check2.Value, "0")
'
'            Screen.MousePointer = vbHourglass
'
'            Label9(10).Caption = Format(RS.Fields(0), "000000") & " - " & RS.Fields(1) & " - " & RS.Fields(2)
'            DoEvents
'
'            'Cargamos en el fichero el resultado de enviar un mensaje
'            GetFileFromUrl Direccion, App.Path & "\RESULT.TXT"
'
'            NF = FreeFile
'            Open App.Path & "\RESULT.TXT" For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
'            cad = ""
'            Line Input #NF, cad
'            Close NF
'
'            Select Case Mid(cad, 1, 2)
'                Case "OK"
'                    espera 2
'
'                    Me.Refresh
'                    espera 0.4
'                    cont = cont + 1
'
'                    Sql = "INSERT INTO rsmsenviados (codsocio, movsocio, fechaenvio, horaenvio, texto)"
'                    Sql = Sql & " VALUES (" & DBSet(RS.Fields(0), "N") & "," & DBSet(Cad1, "T") & ","
'                    Sql = Sql & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(1).Text, "H") & "," & DBSet(txtcodigo(2).Text, "T") & ")"
'                    conn.Execute Sql
'
'
'                Case "-1"
'                    MsgBox "Error en el Login, usuario o clave incorrectas", vbExclamation
'                    EstaOk = False
'                Case Else
'                    If Mid(cad, 1, 12) = "Sin Creditos" Then
'                        MsgBox "No tiene créditos. Revise", vbExclamation
'                        b = False
'                    Else
'                        If MsgBox("Error en el envio de mensaje al socio " & DBLet(RS.Fields(0), "N") & ". ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then b = False
'                    End If
'                    EstaOk = False
'            End Select
'
'            Screen.MousePointer = vbDefault
'
'        End If
'        RS.MoveNext
'    Wend
'    Label9(10).Caption = ""
'    DoEvents
'
'    RS.Close
'    Set RS = Nothing
'
'    Screen.MousePointer = vbDefault
'
'    'Mostra mensaje con aquellos socios que no tienen móvil
'    If lista <> "" Then
'        lista = "Socios sin Móvil:" & vbCrLf & vbCrLf & lista
'        MsgBox lista, vbInformation
'    End If
'
'
'EEnviar:
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, "Enviando SMS", Err.Description
'    End If
'End Sub

'###############COPIADO

Private Sub GetFileFromUrl(ByRef url As String, ByRef file As String)
    Dim fileBytes() As Byte
    Dim fileNum As Integer
    
    On Error GoTo DownloadError
    DoEvents
    
    fileBytes() = Inet1.OpenURL(url, icByteArray)
    
    fileNum = FreeFile
    Open file For Binary Access Write As #fileNum
    Put #fileNum, , fileBytes()
    Close #fileNum
    
    Exit Sub
    
DownloadError:
    MsgBox Err.Description
End Sub

Private Sub InsertarTemporal(cadwhere As String, cadSelect As String)
Dim SQL As String
Dim RS As ADODB.Recordset
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

Private Sub EnviarEMailMulti(cadwhere As String, cadTit As String, cadRpt As String, cadTABLA As String)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer

On Error GoTo EEnviar

    Screen.MousePointer = vbHourglass
    
        'seleccionamos todos los socios a los que queremos enviar e-mail
    SQL = "SELECT distinct rsocios.codsocio,nomsocio,maisocio, maisocio "
    SQL = SQL & "FROM " & cadTABLA
    SQL = SQL & " WHERE " & cadwhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
    
    While Not RS.EOF
    'para cada cliente/proveedor enviamos un e-mail
        Cad1 = DBLet(RS.Fields(2), "T") 'e-mail administracion
        Cad2 = DBLet(RS.Fields(3), "T") 'e-mail compras
        
        If Cad1 = "" And Cad2 = "" Then 'no tiene e-mail
'              MsgBox "Sin mail para el proveedor: " & Format(RS!codProve, "000000") & " - " & RS!nomprove, vbExclamation
              lista = lista & Format(RS.Fields(0), "000000") & " - " & RS.Fields(1) & vbCrLf
        ElseIf Cad1 <> "" And Cad2 <> "" Then 'tiene 2 e-mail
            'ver a q e-mail se va a enviar (administracion, compras)
            If Me.OptMailCom(0).Value = True Then Cad1 = Cad2
        Else 'alguno de los 2 tiene valor
            If Cad2 <> "" Then Cad1 = Cad2  'e-mail para compras
        End If
        
        If Cad1 <> "" Then 'HAY email --> ENVIAMOS e-mail
            Label9(10).Caption = Format(RS.Fields(0), "000000") & " - " & RS.Fields(1) & " - " & RS.Fields(2)
            DoEvents


            If ImpresionNormal Then
                With frmImprimir
                    .OtrosParametros = cadParam
                    .NumeroParametros = numParam
                    '[Monica]05/09/2013: FALLO!!!! faltaba la condicion de tmpinformes.codusu
                    SQL = "{rsocios.codsocio}=" & RS.Fields(0) & " and {tmpinformes.codusu} = " & vUsu.Codigo
                    .Opcion = 306
                    .FormulaSeleccion = SQL
                    .EnvioEMail = True
                    CadenaDesdeOtroForm = "GENERANDO"
                    .Titulo = cadTit
                    .NombreRPT = cadRpt
                    .ConSubInforme = True
                    .Show vbModal
    
                    If CadenaDesdeOtroForm = "" Then
                    'si se ha generado el .pdf para enviar
                        SQL = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                        SQL = SQL & " VALUES (" & vUsu.Codigo & "," & DBSet(RS.Fields(0), "N") & "," & DBSet(RS.Fields(1), "T") & "," & DBSet(Cad1, "T") & ")"
                        conn.Execute SQL
                
                        Me.Refresh
                        espera 0.4
                        cont = cont + 1
                        'Se ha generado bien el documento
                        'Lo copiamos sobre app.path & \temp
                        SQL = RS.Fields(0) & ".pdf"
                        FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & SQL
                    End If
                End With
                Label9(10).Caption = ""
                DoEvents
            Else
                If CadenaDesdeOtroForm = "" Then
                'si se ha generado el .pdf para enviar
                    SQL = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                    SQL = SQL & " VALUES (" & vUsu.Codigo & "," & DBSet(RS.Fields(0), "N") & "," & DBSet(RS.Fields(1), "T") & "," & DBSet(Cad1, "T") & ")"
                    conn.Execute SQL
            
                    Me.Refresh
                    espera 0.4
                    cont = cont + 1
                    'Se ha generado bien el documento
                    'Lo copiamos sobre app.path & \temp
'                    Sql = Rs.Fields(0) & ".pdf"
'                    FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Sql
                End If
                Label9(10).Caption = ""
                DoEvents
            End If
        End If
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
      
    If cont > 0 Then
        espera 0.4
        SQL = "Carta: " & txtNombre(63).Text & "|"
            
        frmEMail.Opcion = 2
        frmEMail.DatosEnvio = SQL
        frmEMail.CodCryst = IndRptReport
        If Not ImpresionNormal Then
            frmEMail.Opcion = 5
            frmEMail.Ficheros = cadSelect
        Else
            frmEMail.Ficheros = ""
        End If
        frmEMail.Show vbModal

        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
        
        'Borrar la carpeta con temporales
        If ImpresionNormal Then Kill App.Path & "\temp\*.pdf"
    End If
    
    Screen.MousePointer = vbDefault
   
    'Mostra mensaje con aquellos proveedores que no tienen e-mail
    If lista <> "" Then
        lista = "Socios sin e-mail:" & vbCrLf & vbCrLf & lista
        MsgBox lista, vbInformation
    End If
    
EEnviar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Enviando Informe por e-mail", Err.Description
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
    End If
End Sub

Private Sub CargarIconos()
Dim I As Integer

    For I = 37 To 39
        Me.imgBuscarOfer(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 56 To 57
        Me.imgBuscarOfer(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I

    For I = 0 To imgAyuda.Count - 1
        imgAyuda(I).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next I

End Sub

Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim SQL As String

    b = True
    
    
    '[Monica]19/10/2012: comprobamos que si vamos a enviar mas de un documento es por email.
    If b And OpcionListado = 1 Then
        SQL = DevuelveDesdeBDNew(cAgro, "scartas", "descarta", "codcarta", vParamAplic.CartaPOZ, "N")
        If SQL = "" Then
            MsgBox "No tiene en parámetros la carta de Reclamación. Revise.", vbExclamation
            b = False
        End If

        If b Then
            Documento = DevuelveDesdeBDNew(cAgro, "scartas", "documrpt", "codcarta", txtcodigo(63).Text, "N")
            If Documento <> "" Then
                If InStr(1, Documento, ",") <> 0 And chkMail(0).Value = 0 Then
                    MsgBox "Para enviar más de un archivo adjunto debe seleccionar sólo por email.", vbExclamation
                    b = False
                    PonerFocoChk chkMail(0)
                Else
                    'cualquier otro tipo de documento se tiene que poder enviar por email
                    If InStr(1, Documento, ".rpt") = 0 And chkMail(0).Value = 0 Then
                        MsgBox "Para enviar más de un archivo adjunto debe seleccionar sólo por email.", vbExclamation
                        b = False
                        PonerFocoChk chkMail(0)
                    End If
                End If
                
                If b And InStr(1, Documento, ".rpt") = 0 Then
                    'si no es un rpt a ejecutar de la carpeta de informes comprobamos que exista la carpeta de cartas
                    If Dir(App.Path & "\cartas", vbDirectory) = "" Then
                        MsgBox "No existe el directorio de cartas donde se introducen los archivos a adjuntar. Revise.", vbExclamation
                        b = False
                    End If
                    If b And Dir(App.Path & "\cartas\*.*", vbArchive) = "" Then
                        MsgBox "No existen archivos en el directorio cartas a adjuntar. Revise.", vbExclamation
                        b = False
                    End If
                End If
            End If
        End If
    End If
    
    DatosOk = b
    
End Function


Private Function GeneracionEnvioMail(ByRef RS As ADODB.Recordset) As Boolean
Dim Tipo As Integer
Dim TipoRec As String ' tipo de factura a la que rectifica
Dim Sql5 As String
Dim EsComplemen As Byte

    On Error GoTo EGeneracionEnvioMail
    GeneracionEnvioMail = False
    
    cadSelect = "Select * from tmpinformes where codusu =" & vUsu.Codigo & " ORDER BY importe1,codigo1"
    RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CodClien = ""
    While Not RS.EOF
    
        InicializarVbles
        '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    
        Label14(22).Caption = "Factura: " & RS!importe1 & " " & RS!Nombre1
        Label14(22).Refresh
        
        Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
        Dim nomDocu As String 'Nombre de Informe rpt de crystal
        
        'Facturas socios
        Select Case Mid(RS!Nombre1, 1, 3)
            Case "FRS" ' Impresion de facturas rectificativas
                       ' hacemos caso del codtipom que rectifica
                  TipoRec = DevuelveValor("select rectif_codtipom from rfactsoc where numfactu = " & DBSet(RS!importe1, "N") & " and codtipom = " & DBSet(RS!Nombre1, "T") & " and fecfactu = " & DBSet(RS!fecha1, "T"))
                       
                  Select Case Mid(TipoRec, 1, 3)
                        Case "FLI"
                            indRPT = 38 'Impresion de Factura Socio de Industria
                        Case Else
                            Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Mid(RS!Nombre1, 1, 3), "T"))
                            If Tipo >= 7 And Tipo <= 10 Then
                                indRPT = 42 'Imporesion de Facturas de Bodega o Almazara
                            Else
                                indRPT = 23 'Impresion de Factura Socio
                                
'[Monica]07/02/2012: Hemos marcado las facturas que son complementarias, ya no hace falta esto
'
'                                '[Monica]07/02/2012: Si la factura que rectifico es complementaria le pasamos el parametro
'                                Sql5 = "select esliqcomplem from rfactsoc where (codtipom, numfactu, fecfactu) in (select rectific_codtipom, rectific_numfactu, rectific_fecfactu from rfactsoc where codtipom = " & DBSet(Rs!Nombre1, "T") & " and numfactu = " & DBSet(Rs!importe1, "N") & " and fecfactu = " & DBSet(Rs!Fecha1, "F") & ")"
'                                EsComplemen = DevuelveValor(Sql5)
'
'                                cadParam = cadParam & "pComplem=" & EsComplemen & "|"
'                                numParam = numParam + 1
                            End If
                  End Select
            
            Case "FLI"
                indRPT = 38 'Impresion de Factura Socio de Industria
            Case Else
                Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Mid(RS!Nombre1, 1, 3), "T"))
                If Tipo >= 7 And Tipo <= 10 Then
                    indRPT = 42 'Imporesion de Facturas de Bodega o Almazara
                Else
                    indRPT = 23 'Impresion de Factura Socio

'[Monica]07/02/2012: Hemos marcado las facturas que son complementarias, ya no hace falta esto
'
'                    'Si es complementaria le pasamos el parametro
'                    cadParam = cadParam & "pComplem=" & Rs!campo1 & "|"
'                    numParam = numParam + 1
                End If
       End Select
        
       If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Function
       'Nombre fichero .rpt a Imprimir
        
        
       cadFormula = "({rfactsoc.codtipom}='" & Trim(RS!Nombre1) & "') "
       cadFormula = cadFormula & " AND ({rfactsoc.numfactu}=" & RS!importe1 & ") "
       cadFormula = cadFormula & " AND ({rfactsoc.fecfactu}= Date(" & Year(RS!fecha1) & "," & Month(RS!fecha1) & "," & Day(RS!fecha1) & "))"

   
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = True
            .NombreRPT = nomDocu
            .ConSubInforme = True
            .Opcion = 0
            .Titulo = ""
            .Show vbModal
        End With
    
                    
        'Subo el progress bar
        Label14(22).Caption = "Generando PDF"
        Label14(22).Refresh
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        If (Me.ProgressBar1.Value Mod 25) = 24 Then
            Me.Refresh
            DoEvents
            espera 1
        End If
        Me.Refresh
        DoEvents
        
        
        FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!Nombre1 & Format(RS!importe1, "0000000") & ".pdf" 'RS!importe1 & Format(RS!Codigo1, "0000000") & ".pdf"
        
        RS.MoveNext
    Wend
    RS.Close
    
    Set RS = Nothing
    GeneracionEnvioMail = True
    Exit Function
EGeneracionEnvioMail:
       MuestraError Err.Number
End Function

Private Sub PonerTamnyosMail(peque As Boolean)
    If peque Then
        Me.Height = Me.FrameEnvioMail.Height + 60
        Me.Width = Me.FrameEnvioMail.Width
    Else
        Me.Height = Me.FrameEnvioFacMail.Height
        Me.Width = Me.FrameEnvioFacMail.Width
    End If
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 120
    Me.FrameEnvioMail.visible = peque
    Me.FrameEnvioFacMail.visible = Not peque
    DoEvents
    Me.Refresh
End Sub

Private Sub CargarCombo()
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo

'Lo cargamos con los valores de la tabla stipom que tengan tipo de documento=Albaranes (tipodocu=1)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim I As Byte

    On Error GoTo ECargaCombo

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
    
    
ECargaCombo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub CargaCombo()
Dim RS As ADODB.Recordset
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
        Combo1(1).ItemData(Combo1(1).NewIndex) = 5
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




Private Function SituacionesBloqueo() As String
Dim SQL As String
Dim cadena As String
Dim RS As ADODB.Recordset

    cadena = ""

    SituacionesBloqueo = cadena

    SQL = "select codsitua from rsituacion where bloqueo = 1"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        cadena = cadena & DBLet(RS!codsitua) & ","
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    SituacionesBloqueo = cadena

End Function

'
'    COMO ESTABA ANTES DE IMPRIMIR CARTAS Y ENVIAR POR EMAIL.
'
Private Sub cmdAceptarCartaRecANTES_Click()
'1: Listado para cartas de reclamacion a proveedor
Dim campo As String
Dim Tabla As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

Dim OK As Boolean

Dim Situacion As String
Dim Tipos As String
Dim CodTipom As String
Dim cDesde As String
Dim cHasta As String
Dim nDesde As String
Dim nHasta As String


    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    Tabla = "rrecibpozos"
    
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
    cDesde = Trim(txtcodigo(60).Text)
    cHasta = Trim(txtcodigo(61).Text)
    nDesde = txtNombre(60).Text
    nHasta = txtNombre(61).Text
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
    
    'si es listado de CARTAS/eMAIL a socios comprobar que se ha seleccionado
    'una carta para imprimir
    If txtcodigo(63).Text = "" Then
        MsgBox "Debe seleccionar una carta para imprimir.", vbInformation
        Exit Sub
    End If
    
    'Parametro cod. carta
    cadParam = "|pCodCarta= " & txtcodigo(63).Text & "|"
    numParam = numParam + 1
    
    'Parametro fecha
    cadParam = cadParam & "|pFecha= """ & txtcodigo(0).Text & """|"
    numParam = numParam + 1
    
    'Nombre fichero .rpt a Imprimir
    nomRPT = "rSocioCarta.rpt" '"rComProveCarta.rpt"
    Titulo = "Cartas Reclamación a Socios"
    
    indRPT = 61 'Personalizacion de la carta a socios
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    '[Monica]19/10/2012: nueva variable para indicar que se pasa por visreport o no ImpresionNormal
    ImpresionNormal = True
    Documento = DevuelveDesdeBDNew(cAgro, "scartas", "documrpt", "codcarta", txtcodigo(63).Text, "N")
    
    '[Monica]19/07/2013: dejo introducir una carta que hay creado el usuario
    If Documento <> "" Then
        nomDocu = Documento
    End If
      
    'Nombre fichero .rpt a Imprimir
    nomRPT = nomDocu
    
    conSubRPT = True
    
        
    '[Monica]11/11/2011: añadimos los socios que esten dados de baja en todas las secciones
    ' solo se sacan los socios que no esten dados de baja
    If Not AnyadirAFormula(cadSelect, "{rsocios.fechabaja} is null ") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "isnull({rsocios.fechabaja})") Then Exit Sub

    '[Monica]23/11/2012: si es escalona o utxera seleccionamos que tipo de socio
    If Option1(0).Value Then    ' solo contado
        If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
    End If
    If Option1(1).Value Then    ' solo efecto
        If Not AnyadirAFormula(cadSelect, "{rsocios.cuentaba<>""8888888888""") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}<>""8888888888""") Then Exit Sub
    End If
'   no hacemos nada
'    If Option1(2).Value Then
'    End If
    
    '[Monica]08/11/2012: solo los socios que no tengan situacion de bloqueo
    Situacion = SituacionesBloqueo
    If Situacion <> "" Then
        Situacion = Mid(Situacion, 1, Len(Situacion) - 1)
        If Not AnyadirAFormula(cadFormula, "not {rsocios.codsitua} in [" & Situacion & "]") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "not rsocios.codsitua in (" & Situacion & ")") Then Exit Sub
    End If
    
    ' si es un correo electronico miramos solo los que tienen mail
    If chkMail(0).Value = 1 Then
        If Not AnyadirAFormula(cadSelect, "not {rsocios.maisocio} is null and {rsocios.maisocio}<>''") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "not isnull({rsocios.maisocio}) and {rsocios.maisocio}<>''") Then Exit Sub
    End If
    ' hasta aqui
    
    '=========================================================================================
    '[Monica]20/01/2012: Añadida la condicion de marca de correo del socio
    '                    Sólo se tendrá en cuenta en cartas y etiquetas. No en los sms
    '=========================================================================================
    If chkMail(0).Value = 0 Then
        If chkMail(2).Value = 1 Then
            If Not AnyadirAFormula(cadSelect, "{rsocios.correo} = 1") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.correo} = 1") Then Exit Sub
        End If
    End If
    
    Tabla = "rsocios inner join rrecibpozos on rsocios.codsocio = rrecibpozos.codsocio"
    
    'ver si hay registros seleccionados para mostrar en el informe
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    Set frmMen = New frmMensajes
    frmMen.cadwhere = cadSelect
    frmMen.OpcionMensaje = 42 'Socios
    frmMen.Show vbModal
    Set frmMen = Nothing
    If cadSelect = "" Then Exit Sub
    
    If OpcionListado = 1 And Me.chkMail(0).Value = 1 Then
        'Enviarlo por e-mail
        IndRptReport = indRPT
        EnviarEMailMulti cadSelect, Titulo, nomDocu, Tabla ' "rSocioCarta.rpt", Tabla  'email para socios
        cmdCancel_Click (9)
    Else
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        LlamarImprimir
    End If
End Sub



