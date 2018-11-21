VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmListadoOfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   10305
   Icon            =   "frmListadoOfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTipoSocio 
      Caption         =   "Tipo Socio"
      Height          =   615
      Left            =   7080
      TabIndex        =   68
      Top             =   6120
      Width           =   3705
      Begin VB.OptionButton Option1 
         Caption         =   "Contado"
         Height          =   225
         Index           =   0
         Left            =   270
         TabIndex        =   71
         Top             =   240
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Efecto"
         Height          =   225
         Index           =   1
         Left            =   1515
         TabIndex        =   70
         Top             =   240
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   225
         Index           =   2
         Left            =   2760
         TabIndex        =   69
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   360
         TabIndex        =   31
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
         TabIndex        =   32
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
   Begin VB.Frame FrameEtiqProv 
      Height          =   7155
      Left            =   0
      TabIndex        =   14
      Top             =   30
      Width           =   7035
      Begin VB.CheckBox chkMail 
         Caption         =   "Sin Fec.Revisión"
         Height          =   345
         Index           =   4
         Left            =   3030
         TabIndex        =   79
         Top             =   2850
         Width           =   1515
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   480
         TabIndex        =   75
         Top             =   750
         Width           =   6105
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Tag             =   "Tipo|N|N|||straba|codsecci||N|"
            Top             =   60
            Width           =   1005
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   2
            Left            =   2400
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   90
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fase"
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
            Left            =   120
            TabIndex        =   76
            Top             =   30
            Width           =   345
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   4980
         TabIndex        =   73
         Top             =   2880
         Width           =   2025
         Begin VB.Label Label1 
            Caption         =   "Adjuntar Archivos"
            Height          =   195
            Left            =   60
            TabIndex        =   74
            Top             =   30
            Width           =   1275
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   1
            Left            =   1350
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.CheckBox chkMail 
         Caption         =   "Sólo con marca de Correo"
         Height          =   345
         Index           =   2
         Left            =   570
         TabIndex        =   38
         Top             =   2850
         Width           =   2145
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   5760
         Top             =   5580
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Frame Frame4 
         Caption         =   "SMS"
         Height          =   1965
         Left            =   510
         TabIndex        =   34
         Top             =   4560
         Width           =   6045
         Begin VB.CheckBox Check3 
            Caption         =   "Sólo sin eMail"
            Height          =   345
            Left            =   1980
            TabIndex        =   37
            Top             =   270
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox txtCodigo 
            Height          =   945
            Index           =   2
            Left            =   210
            MaxLength       =   150
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   870
            Width           =   5505
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Con Remitente"
            Height          =   345
            Left            =   3540
            TabIndex        =   9
            Top             =   600
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   4710
            MaxLength       =   10
            TabIndex        =   10
            Top             =   300
            Width           =   1005
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Enviar Ahora"
            Height          =   345
            Left            =   240
            TabIndex        =   8
            Top             =   270
            Width           =   1935
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Texto"
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
            Left            =   240
            TabIndex        =   36
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hora"
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
            Left            =   4140
            TabIndex        =   35
            Top             =   345
            Width           =   345
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   62
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2520
         Width           =   4515
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5490
         TabIndex        =   13
         Top             =   6570
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarEtiqProv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4410
         TabIndex        =   12
         Top             =   6570
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   3225
         Left            =   360
         TabIndex        =   24
         Top             =   3270
         Width           =   6255
         Begin VB.CheckBox chkMail 
            Caption         =   "Imprimir los que no tienen e-mail"
            Height          =   285
            Index           =   5
            Left            =   3720
            TabIndex        =   80
            Top             =   990
            Width           =   2685
         End
         Begin VB.CheckBox chkMail 
            Caption         =   "Enviar SMS"
            Height          =   315
            Index           =   1
            Left            =   3720
            TabIndex        =   7
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chkMail 
            Caption         =   "Enviar por e-mail"
            Height          =   345
            Index           =   0
            Left            =   3720
            TabIndex        =   6
            Top             =   450
            Width           =   1935
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   5
            Top             =   510
            Width           =   1005
         End
         Begin VB.Frame Frame3 
            Caption         =   "e-Mail"
            Enabled         =   0   'False
            Height          =   780
            Left            =   180
            TabIndex        =   27
            Top             =   1350
            Visible         =   0   'False
            Width           =   1695
            Begin VB.OptionButton OptMailAdm 
               Caption         =   "Administración"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   29
               Top             =   210
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton OptMailCom 
               Caption         =   "Compras"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   28
               Top             =   460
               Width           =   1335
            End
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   63
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "Text5"
            Top             =   150
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   63
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   4
            Top             =   165
            Width           =   1005
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1080
            Picture         =   "frmListadoOfer.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   510
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
            TabIndex        =   33
            Top             =   525
            Width           =   435
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   39
            Left            =   1080
            ToolTipText     =   "Buscar carta"
            Top             =   165
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
            TabIndex        =   26
            Top             =   180
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
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   1725
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1725
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   61
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   2070
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   2
         Top             =   2070
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1260
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   58
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   1260
         Width           =   3735
      End
      Begin VB.Image imgAyuda 
         Height          =   255
         Index           =   3
         Left            =   4560
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   2880
         Width           =   255
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
         Left            =   570
         TabIndex        =   67
         Top             =   6390
         Width           =   3705
      End
      Begin VB.Image imgAyuda 
         Height          =   255
         Index           =   0
         Left            =   2760
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "A la atención de:"
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
         Index           =   8
         Left            =   600
         TabIndex        =   18
         Top             =   2520
         Width           =   1215
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
         Left            =   600
         TabIndex        =   23
         Top             =   1485
         Width           =   375
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   37
         Left            =   1440
         ToolTipText     =   "Buscar socio"
         Top             =   1725
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
         Left            =   960
         TabIndex        =   22
         Top             =   1725
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   38
         Left            =   1440
         ToolTipText     =   "Buscar socio"
         Top             =   2070
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
         Left            =   960
         TabIndex        =   21
         Top             =   2070
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Sección"
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
         TabIndex        =   17
         Top             =   1155
         Width           =   540
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   35
         Left            =   1440
         ToolTipText     =   "Buscar sección"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Etiquetas Socios"
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
         Left            =   600
         TabIndex        =   16
         Top             =   270
         Width           =   2460
      End
   End
   Begin VB.Frame FrameEnvioDatosComunica 
      Height          =   6420
      Left            =   90
      TabIndex        =   81
      Top             =   180
      Width           =   6345
      Begin VB.Frame FrameExportar 
         BorderStyle     =   0  'None
         Height          =   3075
         Left            =   225
         TabIndex        =   86
         Top             =   1935
         Width           =   5820
         Begin VB.Frame FrameFecEntradas 
            Enabled         =   0   'False
            Height          =   1320
            Left            =   2700
            TabIndex        =   93
            Top             =   180
            Width           =   2985
            Begin VB.TextBox txtCodigo 
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
               Left            =   1425
               MaxLength       =   10
               TabIndex        =   95
               Top             =   900
               Width           =   1350
            End
            Begin VB.TextBox txtCodigo 
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
               Left            =   1440
               MaxLength       =   10
               TabIndex        =   94
               Top             =   435
               Width           =   1350
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   5
               Left            =   525
               TabIndex        =   100
               Top             =   900
               Width           =   570
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   2
               Left            =   1155
               Picture         =   "frmListadoOfer.frx":0097
               Top             =   900
               Width           =   240
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   1
               Left            =   1155
               Picture         =   "frmListadoOfer.frx":0122
               Top             =   465
               Width           =   240
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Entrada"
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
               Index           =   0
               Left            =   135
               TabIndex        =   98
               Top             =   135
               Width           =   1440
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   4
               Left            =   525
               TabIndex        =   96
               Top             =   480
               Width           =   600
            End
         End
         Begin VB.Frame FrameFecAlbaranes 
            Enabled         =   0   'False
            Height          =   1365
            Left            =   2700
            TabIndex        =   89
            Top             =   1665
            Width           =   2985
            Begin VB.TextBox txtCodigo 
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
               Left            =   1425
               MaxLength       =   10
               TabIndex        =   97
               Top             =   450
               Width           =   1350
            End
            Begin VB.TextBox txtCodigo 
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
               Left            =   1425
               MaxLength       =   10
               TabIndex        =   99
               Top             =   885
               Width           =   1350
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   0
               Left            =   480
               TabIndex        =   92
               Top             =   900
               Width           =   570
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   3
               Left            =   1170
               Picture         =   "frmListadoOfer.frx":01AD
               Top             =   450
               Width           =   240
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   4
               Left            =   1170
               Picture         =   "frmListadoOfer.frx":0238
               Top             =   915
               Width           =   240
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Albarán"
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
               Left            =   135
               TabIndex        =   91
               Top             =   135
               Width           =   1410
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   1
               Left            =   480
               TabIndex        =   90
               Top             =   450
               Width           =   600
            End
         End
         Begin VB.CheckBox ChkAlbaranes 
            Caption         =   "Albaranes de Venta"
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
            Left            =   180
            TabIndex        =   88
            Top             =   1800
            Width           =   2625
         End
         Begin VB.CheckBox ChkEntradas 
            Caption         =   "Entradas Clasificadas"
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
            Left            =   180
            TabIndex        =   87
            Top             =   315
            Width           =   2580
         End
      End
      Begin VB.Frame Frame6 
         Height          =   960
         Left            =   225
         TabIndex        =   83
         Top             =   945
         Width           =   5730
         Begin VB.OptionButton Option2 
            Caption         =   "Importar desde csv"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   2880
            TabIndex        =   85
            Top             =   360
            Width           =   2355
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Exportar a csv"
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
            Left            =   585
            TabIndex        =   84
            Top             =   315
            Width           =   1860
         End
      End
      Begin VB.CommandButton CmdAceptarComunica 
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
         Left            =   3690
         TabIndex        =   101
         Top             =   5865
         Width           =   1065
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
         Left            =   4860
         TabIndex        =   103
         Top             =   5865
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   240
         Left            =   315
         TabIndex        =   104
         Top             =   5040
         Visible         =   0   'False
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   105
         Top             =   5535
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label lblProgres 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   102
         Top             =   5310
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label14 
         Caption         =   "Comunicación de Datos comunicación"
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
         Index           =   3
         Left            =   240
         TabIndex        =   82
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame FrameEnvioFacMail 
      Height          =   6015
      Left            =   45
      TabIndex        =   39
      Top             =   45
      Width           =   10215
      Begin VB.CheckBox Check5 
         Caption         =   "Impresión con Arrobas"
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
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   5130
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Incluir los ya traspasados"
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
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   5370
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   4065
         MaxLength       =   7
         TabIndex        =   47
         Tag             =   "Nº Factura|N|S|||rfactsoc|numfactu|0000000|S|"
         Top             =   3480
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1545
         MaxLength       =   7
         TabIndex        =   46
         Tag             =   "Nº Factura|N|S|||rfactsoc|numfactu|0000000|S|"
         Text            =   "wwwwwww"
         Top             =   3480
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1545
         MaxLength       =   10
         TabIndex        =   44
         Top             =   2640
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   4065
         MaxLength       =   10
         TabIndex        =   45
         Top             =   2640
         Width           =   1350
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
         Index           =   18
         Left            =   8865
         TabIndex        =   55
         Top             =   5370
         Width           =   1065
      End
      Begin VB.CheckBox chkMail 
         Caption         =   "Copia remitente"
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
         Index           =   3
         Left            =   5655
         TabIndex        =   49
         Top             =   1830
         Width           =   2460
      End
      Begin VB.TextBox Text1 
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
         Left            =   5640
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   2625
         Width           =   4335
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
         Index           =   110
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "Text5"
         Top             =   1185
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   41
         Top             =   1185
         Width           =   855
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
         Index           =   111
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Text5"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   43
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Index           =   1
         Left            =   5640
         MultiLine       =   -1  'True
         TabIndex        =   51
         Text            =   "frmListadoOfer.frx":02C3
         Top             =   3480
         Width           =   4335
      End
      Begin VB.CommandButton cmdEnvioMail 
         Caption         =   "&Enviar"
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
         Left            =   7695
         TabIndex        =   53
         Top             =   5370
         Width           =   1065
      End
      Begin VB.ListBox ListTipoMov 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Index           =   1000
         ItemData        =   "frmListadoOfer.frx":02C9
         Left            =   1545
         List            =   "frmListadoOfer.frx":02CB
         Style           =   1  'Checkbox
         TabIndex        =   48
         Top             =   4215
         Width           =   3825
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   4680
         Picture         =   "frmListadoOfer.frx":02CD
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   5040
         Picture         =   "frmListadoOfer.frx":0417
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   13
         Left            =   3360
         TabIndex        =   66
         Top             =   3510
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   14
         Left            =   600
         TabIndex        =   65
         Top             =   3510
         Width           =   600
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         Index           =   15
         Left            =   240
         TabIndex        =   64
         Top             =   3225
         Width           =   1080
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
         TabIndex        =   63
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   17
         Left            =   600
         TabIndex        =   62
         Top             =   2685
         Width           =   600
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   61
         Top             =   2340
         Width           =   1440
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   33
         Left            =   1260
         Picture         =   "frmListadoOfer.frx":0561
         Top             =   2670
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   34
         Left            =   3780
         Picture         =   "frmListadoOfer.frx":05EC
         Top             =   2670
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   18
         Left            =   3120
         TabIndex        =   60
         Top             =   2685
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Asunto"
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
         Left            =   5640
         TabIndex        =   59
         Top             =   2340
         Width           =   690
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   56
         Left            =   1260
         Top             =   1185
         Width           =   240
      End
      Begin VB.Label Label9 
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
         Index           =   32
         Left            =   240
         TabIndex        =   58
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   33
         Left            =   600
         TabIndex        =   57
         Top             =   1185
         Width           =   600
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   57
         Left            =   1260
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   34
         Left            =   600
         TabIndex        =   56
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   20
         Left            =   240
         TabIndex        =   54
         Top             =   4155
         Width           =   1260
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje"
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
         Left            =   5640
         TabIndex        =   52
         Top             =   3180
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmListadoOfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Opcionlistado As Integer
    '(ver opciones en frmListado)
        
        
        
    '315:  Envio por mail de las facturas
        
        
    '320: envio de datos a cooperativa (comunicacion de datos entre coopic y agrocitrica)
    
    
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
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
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

Dim SqlDeta As String
Dim ConDetalle As Boolean

Dim NotasParaGastos As String

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check5_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub ChkAlbaranes_Click()
    Me.FrameFecAlbaranes.Enabled = (ChkAlbaranes.Value = 1)
    txtCodigo(3).Text = ""
    txtCodigo(4).Text = ""
End Sub

Private Sub ChkEntradas_Click()
    Me.FrameFecEntradas.Enabled = (ChkEntradas.Value = 1)
    txtCodigo(5).Text = ""
    txtCodigo(6).Text = ""
End Sub

Private Sub chkmail_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub


Private Sub CmdAceptarComunica_Click()
Dim v_cadena As String
Dim FrmEM As frmEMail
Dim cadTabla As String
Dim Sql As String
Dim cadTitulo As String
Dim cadNombreRPT As String
Dim b As Boolean

    If Option2(0) Then
        If Me.ChkEntradas.Value Then
            If txtCodigo(5).Text = "" Or txtCodigo(6).Text = "" Then
                MsgBox "Debe introducir obligatoriamente el rango de fechas.", vbExclamation
                PonerFoco txtCodigo(5)
                Exit Sub
            End If
            If CDate(txtCodigo(5).Text) > CDate(txtCodigo(6).Text) Then
                MsgBox "Desde no puede ser mayor que hasta", vbExclamation
                PonerFoco txtCodigo(5)
                Exit Sub
            End If
        End If
        
        If ChkAlbaranes.Value Then
            If txtCodigo(3).Text = "" Or txtCodigo(4).Text = "" Then
                MsgBox "Debe introducir obligatoriamente el rango de fechas.", vbExclamation
                PonerFoco txtCodigo(3)
                Exit Sub
            End If
            If CDate(txtCodigo(3).Text) > CDate(txtCodigo(4).Text) Then
                MsgBox "Desde no puede ser mayor que hasta", vbExclamation
                PonerFoco txtCodigo(5)
                Exit Sub
            End If
        End If
  
        If CargarFicheroCsv(txtCodigo(5), txtCodigo(6), txtCodigo(3), txtCodigo(4), ChkEntradas.Value, ChkAlbaranes.Value, Me.cd1) Then
        End If
    Else
        Me.cd1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    
        Me.cd1.DefaultExt = "csv"
        cd1.Filter = "Archivos CSV|*.csv|"
        cd1.FilterIndex = 1
        Me.cd1.FileName = "comunica.csv"
        Me.cd1.CancelError = True
        Me.cd1.ShowOpen
        
        
        If Me.cd1.FileName <> "" Then
            InicializarVbles
            InicializarTabla
                '========= PARAMETROS  =============================
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1

            'Si estamos importando y hay entradas modificadas preguntamos si quieren seguir
            If HayEntradasModificadas(Me.cd1.FileName) Then
                cadTabla = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo

                Sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo

                If TotalRegistros(Sql) <> 0 Then
                    Titulo = "Entradas comunicadas modificadas"
                    cadNombreRPT = "rErroresTrasDatosCoop.rpt"
                    
                    nomRPT = cadNombreRPT
                    
                    LlamarImprimir
                    If MsgBox("Hay Entradas modificadas. ¿Desea continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
                End If
            End If
                
            '[Monica]31/10/2018:
            NotasParaGastos = ""
            
            conn.BeginTrans
            If ProcesarFicheroComunicacion2(Me.cd1.FileName) Then
                cadTabla = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo

                Sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo

                If TotalRegistros(Sql) <> 0 Then
                    conn.RollbackTrans
                
                    MsgBox "Hay errores en el Traspaso de Datos. Debe corregirlos previamente.", vbExclamation
                    Titulo = "Errores de Traspaso Datos"
                    cadNombreRPT = "rErroresTrasDatosCoop.rpt"
                    
                    nomRPT = cadNombreRPT
                    
                    LlamarImprimir
                    Exit Sub
                Else
                    conn.CommitTrans
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                End If
            Else
                conn.RollbackTrans
            End If
        Else
            MsgBox "No ha seleccionado ningún fichero", vbExclamation
            Exit Sub
        End If
    End If
        
End Sub



Private Sub cmdEnvioMail_Click()
Dim Rs As ADODB.Recordset

Dim T1 As Single

    'El proceso constara de varias fases.
    'Fase 1: Montar el select y ver si hay registros
    'Fase 2: Preparar carpetas para los pdf
    'Fase 3: Generar para cada factura (una a una) del select su pdf
    'Fase 4: Enviar por mail, adjuntando los archivos correspondientes
    If Opcionlistado = 315 Then
        If Text1(0).Text = "" Then
            MsgBox "Ponga el asunto", vbExclamation
            Exit Sub
        End If
    Else
        Codigo = ""
        If vParamAplic.PathFacturaE = "" Then
            Codigo = "Falta configurar parametros"
        Else
'            MsgBox vParamAplic.PathFacturaE, vbExclamation
            If Dir(vParamAplic.PathFacturaE & "\", vbDirectory) = "" Then Codigo = "No existe carpeta"
'            MsgBox "todo ok", vbExclamation
        End If
        If Codigo <> "" Then
            MsgBox Codigo, vbExclamation
            Exit Sub
        End If
    End If
    
    'AHora pongo los tipo de facturas
    cadFormula = ""
    cadselect = ""  'ME dira si estan todas o no
    For indCodigo = 0 To Me.ListTipoMov(1000).ListCount - 1
        If Me.ListTipoMov(1000).Selected(indCodigo) Then
            'Esta checkeado
            cadFormula = cadFormula & " OR rfactsoc.codtipom = '" & Trim(Mid(ListTipoMov(1000).List(indCodigo), 1, 3)) & "'"
        Else
            cadselect = "NO"
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
    cadselect = ""
    
'    'Cadena para seleccion D/H Letra Serie
'    '--------------------------------------------
'    If txtCodigo(0).Text <> "" Or txtCodigo(1).Text <> "" Then
'        Codigo = "facturas.codtipom"
'        'Parametro Desde/Hasta Letra Serie
'        If Not PonerDesdeHasta(Codigo, "T", 0, 1, "") Then Exit Sub
'    End If
        
    'Cadena para seleccion D/H Factura
    '--------------------------------------------
    If txtCodigo(106).Text <> "" Or txtCodigo(107).Text <> "" Then
        Codigo = "rfactsoc.numfactu"
        If Not PonerDesdeHasta(Codigo, "N", 106, 107, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Fecha
    '--------------------------------------------
    If txtCodigo(108).Text <> "" Or txtCodigo(109).Text <> "" Then
        Codigo = "rfactsoc.fecfactu"
        If Not PonerDesdeHasta(Codigo, "F", 108, 109, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Socio
    '--------------------------------------------
    If txtCodigo(110).Text <> "" Or txtCodigo(111).Text <> "" Then
        Codigo = "rfactsoc.codsocio"
        If Not PonerDesdeHasta(Codigo, "N", 110, 111, "") Then Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Eliminamos temporales
    conn.Execute "DELETE from tmpinformes where codusu =" & vUsu.Codigo
    
    If cadselect <> "" Then cadselect = cadselect & " AND "
    cadselect = cadselect & NomTabla
    cadselect = " WHERE " & cadselect

    
    Set Rs = New ADODB.Recordset
    DoEvents
        
    If Opcionlistado = 316 Then
        If Me.Check4.Value = 0 Then
            If cadselect <> "" Then cadselect = cadselect & " AND "
            cadselect = cadselect & " (rfactsoc.enfacturae = 0 )"
        End If
    End If
        
    'Ahora insertare en la tabla temporal tmpinformes las facturas que voy a generar pdf
    '                                         codsocio,numfactu,codtipom,fecfactu,totalfac,esliqcomplem
    Codigo = "insert into tmpinformes (codusu,codigo1,importe1, nombre1, fecha1, importe2,campo1) "
    Codigo = Codigo & " values ( " & vUsu.Codigo & ","
    
    If Not PrepararCarpetasEnvioMail Then Exit Sub
        
    Screen.MousePointer = vbHourglass

    'Vamos a meter todas las facturas en la tabla temporal para comprobar si tienen mail
    'los clientes
    
    NomTabla = "Select codtipom,numfactu,codsocio,fecfactu,totalfac,esliqcomplem from rfactsoc  " & cadselect
    'El orden vamos a hacerlo por: Tipo documento
    NomTabla = NomTabla & " ORDER BY codtipom, numfactu, fecfactu "
    Rs.Open NomTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not Rs.EOF
        NomTabla = Rs!Codsocio & "," & Rs!numfactu & ",'" & Trim(Rs!CodTipom) & "','" & Format(Rs!fecfactu, FormatoFecha)
        NomTabla = NomTabla & "'," & TransformaComasPuntos(CStr(DBLet(Rs!TotalFac, "N"))) & "," & DBLet(Rs!Esliqcomplem, "N") & ")"
        conn.Execute Codigo & NomTabla
        NumRegElim = NumRegElim + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    If NumRegElim = 0 Then
        If Opcionlistado = 316 Then
            MsgBox "Ningúna factura para traspasar a FacturaE", vbExclamation
        Else
            MsgBox "Ningun dato a enviar por mail", vbExclamation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Numero de registros
    NomTabla = NumRegElim
    
    
    If Opcionlistado = 315 Then
        
        'AHora ya tengo todos los datos de las facturas que voy  a imprimir
        
        cadselect = "Select codsocio,maisocio "
        cadselect = cadselect & " as email from tmpinformes,rsocios where codusu = " & vUsu.Codigo & " and codsocio=codigo1"
        cadselect = cadselect & " and (maisocio is null or maisocio = '') "
        cadselect = cadselect & " group by codsocio "
        Rs.Open cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        While Not Rs.EOF
            NumRegElim = NumRegElim + 1
            Rs.MoveNext
        Wend
        Rs.Close
        
        If NumRegElim > 0 Then
            If MsgBox("Tiene socio sin mail. Continuar sin sus datos?", vbQuestion + vbYesNo) = vbNo Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
                
            'Si no salimos borramos
            Rs.Open cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cadselect = "DELETE from tmpinformes where codusu =" & vUsu.Codigo & " and codigo1 ="
            While Not Rs.EOF
                conn.Execute cadselect & Rs!Codsocio
                Rs.MoveNext
            Wend
            Rs.Close
            
            
            cadselect = "Select count(*) from tmpinformes where codusu =" & vUsu.Codigo
            Rs.Open cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            If Not Rs.EOF Then
                If Not IsNull(Rs.Fields(0)) Then NumRegElim = DBLet(Rs.Fields(0), "N")
                
            End If
            Rs.Close
            
            If NumRegElim = 0 Then
                'NO hay datos para enviar
                
                Screen.MousePointer = vbDefault
                MsgBox "No hay datos para enviar por mail", vbExclamation
                Exit Sub
            Else
                cadselect = "Hay " & NumRegElim & " facturas para enviar por mail." & vbCrLf & "¿Continuar?"
                If MsgBox(cadselect, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
            End If
            If NumRegElim = 0 Then
                Set Rs = Nothing
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            '[Monica]17/02/2016: si hay alguna factura de anticipo preguntamos si lo detalla por campos o no
            ConDetalle = False
            If vParamAplic.Cooperativa = 4 Then
                SqlDeta = "select count(*) from tmpinformes where codusu =" & vUsu.Codigo & " and nombre1 = 'FAA'"
                If TotalRegistros(SqlDeta) <> 0 Then
                    If MsgBox("¿ Desea impresión detallada por campos ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        ConDetalle = True
                    Else
                        ConDetalle = False
                    End If
                    numParam = numParam + 1
                End If
            End If
                
            NomTabla = NumRegElim
        
        End If
        
    Else
        cadselect = "Hay " & NumRegElim & " facturas para integrar con facturaE." & vbCrLf & "¿Continuar?"
        If MsgBox(cadselect, vbQuestion + vbYesNo) = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        '[Monica]17/02/2016: si hay alguna factura de anticipo preguntamos si lo detalla por campos o no
        ConDetalle = False
        If vParamAplic.Cooperativa = 4 Then
            SqlDeta = "select count(*) from tmpinformes where codusu =" & vUsu.Codigo & " and nombre1 = 'FAA'"
            If TotalRegistros(SqlDeta) <> 0 Then
                If MsgBox("¿ Desea impresión detallada por campos ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    ConDetalle = True
                Else
                    ConDetalle = False
                End If
                numParam = numParam + 1
            End If
        End If
    End If
        
        
        
    PonerTamnyosMail True
    MDIppal.visible = False
    'Voy arriesgar.
    'Confio en que no envien por mail mas de 32000 facturas (un integer)
    Label14(22).Caption = "Preparando datos"
    Me.ProgressBar1.Max = CInt(NomTabla)
    Me.ProgressBar1.Value = 0
    
    
    
    NumRegElim = 0
    
    If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then T1 = Timer
    
    If GeneracionEnvioMail(Rs) Then NumRegElim = 1
    
    Label14(22).Caption = "Preparando envia email"
    Label14(22).Refresh
    
    'Si ha ido todo bien entonces numregelim=1
    If NumRegElim = 1 Then
        If Opcionlistado = 315 Then
            cadselect = "Select nomsocio, maisocio"
            cadselect = cadselect & " as email,tmpinformes.* from tmpinformes,rsocios where codusu = " & vUsu.Codigo & " and codsocio=codigo1"
    '        cadSelect = cadSelect & " group by codclien having email is null"
    
            '[Monica]31/01/2014: esperamos en catadau, antes de abrir la ventana del frmEmail
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then

                T1 = Timer - T1
                If T1 < 3 Then
                    T1 = 3 - T1
                    espera T1
                End If
                T1 = Timer
            End If
    
            
            frmEMail.DatosEnvio = Text1(0).Text & "|" & Text1(1).Text & "|" & Abs(0) & "|" & cadselect & "|"
            frmEMail.Opcion = 4 'Multienvio de facturacion
            frmEMail.Show vbModal
            
            
            'Para tranquilizar las pantallas, borrar los ficheros generados
            'Confio en que no envien por mail mas de 32000 facturas (un integer)
            Label14(22).Caption = "Restaurando ...."
            Me.ProgressBar1.visible = False
        Else
            'Copiar a parametros
            '
            MsgBox "Proceso finalizado", vbExclamation
                
        End If
    
        Me.Refresh
        DoEvents
        espera 1
    '[Monica]07/03/2013: solo borramos si volvemos a ejecutar el envio (como en gasolinera)
    '    PrepararCarpetasEnvioMail
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


Private Sub Command1_Click()

End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "El check de 'Solo con Marca de Correo' indica que se generará el " & vbCrLf & _
                      "documento únicamente a los socios que tengan marcada la casilla " & vbCrLf & _
                      "'Correo' de su ficha. " & vbCrLf
                      
            If Opcionlistado = 306 Then
                      vCadena = vCadena & vbCrLf & "No tendrá ningún efecto en el caso de que estemos mandando un sms. " & vbCrLf
            End If
            
            vCadena = vCadena & "" & vbCrLf & vbCrLf
                      
        Case 1
            vCadena = "Para adjuntar archivos debe seleccionar la carta 12. " & vbCrLf & vbCrLf & _
                      "En la descripción de la carta debe poner el Asunto. " & vbCrLf & _
                      vbCrLf
                      
            If Opcionlistado = 306 Then
                      vCadena = vCadena & "Si se envia a través de SMS no adjuntará ningún archivo. " & vbCrLf & vbCrLf
                      vCadena = vCadena & "Es aconsejable enviar en formato PDF."
            End If
            
            vCadena = vCadena & "" & vbCrLf & vbCrLf
        
        Case 2
            vCadena = "Si se indica sección, seleccionaremos los socios dados de alta en esa sección. " & vbCrLf & vbCrLf & _
                      "Si indica fase, no hay que poner nada en sección, para seleccionar los" & vbCrLf & _
                      "socios que estén en la fase indicada." & vbCrLf & _
                      vbCrLf
                      
            If Opcionlistado = 306 Then
                      vCadena = vCadena & "Si se envia a través de SMS no adjuntará ningún archivo. " & vbCrLf & vbCrLf
                      vCadena = vCadena & "Es aconsejable enviar en formato PDF."
            End If
            
            vCadena = vCadena & "" & vbCrLf & vbCrLf
            
        '[Monica]29/06/2014: check de sin fecha de revision
        Case 3
           ' "____________________________________________________________"
            vCadena = "El check de 'Sin Fecha de Revisión' indica que se generará el " & vbCrLf & _
                      "documento únicamente a los socios que no tengan Fecha de Revisión" & vbCrLf & _
                      "en  su ficha. " & vbCrLf
            
            vCadena = vCadena & "" & vbCrLf & vbCrLf
        
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub


Private Sub cmdAceptarEtiqProv_Click()
'305: Listado para etiquetas de proveedor
'306: Listado para cartas a proveedor
Dim campo As String
Dim tabla As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

Dim OK As Boolean

Dim Situacion As String
    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    'si es listado de CARTAS/eMAIL a socios comprobar que se ha seleccionado
    'una carta para imprimir
    If Opcionlistado = 306 Then
        If chkMail(1).Value = 1 Then
            ' si estamos mandando un SMS no es obligado meter un codigo de carta
        Else
            If txtCodigo(63).Text = "" Then
                MsgBox "Debe seleccionar una carta para imprimir.", vbInformation
                Exit Sub
            End If
        End If
        
        'Parametro cod. carta
        cadParam = "|pCodCarta= " & txtCodigo(63).Text & "|"
        numParam = numParam + 1
        
        'Parametro fecha
        cadParam = cadParam & "|pFecha= """ & txtCodigo(0).Text & """|"
        numParam = numParam + 1
        
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rSocioCarta.rpt" '"rComProveCarta.rpt"
        Titulo = "Cartas a Socios" '"Cartas a Proveedores"
        
        indRPT = 61 'Personalizacion de la carta a socios
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
          
        '[Monica]19/10/2012: nueva variable para indicar que se pasa por visreport o no ImpresionNormal
        ImpresionNormal = True
        Documento = DevuelveDesdeBDNew(cAgro, "scartas", "documrpt", "codcarta", txtCodigo(63).Text, "N")
        If Documento <> "" Then
            '[Monica]19/10/2012: Si hay mas de un documento a imprimir o .rpt
            If InStr(1, Documento, ",") <> 0 Then
                ImpresionNormal = False
            Else
                If InStr(1, Documento, ".rpt") <> 0 Then
                    ImpresionNormal = True
                Else
                    ImpresionNormal = False
                End If
            End If
            nomDocu = Documento
        End If
          
        'Nombre fichero .rpt a Imprimir
        nomRPT = nomDocu
        
        conSubRPT = True
        
    Else 'ETIQUETAS
        cadParam = "|"
    
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rSocioEtiq.rpt" '"rComProveEtiq.rpt"
        Titulo = "Etiquetas de Socios" '"Etiquetas de Proveedores"
    
        '===================================================
        '============ PARAMETROS ===========================
        indRPT = 27 'Impresion de Etiquetas de socios
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
          
        'Nombre fichero .rpt a Imprimir
        nomRPT = nomDocu
        
        conSubRPT = False
    End If
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H Seccion
    '--------------------------------------------
'    If txtCodigo(58).Text <> "" Or txtCodigo(59).Text <> "" Then
'        campo = "{rsocios_seccion.codsecci}"
'        'Parametro Desde/Hasta seccion
'        If Not PonerDesdeHasta(campo, "N", 58, 59, "") Then Exit Sub
'    End If
    
    '[Monica]11/11/2013: cambio para castelduc seccion o fase
    If vParamAplic.Cooperativa <> 5 Then
        
        If txtCodigo(58).Text <> "" Then
            ' solo se sacan los socios que no esten dados de baja
            If Not AnyadirAFormula(cadselect, "{rsocios_seccion.codsecci} = " & txtCodigo(58).Text) Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & txtCodigo(58).Text) Then Exit Sub
        End If
        
        
        'Cadena para seleccion D/H Socio
        '--------------------------------------------
        If txtCodigo(60).Text <> "" Or txtCodigo(61).Text <> "" Then
            campo = "{rsocios_seccion.codsocio}"
            'Parametro Desde/Hasta socio
            If Not PonerDesdeHasta(campo, "N", 60, 61, "") Then Exit Sub
        End If
        
        '====================================================
            
        ' solo se sacan los socios que no esten dados de baja
        If Not AnyadirAFormula(cadselect, "{rsocios_seccion.fecbaja} is null ") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "isnull({rsocios_seccion.fecbaja})") Then Exit Sub

    Else
        '[Monica]11/11/2013: Caso de CASTELDUC si no me dan seccion cojo los datos de la fases
        '                    si me dan la seccion funciona como el resto de cooperativas
        If ComprobarCero(txtCodigo(58).Text) <> 0 Then
            ' solo se sacan los socios que no esten dados de baja
            If Not AnyadirAFormula(cadselect, "{rsocios_seccion.codsecci} = " & txtCodigo(58).Text) Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & txtCodigo(58).Text) Then Exit Sub
            
            'Cadena para seleccion D/H Socio
            '--------------------------------------------
            If txtCodigo(60).Text <> "" Or txtCodigo(61).Text <> "" Then
                campo = "{rsocios_seccion.codsocio}"
                'Parametro Desde/Hasta socio
                If Not PonerDesdeHasta(campo, "N", 60, 61, "") Then Exit Sub
            End If
            
            '====================================================
                
            ' solo se sacan los socios que no esten dados de baja
            If Not AnyadirAFormula(cadselect, "{rsocios_seccion.fecbaja} is null ") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "isnull({rsocios_seccion.fecbaja})") Then Exit Sub
            
        Else
            '[Monica]13/09/2016: antes un select case
            If Combo1(0).ItemData(Combo1(0).ListIndex) <> 0 Then
                ' solo se sacan los socios de la fase que sea
                If Not AnyadirAFormula(cadselect, "{rsocios_pozos.numfases} = " & Combo1(0).ItemData(Combo1(0).ListIndex)) Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rsocios_pozos.numfases} = " & Combo1(0).ItemData(Combo1(0).ListIndex)) Then Exit Sub
            End If
            
            
            'Cadena para seleccion D/H Socio
            '--------------------------------------------
            If txtCodigo(60).Text <> "" Or txtCodigo(61).Text <> "" Then
                campo = "{rsocios_pozos.codsocio}"
                'Parametro Desde/Hasta socio
                If Not PonerDesdeHasta(campo, "N", 60, 61, "") Then Exit Sub
            End If
            
            '====================================================
       
        End If
    
    End If

    '[Monica]11/11/2011: añadimos los socios que esten dados de baja en todas las secciones
    ' solo se sacan los socios que no esten dados de baja
    If Not AnyadirAFormula(cadselect, "{rsocios.fechabaja} is null ") Then Exit Sub
    If Not AnyadirAFormula(cadFormula, "isnull({rsocios.fechabaja})") Then Exit Sub

    '[Monica]23/11/2012: si es escalona o utxera seleccionamos que tipo de socio
    If Option1(0).Value Then    ' solo contado
        If Not AnyadirAFormula(cadselect, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}=""8888888888""") Then Exit Sub
    End If
    If Option1(1).Value Then    ' solo efecto
        If Not AnyadirAFormula(cadselect, "{rsocios.cuentaba<>""8888888888""") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.cuentaba}<>""8888888888""") Then Exit Sub
    End If
'   no hacemos nada
'    If Option1(2).Value Then
'    End If
    
    '[Monica]29/09/2014: para el caso de escalona y de utxera si marcan solo los que NO tienen fecha de revision
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        If chkMail(4).Value = 1 Then
            If Not AnyadirAFormula(cadselect, "({rsocios.fechanac} is null or {rsocios.fechanac}='')") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "(isnull({rsocios.fechanac}) or {rsocios.fechanac}='')") Then Exit Sub
        End If
    End If



    '[Monica]08/11/2012: solo los socios que no tengan situacion de bloqueo
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Situacion = SituacionesBloqueo
        If Situacion <> "" Then
            Situacion = Mid(Situacion, 1, Len(Situacion) - 1)
            If Not AnyadirAFormula(cadFormula, "not {rsocios.codsitua} in [" & Situacion & "]") Then Exit Sub
            If Not AnyadirAFormula(cadselect, "not rsocios.codsitua in (" & Situacion & ")") Then Exit Sub
        End If
    End If
    
    '[Monica]23/12/2011: si estamos mandando sms miramos los que tienen nro de movil
    If chkMail(1).Value = 1 Then
        If Not AnyadirAFormula(cadselect, "not {rsocios.movsocio} is null and {rsocios.movsocio}<>''") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "not isnull({rsocios.movsocio}) and {rsocios.movsocio}<>''") Then Exit Sub
        
        ' solo saldran los que no tegan email
        If Check3.Value = 1 Then
            If Not AnyadirAFormula(cadselect, "({rsocios.maisocio} is null or {rsocios.maisocio}='')") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "(isnull({rsocios.maisocio}) or {rsocios.maisocio}='')") Then Exit Sub
        End If
    End If
    ' si es un correo electronico miramos solo los que tienen mail
    If chkMail(0).Value = 1 Then
        If Not AnyadirAFormula(cadselect, "not {rsocios.maisocio} is null and {rsocios.maisocio}<>''") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "not isnull({rsocios.maisocio}) and {rsocios.maisocio}<>''") Then Exit Sub
    End If
    
    '[Monica]03/03/2015: solo en el caso de que sean cartas miramnos si quieren los que no tienen correo electrónico
    '                    y son escalona o Utxera
    If chkMail(0).Value = 0 And chkMail(1).Value = 0 And chkMail(5).Value = 1 And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) Then
        If chkMail(5).Value = 1 Then
            If Not AnyadirAFormula(cadselect, "({rsocios.maisocio} is null or {rsocios.maisocio}='')") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "(isnull({rsocios.maisocio}) or {rsocios.maisocio}='')") Then Exit Sub
        End If
    End If
    
    
    ' hasta aqui
    
    '=========================================================================================
    '[Monica]20/01/2012: Añadida la condicion de marca de correo del socio
    '                    Sólo se tendrá en cuenta en cartas y etiquetas. No en los sms
    '=========================================================================================
    If chkMail(1).Value = 0 Then
        If chkMail(2).Value = 1 Then
            If Not AnyadirAFormula(cadselect, "{rsocios.correo} = 1") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.correo} = 1") Then Exit Sub
        End If
    End If
    
        
    'Parametro a la Atencion de
    If txtCodigo(62).Text <> "" Then
        cadParam = cadParam & "pAtencion=""Att. " & txtCodigo(62).Text & """|"
    Else
        cadParam = cadParam & "pAtencion=""""|"
    End If
    numParam = numParam + 1
    
    tabla = "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio"
    
    '[Monica]11/11/2013: Castelduc
    If vParamAplic.Cooperativa = 5 And ComprobarCero(txtCodigo(58).Text) = 0 Then
        tabla = "rsocios_pozos inner join rsocios on rsocios_pozos.codsocio = rsocios.codsocio"
    End If
    
    '[Monica]11/11/2013: Castelduc
    If vParamAplic.Cooperativa <> 5 Then
        '[Monica]22/04/2015: mogente tampoco enlaza con rcampos
        If (vParamAplic.Cooperativa = 3) And Opcionlistado = 305 Then
        
        Else
            '[Monica]16/11/2012: En Utxera no tienen campos por socio
                                               '[Monica]26/11/2014: en turis tampoco
                                                    '[Monica]30/04/2015: en montifrut tampoco
                                                        '[Monica]10/07/2015: en bolbaite tampoco
            If vParamAplic.Cooperativa <> 8 And vParamAplic.Cooperativa <> 1 And vParamAplic.Cooperativa <> 12 And vParamAplic.Cooperativa <> 14 Then
                tabla = "(" & tabla & ") inner join rcampos on rsocios.codsocio = rcampos.codsocio "
            End If
        End If
    End If
    
    
    'ver si hay registros seleccionados para mostrar en el informe
    If Not HayRegParaInforme(tabla, cadselect) Then Exit Sub
    
    Set frmMen = New frmMensajes
    
    If txtCodigo(58).Text = "" Then
        frmMen.cadWHERE = "SELECT distinct rsocios.codsocio,nomsocio,nifsocio FROM rsocios inner join rsocios_pozos on rsocios.codsocio = rsocios_pozos.codsocio where " & cadselect & " order by rsocios.codsocio "
        frmMen.OpcionMensaje = 55 'Etiquetas socios
    Else
        '[Monica]24/10/2014: antes solo tenia cadselect
        frmMen.cadWHERE = "rsocios.codsocio in (select rsocios.codsocio from " & tabla & " where " & cadselect & ")"
        frmMen.OpcionMensaje = 9 'Etiquetas socios
    End If
    
    frmMen.Show vbModal
    Set frmMen = Nothing
    
    If cadselect = "" Then Exit Sub
    
    '[Monica]16/11/2012: añadida la condicion de cooperativa <> 8
    If Documento <> "" And ImpresionNormal And vParamAplic.Cooperativa <> 8 Then
        Set frmMen2 = New frmMensajes
        frmMen2.cadWHERE = " and codsocio in (select rsocios.codsocio from " & tabla & " where " & cadselect & ")"
        frmMen2.OpcionMensaje = 15 'Etiquetas socios
        frmMen2.Show vbModal
        Set frmMen2 = Nothing
        If cadselect = "" Then Exit Sub
    End If
    
    If Opcionlistado = 306 And Me.chkMail(0).Value = 1 Then
        'Enviarlo por e-mail
        IndRptReport = indRPT
        EnviarEMailMulti cadselect, Titulo, nomDocu, tabla ' "rSocioCarta.rpt", Tabla  'email para socios
        cmdCancel_Click (9)
    Else
        If Opcionlistado = 306 And Me.chkMail(1).Value = 1 Then
            OK = True
            EnviarSMS cadselect, Titulo, nomDocu, tabla, OK    ' "rSocioCarta.rpt", Tabla  'email para socios
            If OK Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (9)
            End If
        Else                                                                  '[Monica]22/04/2015: añadimos a mogente
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Or (vParamAplic.Cooperativa = 3 And Opcionlistado = 305) Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            End If
            LlamarImprimir
        End If
    End If
End Sub


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcionlistado
            Case 305, 306 '305: Listado Etiquetas proveedor
                          '306: Listado Cartas a proveedores
                PonerFoco txtCodigo(58)
                
                '[Monica]11/11/2013: castelduc
                If vParamAplic.Cooperativa = 5 Then
                    Combo1(0).ListIndex = 0
                    PonerFocoCmb Combo1(0)
                End If
                
            Case 315, 316 ' envio de facturas por email y facturacion electronica
                PonerFoco txtCodigo(110)
                
            Case 320 ' envio datos comunica entre coopic y agrocitrica
                PonerFoco txtCodigo(5)
            
                Option2(0).Value = True
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim indFrame As Single
Dim devuelve As String
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me
    indCodigo = 0
    NomTabla = ""

    'Ocultar todos los Frames de Formulario
    Me.FrameEtiqProv.visible = False
    Me.FrameEnvioFacMail.visible = False
    CommitConexion
    
    CargarIconos
    
    Me.Option1(2).Value = True
    FrameTipoSocio.visible = (Opcionlistado = 306 And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10))
    FrameTipoSocio.Enabled = (Opcionlistado = 306 And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10))
    
    '[Monica]27/06/2013: marcamos bien si quieren adjuntar archivos
    Frame1.visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) And Opcionlistado = 306
    Frame1.Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) And Opcionlistado = 306
    
    FrameEnvioDatosComunica.visible = False
    
    Select Case Opcionlistado
        Case 305, 306 '305: Etiquetas de proveedor
                      '306: Cartas a proveedor
            indFrame = 9
            H = 7155 '5325
            W = 7035
            If Opcionlistado = 305 Then
                H = 5325
                Me.cmdAceptarEtiqProv.Top = Me.cmdAceptarEtiqProv.Top - 2000
                Me.CmdCancel(9).Top = CmdCancel(9).Top - 2000
            End If
            PonerFrameVisible Me.FrameEtiqProv, True, H, W
            Me.Frame2.visible = (Opcionlistado = 306)
            If (Opcionlistado = 306) Then Me.Label9(1).Caption = "Cartas a Socios"
            Me.Frame3.visible = (chkMail(0).Value = True)
            Me.Frame3.Enabled = (chkMail(0).Value = True)
            Me.Frame4.visible = (chkMail(1).Value = True)
            Me.Frame4.Enabled = (chkMail(1).Value = True)
            txtCodigo(0).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(1).Text = Format(Now, "hh:mm:ss")
            
            '[Monica]11/11/2013: solo para Castelduc si me da fases no hacemos caso de la seccion
            Frame5.Enabled = (vParamAplic.Cooperativa = 5)
            Frame5.visible = (vParamAplic.Cooperativa = 5)
            CargaCombo
            
            
            '[Monica]29/09/2014: solo para el caso de Escalona y de utxera
            chkMail(4).visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            chkMail(4).Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            imgAyuda(3).visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            imgAyuda(3).Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            
            '[Monica]03/03/2015: solo para el caso de Escalona y de utxera si es carta, solo los que no tienen correo electronico
            chkMail(5).visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            chkMail(5).Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            
        Case 315, 316 'Envio masivo de Facturas
            indFrame = 18
            
            If Opcionlistado = 316 Then Me.FrameEnvioFacMail.Width = 5535
            
            H = FrameEnvioFacMail.Height
            W = FrameEnvioFacMail.Width
            PonerFrameVisible FrameEnvioFacMail, True, H, W
            CargarComboTipoMov 1000
            
            chkMail(3).visible = Opcionlistado = 316 'Solo para facturae
            If Opcionlistado = 316 Then
                cmdEnvioMail.Left = 3240
                CmdCancel(indFrame).Left = 4320
                Label14(16).Caption = "Facturacion E"
                cmdEnvioMail.TabIndex = 474
                Check4.Enabled = True
                Check4.visible = True
            Else
                Label14(16).Caption = "Envio facturas por mail"
            End If
            
            '[Monica]28/01/2014: impresion de facturas con arrobas para montifrut
            Check5.visible = (vParamAplic.Cooperativa = 12)
            Check5.Enabled = (vParamAplic.Cooperativa = 12)
            Check5.Value = 0
        
        Case 320
            H = FrameEnvioDatosComunica.Height
            W = FrameEnvioDatosComunica.Width
            PonerFrameVisible FrameEnvioDatosComunica, True, H, W
                
    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.CmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    
End Sub

Private Sub frmCar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Cartas de Socio
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Or (vParamAplic.Cooperativa = 3 And Opcionlistado = 305) Then
            InsertarTemporal CadenaSeleccion
            cadselect = "rsocios_seccion.codsocio IN (" & CadenaSeleccion & ")"
        Else
            If Opcionlistado = 305 Or Opcionlistado = 306 Then 'Socios
                cadFormula = "{rsocios_seccion.codsocio} IN [" & CadenaSeleccion & "]"
                cadselect = "rsocios_seccion.codsocio IN (" & CadenaSeleccion & ")"
            End If
            If vParamAplic.Cooperativa = 5 And txtCodigo(58).Text = "" Then
                cadFormula = "{rsocios.codsocio} IN [" & CadenaSeleccion & "]"
                cadselect = "rsocios.codsocio IN (" & CadenaSeleccion & ")"
            End If
        End If
    Else 'no seleccionamos ningun cliente
        cadFormula = ""
        cadselect = ""
    End If
End Sub

Private Sub frmMen2_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        If Opcionlistado = 306 Then 'Socios
            cadFormula = "{rcampos.codcampo} IN [" & CadenaSeleccion & "]"
            cadselect = "rcampos.codcampo IN (" & CadenaSeleccion & ")"
        End If
    Else 'no seleccionamos ningun cliente
        cadFormula = ""
        cadselect = ""
    End If
End Sub

Private Sub frmMen3_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        cadselect = CadenaSeleccion
    Else 'no seleccionamos ningun archivo
        cadselect = ""
    End If
End Sub



Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Secciones
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Socios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)

    Select Case Index
        Case 39 'Cod. Carta
            indCodigo = 63
            Set frmCar = New frmCartasSocio
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
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgCheck_Click(Index As Integer)
Dim b As Boolean
Dim TotalArray As Integer

    'En el listview3
    b = Index = 1
    For TotalArray = 0 To ListTipoMov(1000).ListCount - 1
        ListTipoMov(1000).Selected(TotalArray) = b
        If (TotalArray Mod 50) = 0 Then DoEvents
    Next TotalArray
    
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
        Case 1, 2
            indCodigo = Index + 4
        Case 3, 4
            indCodigo = Index
   End Select

   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
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

    '[Monica]03/03/2015: solo utxera y escalona
    Select Case Index
        Case 0
            If chkMail(Index) = 1 Then
                chkMail(1).Value = 0
                chkMail(1).Enabled = False
                chkMail(5).Value = 0
                chkMail(5).Enabled = False
            Else
                chkMail(1).Value = 0
                chkMail(1).Enabled = True
                chkMail(5).Value = 0
                chkMail(5).Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            End If
        Case 1
            If chkMail(Index) = 1 Then
                chkMail(0).Value = 0
                chkMail(0).Enabled = False
                chkMail(5).Value = 0
                chkMail(5).Enabled = False
            Else
                chkMail(0).Value = 0
                chkMail(0).Enabled = True
                chkMail(5).Value = 0
                chkMail(5).Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
            End If
        
        Case 5
            If chkMail(Index) = 1 Then
                chkMail(0).Value = 0
                chkMail(0).Enabled = False
                chkMail(1).Value = 0
                chkMail(1).Enabled = False
            Else
                chkMail(0).Value = 0
                chkMail(0).Enabled = True
                chkMail(1).Value = 0
                chkMail(1).Enabled = True
            End If
        
    End Select
    

    Frame3.visible = (chkMail(0).Value = 1)
    Frame3.Enabled = (chkMail(0).Value = 1)
    Frame4.visible = (chkMail(1).Value = 1)
    Frame4.Enabled = (chkMail(1).Value = 1)
    
    '[Monica]22/12/2011: si tenemos metido el numero de carta que me traiga el texto del sms
    If chkMail(1).Value = 1 And txtCodigo(63).Text <> "" Then
        txtCodigo(2).Text = DevuelveValor("select textosms from scartas where codcarta = " & DBSet(txtCodigo(63).Text, "N"))
        If txtCodigo(2).Text = "0" Then txtCodigo(2).Text = ""
    End If
    
    
End Sub






Private Sub Option2_Click(Index As Integer)
    If Index = 0 Or Index = 1 Then
        Me.FrameExportar.Enabled = (Option2(0).Value)
        If Not FrameExportar.Enabled Then
            Me.ChkAlbaranes.Value = 0
            Me.ChkEntradas.Value = 0
            txtCodigo(5).Text = ""
            txtCodigo(6).Text = ""
            txtCodigo(3).Text = ""
            txtCodigo(4).Text = ""
        End If
    End If
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
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
        
        
            ' envio de csv de comunicaciones
            Case 5: KEYFecha KeyAscii, 1 'fecha desde
            Case 6: KEYFecha KeyAscii, 2 'fecha hasta
            Case 3: KEYFecha KeyAscii, 3 'fecha desde
            Case 4: KEYFecha KeyAscii, 4 'fecha hasta
        
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
Dim tabla As String
Dim codCampo As String, nomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        Case 0, 5, 6, 3, 4 ' Fecha de la carta
            PonerFormatoFecha txtCodigo(Index), False
            
        Case 108, 109 ' Fecha factura
            PonerFormatoFecha txtCodigo(Index), True
        
        Case 1
            PonerFormatoHora txtCodigo(Index)
        
        Case 63, 64 'CARTA de la Oferta
            EsNomCod = True
            tabla = "scartas"
            codCampo = "codcarta"
            nomCampo = "descarta"
            Formato = "000"
            Titulo = "cod. de Carta"
                    
        Case 60, 61, 110, 111 'Socio
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = txtCodigo(Index).Text
            
         Case 58, 59 'Cod. Seccion
            EsNomCod = True
            tabla = "rseccion"
            codCampo = "codsecci"
            nomCampo = "nomsecci"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Sección"
            
        Case 106, 107 ' nro de factura
            PonerFormatoEntero txtCodigo(Index)
            
    End Select
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomCampo, codCampo, TipCampo)
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
                
                If Index = 63 And chkMail(1).Value = 1 Then
                    txtCodigo(2).Text = DevuelveValor("select textosms from scartas where codcarta = " & DBSet(txtCodigo(63).Text, "N"))
                    If txtCodigo(2).Text = "0" Then txtCodigo(2).Text = ""
                End If
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomCampo, codCampo, TipCampo)
        End If
    End If
End Sub

Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
    
    Documento = ""
End Sub


Private Sub InicializarTabla()
Dim Sql As String

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

End Sub



Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadselect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
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
        .Opcion = Opcionlistado
        .Titulo = Titulo
        .NombreRPT = nomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub



Private Sub EnviarSMS(cadWHERE As String, cadTit As String, cadRpt As String, cadTabla As String, ByRef EstaOk As Boolean)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer
Dim Direccion As String

Dim NF As Integer
Dim Cad As String
Dim b As Boolean

On Error GoTo EEnviar


    If vParamAplic.SMSclave = "" Or vParamAplic.SMSemail = "" Or vParamAplic.SMSremitente = "" Then
        MsgBox "No tiene configurados los parámetros de Envio de SMS. Revise.", vbExclamation
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    
    If cadTabla = "(rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio) inner join rcampos on rsocios.codsocio = rcampos.codsocio " Or _
       cadTabla = "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio" Or _
       cadTabla = "rsocios_pozos inner join rsocios on rsocios_pozos.codsocio = rsocios.codsocio" Then
        'seleccionamos todos los socios a los que queremos enviar un SMS
        Sql = "SELECT distinct rsocios.codsocio,nomsocio,rsocios.movsocio "
    End If
    Sql = Sql & "FROM " & cadTabla
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cont = 0
    lista = ""
    
    b = True
    
    
    While Not Rs.EOF And b
    'para cada socio enviamos un SMS
        Cad1 = DBLet(Rs.Fields(2), "T") 'movil socio
        
        If Cad1 = "" Then 'no tiene movil
            lista = lista & Format(Rs.Fields(0), "000000") & " - " & Rs.Fields(1) & vbCrLf
            EstaOk = False
        End If
        
        If Cad1 <> "" Then 'HAY movil  --> ENVIAMOS el mensaje
            Direccion = "http://www.afilnet.com/http/sms/?email=" & Trim(vParamAplic.SMSemail) & "&pass=" & Trim(vParamAplic.SMSclave)
            Direccion = Direccion & "&mobile=" & Trim(Cad1) & "&id=" & Trim(vParamAplic.SMSremitente)
            Direccion = Direccion & "&country=0034" & "&sms=" & txtCodigo(2).Text & "&now=" & Format(Check1.Value, "0")
            Direccion = Direccion & "&date=" & Format(txtCodigo(0).Text, "yyyy/mm/dd") & " " & Format(txtCodigo(1).Text, "hh:mm")
            Direccion = Direccion & "&type=" & Format(Check2.Value, "0")
            
            Screen.MousePointer = vbHourglass
            
            Label9(10).Caption = Format(Rs.Fields(0), "000000") & " - " & Rs.Fields(1) & " - " & Rs.Fields(2)
            DoEvents
           
            'Cargamos en el fichero el resultado de enviar un mensaje
            GetFileFromUrl Direccion, App.Path & "\RESULT.TXT"
    
            NF = FreeFile
            Open App.Path & "\RESULT.TXT" For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
            Cad = ""
            Line Input #NF, Cad
            Close NF

            Select Case Mid(Cad, 1, 2)
                Case "OK"
                    espera 2
                
                    Me.Refresh
                    DoEvents
                    espera 0.4
                    cont = cont + 1
                    
                    Sql = "INSERT INTO rsmsenviados (codsocio, movsocio, fechaenvio, horaenvio, texto)"
                    Sql = Sql & " VALUES (" & DBSet(Rs.Fields(0), "N") & "," & DBSet(Cad1, "T") & ","
                    Sql = Sql & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(1).Text, "H") & "," & DBSet(txtCodigo(2).Text, "T") & ")"
                    conn.Execute Sql
            
            
                Case "-1"
                    MsgBox "Error en el Login, usuario o clave incorrectas", vbExclamation
                    EstaOk = False
                Case Else
                    If Mid(Cad, 1, 12) = "Sin Creditos" Then
                        MsgBox "No tiene créditos. Revise", vbExclamation
                        b = False
                    Else
                        If MsgBox("Error en el envio de mensaje al socio " & DBLet(Rs.Fields(0), "N") & ". ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then b = False
                    End If
                    EstaOk = False
            End Select
            
            Screen.MousePointer = vbDefault
            
        End If
        Rs.MoveNext
    Wend
    Label9(10).Caption = ""
    DoEvents
    
    Rs.Close
    Set Rs = Nothing
    
    Screen.MousePointer = vbDefault
   
    'Mostra mensaje con aquellos socios que no tienen móvil
    If lista <> "" Then
        lista = "Socios sin Móvil:" & vbCrLf & vbCrLf & lista
        MsgBox lista, vbInformation
    End If
    
    
EEnviar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Enviando SMS", Err.Description
    End If
End Sub

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

Private Sub InsertarTemporal(cadWHERE As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer
Dim Sql2 As String

    
    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    'seleccionamos todos los socios a los que queremos enviar e-mail
    Sql = "SELECT distinct " & vUsu.Codigo & ", rsocios.codsocio from rsocios where codsocio in (" & cadWHERE & ")"
    
    Sql2 = "insert into tmpinformes (codusu, codigo1) " & Sql
    conn.Execute Sql2

End Sub

Private Sub EnviarEMailMulti(cadWHERE As String, cadTit As String, cadRpt As String, cadTabla As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer

On Error GoTo EEnviar

    Screen.MousePointer = vbHourglass
    
    If cadTabla = "(rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio) inner join rcampos on rsocios.codsocio = rcampos.codsocio " Or _
       cadTabla = "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio" Or _
       cadTabla = "rsocios_pozos inner join rsocios on rsocios_pozos.codsocio = rsocios.codsocio" Then
       
        'seleccionamos todos los socios a los que queremos enviar e-mail
        Sql = "SELECT distinct rsocios.codsocio,nomsocio,maisocio, maisocio "
        If ImpresionNormal Then
            Sql = Sql & "FROM " & cadTabla
        Else
            Sql = Sql & "FROM " & "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio"
        End If
    ElseIf cadTabla = "sclien" Then
        'seleccionamos todos los clientes a los que queremos enviar e-mail
        Sql = "SELECT codclien,nomclien,maiclie1,maiclie2 "
        Sql = Sql & "FROM " & cadTabla
    End If
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
        Cad2 = DBLet(Rs.Fields(3), "T") 'e-mail compras
        
        If Cad1 = "" And Cad2 = "" Then 'no tiene e-mail
'              MsgBox "Sin mail para el proveedor: " & Format(RS!codProve, "000000") & " - " & RS!nomprove, vbExclamation
              lista = lista & Format(Rs.Fields(0), "000000") & " - " & Rs.Fields(1) & vbCrLf
        ElseIf Cad1 <> "" And Cad2 <> "" Then 'tiene 2 e-mail
            'ver a q e-mail se va a enviar (administracion, compras)
            If cadTabla = "(rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio) inner join rcampos on rsocios.codsocio = rcampos.codsocio " Or _
                cadTabla = "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio" Or _
                cadTabla = "rsocios_pozos inner join rsocios on rsocios_pozos.codsocio = rsocios.codsocio" Then
                If Me.OptMailCom(0).Value = True Then Cad1 = Cad2
            Else
                If Me.OptMailCom(1).Value = True Then Cad1 = Cad2
            End If
        Else 'alguno de los 2 tiene valor
            If Cad2 <> "" Then Cad1 = Cad2  'e-mail para compras
        End If
        
        If Cad1 <> "" Then 'HAY email --> ENVIAMOS e-mail
            Label9(10).Caption = Format(Rs.Fields(0), "000000") & " - " & Rs.Fields(1) & " - " & Rs.Fields(2)
            DoEvents


            If ImpresionNormal Then
                With frmImprimir
                    .OtrosParametros = cadParam
                    .NumeroParametros = numParam
                    If cadTabla = "(rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio) inner join rcampos on rsocios.codsocio = rcampos.codsocio " Or _
                       cadTabla = "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio" Or _
                       cadTabla = "rsocios_pozos inner join rsocios on rsocios_pozos.codsocio = rsocios.codsocio" Then
                        Sql = "{rsocios.codsocio}=" & Rs.Fields(0)
                        .Opcion = 306
                    Else
                        Sql = "{sclien.codclien}=" & Rs.Fields(0)
                        .Opcion = 91
                    End If
                    .FormulaSeleccion = Sql
                    .EnvioEMail = True
                    CadenaDesdeOtroForm = "GENERANDO"
                    .Titulo = cadTit
                    .NombreRPT = cadRpt
                    .ConSubInforme = True
                    .Show vbModal
    
                    If CadenaDesdeOtroForm = "" Then
                    'si se ha generado el .pdf para enviar
                        Sql = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                        Sql = Sql & " VALUES (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "T") & "," & DBSet(Cad1, "T") & ")"
                        conn.Execute Sql
                
                        Me.Refresh
                        DoEvents
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
            Else
                If CadenaDesdeOtroForm = "" Then
                'si se ha generado el .pdf para enviar
                    Sql = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                    Sql = Sql & " VALUES (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "T") & "," & DBSet(Cad1, "T") & ")"
                    conn.Execute Sql
            
                    Me.Refresh
                    DoEvents
                    
'[Monica]28/06/2013: quito el espera para agilizar el proceso
'                    espera 0.4

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
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
      
    If cont > 0 Then
        espera 0.4
        If cadTabla = "(rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio) inner join rcampos on rsocios.codsocio = rcampos.codsocio " Or _
           cadTabla = "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio" Or _
           cadTabla = "rsocios_pozos inner join rsocios on rsocios_pozos.codsocio = rsocios.codsocio" Then
            Sql = "Carta: " & txtNombre(63).Text & "|"
            
            '[Monica]08/07/2011: si no hay a la atencion no se pone nada en el cuerpo del mensaje
            '                    añadida la condicion
            If txtCodigo(62).Text <> "" Then
                Sql = Sql & "Att : " & txtCodigo(62).Text & "|"
            End If
        Else
            Sql = "Carta: " & txtNombre(63).Text & "|"
            Sql = Sql & "Att : " & txtCodigo(62).Text & "|"
        End If
       
        If Not ImpresionNormal Then
            Set frmMen3 = New frmMensajes
            frmMen3.cadWHERE = ""
            frmMen3.OpcionMensaje = 40 'archivos a seleccionar
            frmMen3.Show vbModal
            Set frmMen3 = Nothing
            If cadselect = "" Then Exit Sub
        End If
       
        frmEMail.Opcion = 2
        frmEMail.DatosEnvio = Sql
        frmEMail.CodCryst = IndRptReport
        If Not ImpresionNormal Then
            frmEMail.Opcion = 5
            frmEMail.Ficheros = cadselect
        Else
            frmEMail.Ficheros = ""
        End If
        frmEMail.Show vbModal

        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute Sql
        
        'Borrar la carpeta con temporales
        If ImpresionNormal Then Kill App.Path & "\temp\*.pdf"
    End If
    
    Screen.MousePointer = vbDefault
   
    'Mostra mensaje con aquellos proveedores que no tienen e-mail
    If lista <> "" Then
        If cadTabla = "(rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio) inner join rcampos on rsocios.codsocio = rcampos.codsocio " Or _
           cadTabla = "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio" Or _
           cadTabla = "rsocios_pozos inner join rsocios on rsocios_pozos.codsocio = rsocios.codsocio" Then
            lista = "Socios sin e-mail:" & vbCrLf & vbCrLf & lista
        Else
            lista = "Clientes sin e-mail:" & vbCrLf & vbCrLf & lista
        End If
        MsgBox lista, vbInformation
    End If
    
EEnviar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Enviando Informe por e-mail", Err.Description
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute Sql
    End If
End Sub


Private Sub CargarIconos()
Dim i As Integer

    Me.imgBuscarOfer(35).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    For i = 37 To 39
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 56 To 57
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i



End Sub

Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean

    b = True
    
    '[Monica]11/11/2013: dejamos poner unicamente las fases de Castelduc
    If vParamAplic.Cooperativa = 5 Then
        If Combo1(0).ListIndex = -1 And ComprobarCero(txtCodigo(58).Text) = 0 Then
            b = False
            MsgBox "Debe introducir sección o fase. Revise.", vbExclamation
            PonerFocoCmb Combo1(0)
        End If
    End If
    
    '[Monica]11/11/2013: en Castelduc dejamos poner unicamente las fases
    If txtCodigo(58).Text = "" And vParamAplic.Cooperativa <> 5 Then
        MsgBox "Debe introducir un valor en la Sección. Revise.", vbExclamation
        b = False
        PonerFoco txtCodigo(58)
    End If
    
    
    If b And Opcionlistado = 306 And chkMail(1).Value = 1 Then
        If txtCodigo(0).Text = "" Then
            MsgBox "Debe introducir un valor en la Fecha del SMS. Revise.", vbExclamation
            b = False
            PonerFoco txtCodigo(0)
        End If
        If b And txtCodigo(1).Text = "" Then
            MsgBox "Debe introducir un valor en la Hora del SMS. Revise.", vbExclamation
            b = False
            PonerFoco txtCodigo(1)
        End If
        If b And txtCodigo(2).Text = "" Then
            MsgBox "Debe introducir un valor en el Texto del SMS. Revise.", vbExclamation
            b = False
            PonerFoco txtCodigo(2)
        End If
    End If
    
    '[Monica]19/10/2012: comprobamos que si vamos a enviar mas de un documento es por email.
    If b And Opcionlistado = 306 Then
        Documento = DevuelveDesdeBDNew(cAgro, "scartas", "documrpt", "codcarta", txtCodigo(63).Text, "N")
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
    
    DatosOk = b
    
End Function


Private Function GeneracionEnvioMail(ByRef Rs As ADODB.Recordset) As Boolean
Dim Tipo As Integer
Dim TipoRec As String ' tipo de factura a la que rectifica
Dim Sql5 As String
Dim EsComplemen As Byte
Dim letraser As String

    On Error GoTo EGeneracionEnvioMail
    GeneracionEnvioMail = False
    
    cadselect = "Select * from tmpinformes where codusu =" & vUsu.Codigo & " ORDER BY importe1,codigo1"
    Rs.Open cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CodClien = ""
    While Not Rs.EOF
    
        InicializarVbles
        '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        
        '[Monica]28/01/2013: impresion con arrobas o sin arrobas para Montifrut
        If vParamAplic.Cooperativa = 12 Then
            If Check5.Value = 1 Then
                cadParam = cadParam & "pConArrobas=1|"
            Else
                cadParam = cadParam & "pConArrobas=0|"
            End If
            numParam = numParam + 1
        End If
        
        
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    
        Label14(22).Caption = "Factura: " & Rs!importe1 & " " & Rs!Nombre1
        Label14(22).Refresh
        
        Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
        Dim nomDocu As String 'Nombre de Informe rpt de crystal
        
        'Facturas socios
        Select Case Mid(Rs!Nombre1, 1, 3)
            Case "FRS" ' Impresion de facturas rectificativas
                       ' hacemos caso del codtipom que rectifica
                  TipoRec = DevuelveValor("select rectif_codtipom from rfactsoc where numfactu = " & DBSet(Rs!importe1, "N") & " and codtipom = " & DBSet(Rs!Nombre1, "T") & " and fecfactu = " & DBSet(Rs!fecha1, "T"))
                       
                  Select Case Mid(TipoRec, 1, 3)
                        Case "FLI"
                            indRPT = 38 'Impresion de Factura Socio de Industria
                        Case Else
                            Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Mid(Rs!Nombre1, 1, 3), "T"))
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
                Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Mid(Rs!Nombre1, 1, 3), "T"))
                If Tipo >= 7 And Tipo <= 10 Then
                    indRPT = 42 'Imporesion de Facturas de Bodega o Almazara
                Else
                    indRPT = 23 'Impresion de Factura Socio

'[Monica]07/02/2012: Hemos marcado las facturas que son complementarias, ya no hace falta esto
'
'                    'Si es complementaria le pasamos el parametro
'                    cadParam = cadParam & "pComplem=" & Rs!campo1 & "|"
'                    numParam = numParam + 1

                    '[Monica]17/02/2016: si la factura es de anticipo
                    If vParamAplic.Cooperativa = 4 Then
                        If Mid(Rs!Nombre1, 1, 3) = "FAA" Then
                            If ConDetalle Then
                                cadParam = cadParam & "pDetalle=1|"
                            Else
                                cadParam = cadParam & "pDetalle=0|"
                            End If
                        Else
                            cadParam = cadParam & "pDetalle=1|"
                        End If
                        numParam = numParam + 1
                    End If
                End If
       End Select
        
       If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Function
       'Nombre fichero .rpt a Imprimir
        
        
       cadFormula = "({rfactsoc.codtipom}='" & Trim(Rs!Nombre1) & "') "
       cadFormula = cadFormula & " AND ({rfactsoc.numfactu}=" & Rs!importe1 & ") "
       cadFormula = cadFormula & " AND ({rfactsoc.fecfactu}= Date(" & Year(Rs!fecha1) & "," & Month(Rs!fecha1) & "," & Day(Rs!fecha1) & "))"

   
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
        
        
        If Opcionlistado = 315 Then
   
            EsperaFichero
            
            FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Rs!Nombre1 & Format(Rs!importe1, "0000000") & ".pdf" 'RS!importe1 & Format(RS!Codigo1, "0000000") & ".pdf"
        Else
            ' Se tiene que llamar base_letradeserie_numFactura_yyyymmdd_tipo.pdf
            ' tipo: F = Factura de cliente
            '       S = Factura de socio
            
            ' Sacamos la letra de serie
            letraser = ""
            letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!Nombre1, "T"))
            
            '[Monica]07/03/2013: tengo que incluir el nro de la base de datos, antes era "ariagro_" y la letra de serie
            cadFormula = vEmpresa.BDAriagro & "_" & Trim(letraser) & "_" & Rs!importe1 & "_" & Format(Rs!fecha1, "yyyymmdd") & "_S" & Rs!Nombre1 & ".pdf"
            cadFormula = vParamAplic.PathFacturaE & "\" & cadFormula
            
            Label14(22).Caption = cadFormula
            Label14(22).Refresh
        
            FileCopy App.Path & "\docum.pdf", cadFormula
            
            'Ha copiado, luego yo la pongo como en facturaE
            cadFormula = "UPDATE rfactsoc set enfacturae=1 WHERE codtipom = '" & Rs!Nombre1 & "' AND numfactu=" & Rs!importe1
            cadFormula = cadFormula & " AND fecfactu='" & Format(Rs!fecha1, FormatoFecha) & "'"
            
            conn.Execute cadFormula
        
        End If
       
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    Set Rs = Nothing
    GeneracionEnvioMail = True
    Exit Function
EGeneracionEnvioMail:
       MuestraError Err.Number
End Function

Private Sub EsperaFichero()
Dim Cad As String
Dim T1 As Single

    On Error GoTo eEsperaFichero
    
    T1 = Timer
    Do
        Cad = Dir(App.Path & "\docum.pdf", vbArchive)
        If Cad = "" Then
            If Timer - T1 > 2 Then Cad = "SAL"
        End If
    Loop Until Cad <> ""
    
eEsperaFichero:
    If Err.Number <> 0 Then Err.Clear

End Sub


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
    Me.Refresh
    DoEvents
End Sub

Private Sub CargarComboTipoMov(indice As Integer)
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo

'Lo cargamos con los valores de la tabla stipom que tengan tipo de documento=Albaranes (tipodocu=1)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim i As Byte

    On Error GoTo ECargaCombo
    
    '[Monica]18/02/2013: excluimos los movimientos de facturas varias
    Sql = "select codtipom, nomtipom from usuarios.stipom where (tipodocu <> 0) and tipodocu <> 12"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 0
    
    ListTipoMov(indice).Clear
    
    'LOS TIKCETS NO LOS ENVIO POR MAIL
    While Not Rs.EOF
        ListTipoMov(indice).AddItem Rs.Fields(0).Value & "-" & Rs.Fields(1).Value
        'ListTipoMov(indice).List (ListTipoMov(indice).NewIndex)
        ListTipoMov(indice).Selected((ListTipoMov(indice).NewIndex)) = True
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    
ECargaCombo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function SituacionesBloqueo() As String
Dim Sql As String
Dim cadena As String
Dim Rs As ADODB.Recordset

    cadena = ""

    SituacionesBloqueo = cadena

    Sql = "select codsitua from rsituacion where bloqueo = 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        cadena = cadena & DBLet(Rs!codsitua) & ","
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    SituacionesBloqueo = cadena

End Function


Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer
Dim Sql As String
Dim Rs As ADODB.Recordset

    Combo1(0).Clear
    
    Sql = "select distinct numfases from rsocios_pozos "

    Combo1(0).AddItem "Todas" 'campo del codigo
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql = "Fase " & Rs.Fields(0).Value
        
        Combo1(0).AddItem Sql 'campo del codigo
        Combo1(0).ItemData(Combo1(0).NewIndex) = Rs.Fields(0).Value
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    
End Sub


Public Function ProcesarFicheroComunicacion2(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eProcesarFicheroComunicacion2
    
    ProcesarFicheroComunicacion2 = False
    
    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    DoEvents
    ' PROCESO DEL FICHERO VENTAS.TXT

    b = True

    While Not EOF(NF)
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        DoEvents
        b = ComprobarRegistro(Cad)
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" Then
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        DoEvents
        b = ComprobarRegistro(Cad)
    End If
    
    
    '[Monica]31/10/2018: para el caso de picassent calculamos los gastos de acarreo y recoleccion
    If vParamAplic.Cooperativa = 2 And NotasParaGastos <> "" Then
        CalcularGastosNotas Mid(NotasParaGastos, 2)
    End If
    
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFicheroComunicacion2 = True
    Exit Function

eProcesarFicheroComunicacion2:
    ProcesarFicheroComunicacion2 = False
End Function

' Funcion para el calculo de gastos de las notas de las que se han insertado
Private Sub CalcularGastosNotas(vNotas As String)
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCalcularGastosNotas


    Sql = "select * from rclasifica where numnotac in (" & vNotas & ")"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        CalcularGastos Rs
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    Exit Sub
    
eCalcularGastosNotas:
    MuestraError Err.Number, "Calcular Gastos de notas", Err.Description
End Sub

Private Sub CalcularGastos(ByRef Rs As ADODB.Recordset)
Dim Rs1 As ADODB.Recordset
Dim SQL1 As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency
Dim GasRecol As Currency
Dim GasAcarreo As Currency
Dim KilosTria As Long
Dim KilosNet As Long
Dim KilosTrans As Long
Dim EurDesta As Currency
Dim EurRecol As Currency
Dim PrecAcarreo As Currency
Dim i As Integer
Dim Sql As String

    On Error Resume Next
    
    GasRecol = 0
    GasAcarreo = 0
    
    SQL1 = "select eurdesta, eurecole from variedades where codvarie = " & DBSet(Rs!codvarie, "N")
    
    Set Rs1 = New ADODB.Recordset
    Rs1.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs1.EOF Then
        EurDesta = DBLet(Rs1.Fields(0).Value, "N")
        EurRecol = DBLet(Rs1.Fields(1).Value, "N")
    End If

    Set Rs1 = Nothing

    KilosNet = DBLet(Rs!KilosNet, "N")

    '[Monica]14/10/2010: para picassent los kilos son los de transporte
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then KilosNet = DBLet(Rs!KilosTra, "N")


    'recolecta socio
    If DBLet(Rs!Recolect, "N") = 1 Then
        Sql = "select sum(kilosnet) from rclasifica_clasif, rcalidad  where numnotac = " & DBSet(Rs!nunotac, "N")
        Sql = Sql & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
        Sql = Sql & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
        Sql = Sql & " and rcalidad.gastosrec = 1"
        
        KilosTria = TotalRegistros(Sql)
        
        GasRecol = Round2(KilosTria * EurRecol, 2)
        
    Else
    'recolecta cooperativa
        If DBLet(Rs!tiporecol, "N") = 0 Then
            'horas
            'gastosrecol = horas * personas * rparam.(costeshora + costesegso)
            GasRecol = Round2(HorasDecimal(DBLet(Rs!horastra, "N")) * CCur(DBLet(Rs!numtraba, "N")) * (vParamAplic.CosteHora + vParamAplic.CosteSegSo), 2)
        Else
            'destajo
            GasRecol = Round2(KilosNet * EurDesta, 2)
        End If
    End If
    
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then GasRecol = Round2(KilosNet * EurDesta, 2)
    

    PrecAcarreo = 0
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", DBLet(Rs!Codtarif, "N"), "N")
    If Sql <> "" Then
        PrecAcarreo = CCur(Sql)
    End If
    
    If vParamAplic.Cooperativa = 4 Then
        Sql = "select sum(kilosnet) from rclasifica_clasif, rcalidad  where numnotac = " & DBSet(Rs!NumNotac, "N")
        Sql = Sql & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
        Sql = Sql & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
        Sql = Sql & " and rcalidad.gastosrec = 1"
        
        KilosTria = TotalRegistros(Sql)
        
        If DBLet(Rs!trasnportadopor, "N") = 1 Then ' transportado por socio
            GasAcarreo = Round2(PrecAcarreo * KilosTria, 2)
        Else
            GasAcarreo = Round2(PrecAcarreo * KilosNet, 2)
        End If
        ' cargamos los kilos de transporte
        ' Text1(23).text = Format(KilosTria, "###,##0")
    Else
        GasAcarreo = Round2(PrecAcarreo * KilosNet, 2)
    End If
    
'    Text1(16).Text = Format(GasRecol, "#,##0.00")
'    Text1(15).Text = Format(GasAcarreo, "#,##0.00")
    
    Sql = "update rclasifica set imprecol = " & DBSet(GasRecol, "N") & ", impacarr = " & DBSet(GasAcarreo, "N")
    If vParamAplic.Cooperativa = 4 Then
        Sql = Sql & ", kilostrans = " & DBSet(KilosTria, "N")
    End If
    Sql = Sql & " where numnotac = " & DBSet(Rs!NumNotac, "N")
    conn.Execute Sql

End Sub






Private Function ComprobarRegistro(Cad As String) As Boolean
Dim Sql As String
Dim Id As String
Dim Fecha As String
Dim Usuario As String
Dim Tipo As String
Dim tabla As String
Dim Observaciones As String
Dim SqlEjec As String
Dim SqlActualizar As String
Dim i As Long
Dim vAux As String
Dim vNota As String
Dim Mens As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = False

    Id = RecuperaValorNew(Cad, ";", 1)
    Fecha = RecuperaValorNew(Cad, ";", 2)
    Usuario = RecuperaValorNew(Cad, ";", 3)
    Tipo = RecuperaValorNew(Cad, ";", 4)
    tabla = RecuperaValorNew(Cad, ";", 5)
    SqlEjec = RecuperaValorNew(Cad, ";", 6)
    Observaciones = RecuperaValorNew(Cad, ";", 7)
    
    
    Mens = ""
    
    ' id existente
    Sql = "select count(*) from comunica_rec where id = " & DBSet(Id, "N")
    If TotalRegistros(Sql) <> 0 Then
        Mens = "Id Existente"
        Sql = "insert into tmpinformes (codusu, fecha1, importe1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Fecha, "F") & "," & _
              DBSet(Id, "N") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
    End If
    
    'Comprobamos fechas
    If Not IsDate(Fecha) Then
        Mens = "Fecha incorrecta"
        Sql = "insert into tmpinformes (codusu, fecha1, importe1,  nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Fecha, "F") & "," & _
              DBSet(Id, "N") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
    End If
    
    
    'vemos si da error el sql a ejecutar y lo registramos
    
    conn.Execute SqlEjec
    
    
    If Mens = "" Then
        '[Monica]31/10/2018: en el caso de que sea una entrada de clasificacion me las guardo para calcular los costes
        If tabla = "rclasifica" And Tipo = "I" Then
            'Parte correspondiente a encontrar el numero de nota de la cadena de insert
            i = InStr(1, SqlEjec, "values (")
            If i > 0 Then
                i = i + 8
                vAux = Mid(SqlEjec, i)
                vNota = RecuperaValorNew(vAux, ",", 1)
                NotasParaGastos = NotasParaGastos & "," & vNota
            End If
        End If
    
        Sql = "insert into comunica_rec (id,fechacreacion,usuariocreacion,tipo,tabla,sqlaejecutar,observaciones) values ("
        Sql = Sql & DBSet(Id, "N") & "," & DBSet(Fecha, "FH") & "," & DBSet(Usuario, "T") & "," & DBSet(Tipo, "T") & ","
        Sql = Sql & DBSet(tabla, "T") & "," & DBSet(SqlEjec, "T") & "," & DBSet(Observaciones, "T") & ")"
        
        conn.Execute Sql
    
        SqlActualizar = "update comunica_rec set fechaactualizacion = " & DBSet(Now(), "FH")
        SqlActualizar = SqlActualizar & ", usuarioactualizacion = " & DBSet(vUsu.Nombre, "T")
        SqlActualizar = SqlActualizar & " where id = " & DBSet(Id, "N")
        
        conn.Execute SqlActualizar
    End If
    
    ComprobarRegistro = True
    Exit Function
    
eComprobarRegistro:
'    MuestraError Err.Number, "Comprobar registro", Err.Description
    Mens = "Error sql"
    Sql = "insert into tmpinformes (codusu, fecha1, importe1, nombre1, text1) values (" & _
              vUsu.Codigo & "," & DBSet(Fecha, "F") & "," & _
              DBSet(Id, "N") & "," & DBSet(Mens, "T") & "," & DBSet(Err.Description, "T") & ")"
    conn.Execute Sql
End Function



Public Function ProcesarFicheroComunicacion() As Boolean
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim b As Boolean

Dim SqlActualizar As String
Dim SqlEjecutar As String
Dim cadError As String

    On Error GoTo eProcesarFicheroComunicacion
    
    ProcesarFicheroComunicacion = False
    
    cadError = ""
    
    Sql = "select * from comunica_rec where fechaactualizacion is null"
    Sql = Sql & " order by fechacreacion, tipo"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
    
        cadError = "Id: " & Rs!Id & vbCrLf & "Tabla: " & Rs!tabla & vbCrLf & "SQL: " & Rs!SQLAEJECUTAR
    
        SqlEjecutar = DBLet(Rs!SQLAEJECUTAR, "T")
        
        conn.Execute SqlEjecutar
    
        SqlActualizar = "update comunica_rec set fechaactualizacion = " & DBSet(Now(), "FH")
        SqlActualizar = SqlActualizar & ", usuarioactualizacion = " & DBSet(vUsu.Nombre, "T")
        SqlActualizar = SqlActualizar & " where id = " & DBSet(Rs!Id, "N")
        
        conn.Execute SqlActualizar
    
        Rs.MoveNext
    Wend
    
    ProcesarFicheroComunicacion = True
    Exit Function

eProcesarFicheroComunicacion:
    MuestraError Err.Number, "Procesar fichero comunicación:" & vbCrLf & cadError, Err.Description
    ProcesarFicheroComunicacion = False
End Function





Public Function HayEntradasModificadas(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eProcesarFicheroComunicacion2
    
    HayEntradasModificadas = False
    
    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    DoEvents
    ' PROCESO DEL FICHERO VENTAS.TXT

    b = True

    While Not EOF(NF) 'And B
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        DoEvents
        b = ComprobarEntrada(Cad)
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
'    If B Then
        If Cad <> "" Then
            i = i + 1
            
            Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
            lblProgres(1).Caption = "Linea " & i
            Me.Refresh
            DoEvents
            b = ComprobarEntrada(Cad)
        End If
'    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    HayEntradasModificadas = True 'B
    Exit Function

eProcesarFicheroComunicacion2:
    HayEntradasModificadas = False
End Function

Private Function ComprobarEntrada(Cad As String) As Boolean
Dim Sql As String
Dim Id As String
Dim Fecha As String
Dim Usuario As String
Dim Tipo As String
Dim tabla As String
Dim Observaciones As String
Dim SqlEjec As String
Dim SqlActualizar As String

Dim Mens As String

    On Error GoTo eComprobarRegistro

    ComprobarEntrada = False

    Id = RecuperaValorNew(Cad, ";", 1)
    Fecha = RecuperaValorNew(Cad, ";", 2)
    Usuario = RecuperaValorNew(Cad, ";", 3)
    Tipo = RecuperaValorNew(Cad, ";", 4)
    tabla = RecuperaValorNew(Cad, ";", 5)
    SqlEjec = RecuperaValorNew(Cad, ";", 6)
    Observaciones = RecuperaValorNew(Cad, ";", 7)
    
    Mens = ""
    
    'damos aviso de las entradas comunicadas que han sido modificadas
    If tabla = "rclasifica" And Tipo = "U" Then
        Mens = "Entrada comunicada modif."
        
        Sql = "insert into tmpinformes (codusu, fecha1, importe1,  nombre1, text1) values (" & _
              vUsu.Codigo & "," & DBSet(Fecha, "F") & "," & _
              DBSet(Id, "N") & "," & DBSet(Mens, "T") & "," & DBSet(Observaciones, "T") & ")"
        
        conn.Execute Sql
    End If
    
    'idem de los albaranes
    If tabla = "albaran" And Tipo = "U" Then
        Mens = "Albarán comunicado modif."
        
        Sql = "insert into tmpinformes (codusu, fecha1, importe1,  nombre1, text1) values (" & _
              vUsu.Codigo & "," & DBSet(Fecha, "F") & "," & _
              DBSet(Id, "N") & "," & DBSet(Mens, "T") & "," & DBSet(Observaciones, "T") & ")"
        
        conn.Execute Sql
    End If
    
    ComprobarEntrada = True
    Exit Function
    
eComprobarRegistro:
    MuestraError Err.Number, "Comprobar Entrada", Err.Description
End Function


