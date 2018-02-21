VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmListadoOfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   10305
   Icon            =   "frmListadoOfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
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
   Begin VB.Frame FrameEnvioFacMail 
      Height          =   6015
      Left            =   30
      TabIndex        =   39
      Top             =   90
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
         Text            =   "frmListadoOfer.frx":0097
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
         ItemData        =   "frmListadoOfer.frx":009D
         Left            =   1545
         List            =   "frmListadoOfer.frx":009F
         Style           =   1  'Checkbox
         TabIndex        =   48
         Top             =   4215
         Width           =   3825
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   4680
         Picture         =   "frmListadoOfer.frx":00A1
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   5040
         Picture         =   "frmListadoOfer.frx":01EB
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
         Picture         =   "frmListadoOfer.frx":0335
         Top             =   2670
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   34
         Left            =   3780
         Picture         =   "frmListadoOfer.frx":03C0
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

Public OpcionListado As Integer
    '(ver opciones en frmListado)
        
        
        
    '315:  Envio por mail de las facturas
        
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
Private CadParam As String 'cadena con los parametros q se pasan a Crystal R.
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

Dim SqlDeta As String
Dim ConDetalle As Boolean



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check5_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkmail_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub


Private Sub cmdEnvioMail_Click()
Dim Rs As ADODB.Recordset

Dim T1 As Single

    'El proceso constara de varias fases.
    'Fase 1: Montar el select y ver si hay registros
    'Fase 2: Preparar carpetas para los pdf
    'Fase 3: Generar para cada factura (una a una) del select su pdf
    'Fase 4: Enviar por mail, adjuntando los archivos correspondientes
    If OpcionListado = 315 Then
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
    
    If cadSelect <> "" Then cadSelect = cadSelect & " AND "
    cadSelect = cadSelect & NomTabla
    cadSelect = " WHERE " & cadSelect

    
    Set Rs = New ADODB.Recordset
    DoEvents
        
    If OpcionListado = 316 Then
        If Me.Check4.Value = 0 Then
            If cadSelect <> "" Then cadSelect = cadSelect & " AND "
            cadSelect = cadSelect & " (rfactsoc.enfacturae = 0 )"
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
    
    NomTabla = "Select codtipom,numfactu,codsocio,fecfactu,totalfac,esliqcomplem from rfactsoc  " & cadSelect
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
        If OpcionListado = 316 Then
            MsgBox "Ningúna factura para traspasar a FacturaE", vbExclamation
        Else
            MsgBox "Ningun dato a enviar por mail", vbExclamation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Numero de registros
    NomTabla = NumRegElim
    
    
    If OpcionListado = 315 Then
        
        'AHora ya tengo todos los datos de las facturas que voy  a imprimir
        
        cadSelect = "Select codsocio,maisocio "
        cadSelect = cadSelect & " as email from tmpinformes,rsocios where codusu = " & vUsu.Codigo & " and codsocio=codigo1"
        cadSelect = cadSelect & " and (maisocio is null or maisocio = '') "
        cadSelect = cadSelect & " group by codsocio "
        Rs.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            Rs.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cadSelect = "DELETE from tmpinformes where codusu =" & vUsu.Codigo & " and codigo1 ="
            While Not Rs.EOF
                conn.Execute cadSelect & Rs!Codsocio
                Rs.MoveNext
            Wend
            Rs.Close
            
            
            cadSelect = "Select count(*) from tmpinformes where codusu =" & vUsu.Codigo
            Rs.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                cadSelect = "Hay " & NumRegElim & " facturas para enviar por mail." & vbCrLf & "¿Continuar?"
                If MsgBox(cadSelect, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
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
        cadSelect = "Hay " & NumRegElim & " facturas para integrar con facturaE." & vbCrLf & "¿Continuar?"
        If MsgBox(cadSelect, vbQuestion + vbYesNo) = vbNo Then
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
    
    If vParamAplic.Cooperativa = 0 Then T1 = Timer
    
    If GeneracionEnvioMail(Rs) Then NumRegElim = 1
    
    Label14(22).Caption = "Preparando envia email"
    Label14(22).Refresh
    
    'Si ha ido todo bien entonces numregelim=1
    If NumRegElim = 1 Then
        If OpcionListado = 315 Then
            cadSelect = "Select nomsocio, maisocio"
            cadSelect = cadSelect & " as email,tmpinformes.* from tmpinformes,rsocios where codusu = " & vUsu.Codigo & " and codsocio=codigo1"
    '        cadSelect = cadSelect & " group by codclien having email is null"
    
            '[Monica]31/01/2014: esperamos en catadau, antes de abrir la ventana del frmEmail
            If vParamAplic.Cooperativa = 0 Then

                T1 = Timer - T1
                If T1 < 3 Then
                    T1 = 3 - T1
                    espera T1
                End If
                T1 = Timer
            End If
    
            
            frmEMail.DatosEnvio = Text1(0).Text & "|" & Text1(1).Text & "|" & Abs(0) & "|" & cadSelect & "|"
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
                      
        Case 1
            vCadena = "Para adjuntar archivos debe seleccionar la carta 12. " & vbCrLf & vbCrLf & _
                      "En la descripción de la carta debe poner el Asunto. " & vbCrLf & _
                      vbCrLf
                      
            If OpcionListado = 306 Then
                      vCadena = vCadena & "Si se envia a través de SMS no adjuntará ningún archivo. " & vbCrLf & vbCrLf
                      vCadena = vCadena & "Es aconsejable enviar en formato PDF."
            End If
            
            vCadena = vCadena & "" & vbCrLf & vbCrLf
        
        Case 2
            vCadena = "Si se indica sección, seleccionaremos los socios dados de alta en esa sección. " & vbCrLf & vbCrLf & _
                      "Si indica fase, no hay que poner nada en sección, para seleccionar los" & vbCrLf & _
                      "socios que estén en la fase indicada." & vbCrLf & _
                      vbCrLf
                      
            If OpcionListado = 306 Then
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
    
    If Not DatosOK Then Exit Sub
    
    'si es listado de CARTAS/eMAIL a socios comprobar que se ha seleccionado
    'una carta para imprimir
    If OpcionListado = 306 Then
        If chkMail(1).Value = 1 Then
            ' si estamos mandando un SMS no es obligado meter un codigo de carta
        Else
            If txtCodigo(63).Text = "" Then
                MsgBox "Debe seleccionar una carta para imprimir.", vbInformation
                Exit Sub
            End If
        End If
        
        'Parametro cod. carta
        CadParam = "|pCodCarta= " & txtCodigo(63).Text & "|"
        numParam = numParam + 1
        
        'Parametro fecha
        CadParam = CadParam & "|pFecha= """ & txtCodigo(0).Text & """|"
        numParam = numParam + 1
        
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rSocioCarta.rpt" '"rComProveCarta.rpt"
        Titulo = "Cartas a Socios" '"Cartas a Proveedores"
        
        indRPT = 61 'Personalizacion de la carta a socios
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
          
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
        CadParam = "|"
    
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rSocioEtiq.rpt" '"rComProveEtiq.rpt"
        Titulo = "Etiquetas de Socios" '"Etiquetas de Proveedores"
    
        '===================================================
        '============ PARAMETROS ===========================
        indRPT = 27 'Impresion de Etiquetas de socios
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
          
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
            If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & txtCodigo(58).Text) Then Exit Sub
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
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.fecbaja} is null ") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "isnull({rsocios_seccion.fecbaja})") Then Exit Sub

    Else
        '[Monica]11/11/2013: Caso de CASTELDUC si no me dan seccion cojo los datos de la fases
        '                    si me dan la seccion funciona como el resto de cooperativas
        If ComprobarCero(txtCodigo(58).Text) <> 0 Then
            ' solo se sacan los socios que no esten dados de baja
            If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & txtCodigo(58).Text) Then Exit Sub
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
            If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.fecbaja} is null ") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "isnull({rsocios_seccion.fecbaja})") Then Exit Sub
            
        Else
            '[Monica]13/09/2016: antes un select case
            If Combo1(0).ItemData(Combo1(0).ListIndex) <> 0 Then
                ' solo se sacan los socios de la fase que sea
                If Not AnyadirAFormula(cadSelect, "{rsocios_pozos.numfases} = " & Combo1(0).ItemData(Combo1(0).ListIndex)) Then Exit Sub
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
    
    '[Monica]29/09/2014: para el caso de escalona y de utxera si marcan solo los que NO tienen fecha de revision
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        If chkMail(4).Value = 1 Then
            If Not AnyadirAFormula(cadSelect, "({rsocios.fechanac} is null or {rsocios.fechanac}='')") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "(isnull({rsocios.fechanac}) or {rsocios.fechanac}='')") Then Exit Sub
        End If
    End If



    '[Monica]08/11/2012: solo los socios que no tengan situacion de bloqueo
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Situacion = SituacionesBloqueo
        If Situacion <> "" Then
            Situacion = Mid(Situacion, 1, Len(Situacion) - 1)
            If Not AnyadirAFormula(cadFormula, "not {rsocios.codsitua} in [" & Situacion & "]") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "not rsocios.codsitua in (" & Situacion & ")") Then Exit Sub
        End If
    End If
    
    '[Monica]23/12/2011: si estamos mandando sms miramos los que tienen nro de movil
    If chkMail(1).Value = 1 Then
        If Not AnyadirAFormula(cadSelect, "not {rsocios.movsocio} is null and {rsocios.movsocio}<>''") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "not isnull({rsocios.movsocio}) and {rsocios.movsocio}<>''") Then Exit Sub
        
        ' solo saldran los que no tegan email
        If Check3.Value = 1 Then
            If Not AnyadirAFormula(cadSelect, "({rsocios.maisocio} is null or {rsocios.maisocio}='')") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "(isnull({rsocios.maisocio}) or {rsocios.maisocio}='')") Then Exit Sub
        End If
    End If
    ' si es un correo electronico miramos solo los que tienen mail
    If chkMail(0).Value = 1 Then
        If Not AnyadirAFormula(cadSelect, "not {rsocios.maisocio} is null and {rsocios.maisocio}<>''") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "not isnull({rsocios.maisocio}) and {rsocios.maisocio}<>''") Then Exit Sub
    End If
    
    '[Monica]03/03/2015: solo en el caso de que sean cartas miramnos si quieren los que no tienen correo electrónico
    '                    y son escalona o Utxera
    If chkMail(0).Value = 0 And chkMail(1).Value = 0 And chkMail(5).Value = 1 And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) Then
        If chkMail(5).Value = 1 Then
            If Not AnyadirAFormula(cadSelect, "({rsocios.maisocio} is null or {rsocios.maisocio}='')") Then Exit Sub
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
            If Not AnyadirAFormula(cadSelect, "{rsocios.correo} = 1") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.correo} = 1") Then Exit Sub
        End If
    End If
    
        
    'Parametro a la Atencion de
    If txtCodigo(62).Text <> "" Then
        CadParam = CadParam & "pAtencion=""Att. " & txtCodigo(62).Text & """|"
    Else
        CadParam = CadParam & "pAtencion=""""|"
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
        If (vParamAplic.Cooperativa = 3) And OpcionListado = 305 Then
        
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
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    Set frmMen = New frmMensajes
    
    If txtCodigo(58).Text = "" Then
        frmMen.cadWHERE = "SELECT distinct rsocios.codsocio,nomsocio,nifsocio FROM rsocios inner join rsocios_pozos on rsocios.codsocio = rsocios_pozos.codsocio where " & cadSelect & " order by rsocios.codsocio "
        frmMen.OpcionMensaje = 55 'Etiquetas socios
    Else
        '[Monica]24/10/2014: antes solo tenia cadselect
        frmMen.cadWHERE = "rsocios.codsocio in (select rsocios.codsocio from " & tabla & " where " & cadSelect & ")"
        frmMen.OpcionMensaje = 9 'Etiquetas socios
    End If
    
    frmMen.Show vbModal
    Set frmMen = Nothing
    
    If cadSelect = "" Then Exit Sub
    
    '[Monica]16/11/2012: añadida la condicion de cooperativa <> 8
    If Documento <> "" And ImpresionNormal And vParamAplic.Cooperativa <> 8 Then
        Set frmMen2 = New frmMensajes
        frmMen2.cadWHERE = " and codsocio in (select rsocios.codsocio from " & tabla & " where " & cadSelect & ")"
        frmMen2.OpcionMensaje = 15 'Etiquetas socios
        frmMen2.Show vbModal
        Set frmMen2 = Nothing
        If cadSelect = "" Then Exit Sub
    End If
    
    If OpcionListado = 306 And Me.chkMail(0).Value = 1 Then
        'Enviarlo por e-mail
        IndRptReport = indRPT
        EnviarEMailMulti cadSelect, Titulo, nomDocu, tabla ' "rSocioCarta.rpt", Tabla  'email para socios
        cmdCancel_Click (9)
    Else
        If OpcionListado = 306 And Me.chkMail(1).Value = 1 Then
            OK = True
            EnviarSMS cadSelect, Titulo, nomDocu, tabla, OK    ' "rSocioCarta.rpt", Tabla  'email para socios
            If OK Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (9)
            End If
        Else                                                                  '[Monica]22/04/2015: añadimos a mogente
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Or (vParamAplic.Cooperativa = 3 And OpcionListado = 305) Then
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
        Select Case OpcionListado
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
    FrameTipoSocio.visible = (OpcionListado = 306 And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10))
    FrameTipoSocio.Enabled = (OpcionListado = 306 And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10))
    
    '[Monica]27/06/2013: marcamos bien si quieren adjuntar archivos
    Frame1.visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) And OpcionListado = 306
    Frame1.Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) And OpcionListado = 306
    
    
    
    Select Case OpcionListado
        Case 305, 306 '305: Etiquetas de proveedor
                      '306: Cartas a proveedor
            indFrame = 9
            H = 7155 '5325
            W = 7035
            If OpcionListado = 305 Then
                H = 5325
                Me.cmdAceptarEtiqProv.Top = Me.cmdAceptarEtiqProv.Top - 2000
                Me.cmdCancel(9).Top = cmdCancel(9).Top - 2000
            End If
            PonerFrameVisible Me.FrameEtiqProv, True, H, W
            Me.Frame2.visible = (OpcionListado = 306)
            If (OpcionListado = 306) Then Me.Label9(1).Caption = "Cartas a Socios"
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
            
            If OpcionListado = 316 Then Me.FrameEnvioFacMail.Width = 5535
            
            H = FrameEnvioFacMail.Height
            W = FrameEnvioFacMail.Width
            PonerFrameVisible FrameEnvioFacMail, True, H, W
            CargarComboTipoMov 1000
            
            chkMail(3).visible = OpcionListado = 316 'Solo para facturae
            If OpcionListado = 316 Then
                cmdEnvioMail.Left = 3240
                cmdCancel(indFrame).Left = 4320
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
        
    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
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
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Or (vParamAplic.Cooperativa = 3 And OpcionListado = 305) Then
            InsertarTemporal CadenaSeleccion
            cadSelect = "rsocios_seccion.codsocio IN (" & CadenaSeleccion & ")"
        Else
            If OpcionListado = 305 Or OpcionListado = 306 Then 'Socios
                cadFormula = "{rsocios_seccion.codsocio} IN [" & CadenaSeleccion & "]"
                cadSelect = "rsocios_seccion.codsocio IN (" & CadenaSeleccion & ")"
            End If
            If vParamAplic.Cooperativa = 5 And txtCodigo(58).Text = "" Then
                cadFormula = "{rsocios.codsocio} IN [" & CadenaSeleccion & "]"
                cadSelect = "rsocios.codsocio IN (" & CadenaSeleccion & ")"
            End If
        End If
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
Dim B As Boolean
Dim TotalArray As Integer

    'En el listview3
    B = Index = 1
    For TotalArray = 0 To ListTipoMov(1000).ListCount - 1
        ListTipoMov(1000).Selected(TotalArray) = B
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
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscarOfer_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim tabla As String
Dim codcampo As String, nomCampo As String
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
        Case 0 ' Fecha de la carta
            PonerFormatoFecha txtCodigo(Index), False
            
        Case 108, 109 ' Fecha factura
            PonerFormatoFecha txtCodigo(Index), True
        
        Case 1
            PonerFormatoHora txtCodigo(Index)
        
        Case 63, 64 'CARTA de la Oferta
            EsNomCod = True
            tabla = "scartas"
            codcampo = "codcarta"
            nomCampo = "descarta"
            Formato = "000"
            Titulo = "cod. de Carta"
                    
        Case 60, 61, 110, 111 'Socio
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = txtCodigo(Index).Text
            
         Case 58, 59 'Cod. Seccion
            EsNomCod = True
            tabla = "rseccion"
            codcampo = "codsecci"
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
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomCampo, codcampo, TipCampo)
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
                
                If Index = 63 And chkMail(1).Value = 1 Then
                    txtCodigo(2).Text = DevuelveValor("select textosms from scartas where codcarta = " & DBSet(txtCodigo(63).Text, "N"))
                    If txtCodigo(2).Text = "0" Then txtCodigo(2).Text = ""
                End If
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomCampo, codcampo, TipCampo)
        End If
    End If
End Sub

Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    If txtCodigo(indD).Text <> "" Then
        cad = cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        cad = cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
    
    Documento = ""
End Sub

Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            CadParam = CadParam & AnyadirParametroDH(param, indD, indH) & """|"
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
        .Opcion = OpcionListado
        .Titulo = Titulo
        .NombreRPT = nomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub



Private Sub EnviarSMS(cadWHERE As String, cadTit As String, cadRpt As String, cadTabla As String, ByRef EstaOk As Boolean)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer
Dim Direccion As String

Dim NF As Integer
Dim cad As String
Dim B As Boolean

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
        SQL = "SELECT distinct rsocios.codsocio,nomsocio,rsocios.movsocio "
    End If
    SQL = SQL & "FROM " & cadTabla
    SQL = SQL & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cont = 0
    lista = ""
    
    B = True
    
    
    While Not Rs.EOF And B
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
            cad = ""
            Line Input #NF, cad
            Close NF

            Select Case Mid(cad, 1, 2)
                Case "OK"
                    espera 2
                
                    Me.Refresh
                    DoEvents
                    espera 0.4
                    cont = cont + 1
                    
                    SQL = "INSERT INTO rsmsenviados (codsocio, movsocio, fechaenvio, horaenvio, texto)"
                    SQL = SQL & " VALUES (" & DBSet(Rs.Fields(0), "N") & "," & DBSet(Cad1, "T") & ","
                    SQL = SQL & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(1).Text, "H") & "," & DBSet(txtCodigo(2).Text, "T") & ")"
                    conn.Execute SQL
            
            
                Case "-1"
                    MsgBox "Error en el Login, usuario o clave incorrectas", vbExclamation
                    EstaOk = False
                Case Else
                    If Mid(cad, 1, 12) = "Sin Creditos" Then
                        MsgBox "No tiene créditos. Revise", vbExclamation
                        B = False
                    Else
                        If MsgBox("Error en el envio de mensaje al socio " & DBLet(Rs.Fields(0), "N") & ". ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then B = False
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
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer
Dim Sql2 As String

    
    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    'seleccionamos todos los socios a los que queremos enviar e-mail
    SQL = "SELECT distinct " & vUsu.Codigo & ", rsocios.codsocio from rsocios where codsocio in (" & cadWHERE & ")"
    
    Sql2 = "insert into tmpinformes (codusu, codigo1) " & SQL
    conn.Execute Sql2

End Sub

Private Sub EnviarEMailMulti(cadWHERE As String, cadTit As String, cadRpt As String, cadTabla As String)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Cad1 As String, Cad2 As String, lista As String
Dim cont As Integer

On Error GoTo EEnviar

    Screen.MousePointer = vbHourglass
    
    If cadTabla = "(rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio) inner join rcampos on rsocios.codsocio = rcampos.codsocio " Or _
       cadTabla = "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio" Or _
       cadTabla = "rsocios_pozos inner join rsocios on rsocios_pozos.codsocio = rsocios.codsocio" Then
       
        'seleccionamos todos los socios a los que queremos enviar e-mail
        SQL = "SELECT distinct rsocios.codsocio,nomsocio,maisocio, maisocio "
        If ImpresionNormal Then
            SQL = SQL & "FROM " & cadTabla
        Else
            SQL = SQL & "FROM " & "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio"
        End If
    ElseIf cadTabla = "sclien" Then
        'seleccionamos todos los clientes a los que queremos enviar e-mail
        SQL = "SELECT codclien,nomclien,maiclie1,maiclie2 "
        SQL = SQL & "FROM " & cadTabla
    End If
    SQL = SQL & " WHERE " & cadWHERE
    
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
                    .OtrosParametros = CadParam
                    .NumeroParametros = numParam
                    If cadTabla = "(rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio) inner join rcampos on rsocios.codsocio = rcampos.codsocio " Or _
                       cadTabla = "rsocios_seccion inner join rsocios on rsocios_seccion.codsocio = rsocios.codsocio" Or _
                       cadTabla = "rsocios_pozos inner join rsocios on rsocios_pozos.codsocio = rsocios.codsocio" Then
                        SQL = "{rsocios.codsocio}=" & Rs.Fields(0)
                        .Opcion = 306
                    Else
                        SQL = "{sclien.codclien}=" & Rs.Fields(0)
                        .Opcion = 91
                    End If
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
                        SQL = SQL & " VALUES (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "T") & "," & DBSet(Cad1, "T") & ")"
                        conn.Execute SQL
                
                        Me.Refresh
                        DoEvents
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
            Else
                If CadenaDesdeOtroForm = "" Then
                'si se ha generado el .pdf para enviar
                    SQL = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                    SQL = SQL & " VALUES (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "T") & "," & DBSet(Cad1, "T") & ")"
                    conn.Execute SQL
            
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
            SQL = "Carta: " & txtNombre(63).Text & "|"
            
            '[Monica]08/07/2011: si no hay a la atencion no se pone nada en el cuerpo del mensaje
            '                    añadida la condicion
            If txtCodigo(62).Text <> "" Then
                SQL = SQL & "Att : " & txtCodigo(62).Text & "|"
            End If
        Else
            SQL = "Carta: " & txtNombre(63).Text & "|"
            SQL = SQL & "Att : " & txtCodigo(62).Text & "|"
        End If
       
        If Not ImpresionNormal Then
            Set frmMen3 = New frmMensajes
            frmMen3.cadWHERE = ""
            frmMen3.OpcionMensaje = 40 'archivos a seleccionar
            frmMen3.Show vbModal
            Set frmMen3 = Nothing
            If cadSelect = "" Then Exit Sub
        End If
       
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
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
    End If
End Sub


Private Sub CargarIconos()
Dim I As Integer

    Me.imgBuscarOfer(35).Picture = frmPpal.imgListImages16.ListImages(1).Picture
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

Private Function DatosOK() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean

    B = True
    
    '[Monica]11/11/2013: dejamos poner unicamente las fases de Castelduc
    If vParamAplic.Cooperativa = 5 Then
        If Combo1(0).ListIndex = -1 And ComprobarCero(txtCodigo(58).Text) = 0 Then
            B = False
            MsgBox "Debe introducir sección o fase. Revise.", vbExclamation
            PonerFocoCmb Combo1(0)
        End If
    End If
    
    '[Monica]11/11/2013: en Castelduc dejamos poner unicamente las fases
    If txtCodigo(58).Text = "" And vParamAplic.Cooperativa <> 5 Then
        MsgBox "Debe introducir un valor en la Sección. Revise.", vbExclamation
        B = False
        PonerFoco txtCodigo(58)
    End If
    
    
    If B And OpcionListado = 306 And chkMail(1).Value = 1 Then
        If txtCodigo(0).Text = "" Then
            MsgBox "Debe introducir un valor en la Fecha del SMS. Revise.", vbExclamation
            B = False
            PonerFoco txtCodigo(0)
        End If
        If B And txtCodigo(1).Text = "" Then
            MsgBox "Debe introducir un valor en la Hora del SMS. Revise.", vbExclamation
            B = False
            PonerFoco txtCodigo(1)
        End If
        If B And txtCodigo(2).Text = "" Then
            MsgBox "Debe introducir un valor en el Texto del SMS. Revise.", vbExclamation
            B = False
            PonerFoco txtCodigo(2)
        End If
    End If
    
    '[Monica]19/10/2012: comprobamos que si vamos a enviar mas de un documento es por email.
    If B And OpcionListado = 306 Then
        Documento = DevuelveDesdeBDNew(cAgro, "scartas", "documrpt", "codcarta", txtCodigo(63).Text, "N")
        If Documento <> "" Then
            If InStr(1, Documento, ",") <> 0 And chkMail(0).Value = 0 Then
                MsgBox "Para enviar más de un archivo adjunto debe seleccionar sólo por email.", vbExclamation
                B = False
                PonerFocoChk chkMail(0)
            Else
                'cualquier otro tipo de documento se tiene que poder enviar por email
                If InStr(1, Documento, ".rpt") = 0 And chkMail(0).Value = 0 Then
                    MsgBox "Para enviar más de un archivo adjunto debe seleccionar sólo por email.", vbExclamation
                    B = False
                    PonerFocoChk chkMail(0)
                End If
            End If
            
            If B And InStr(1, Documento, ".rpt") = 0 Then
                'si no es un rpt a ejecutar de la carpeta de informes comprobamos que exista la carpeta de cartas
                If Dir(App.Path & "\cartas", vbDirectory) = "" Then
                    MsgBox "No existe el directorio de cartas donde se introducen los archivos a adjuntar. Revise.", vbExclamation
                    B = False
                End If
                If B And Dir(App.Path & "\cartas\*.*", vbArchive) = "" Then
                    MsgBox "No existen archivos en el directorio cartas a adjuntar. Revise.", vbExclamation
                    B = False
                End If
            End If
        End If
    End If
    
    DatosOK = B
    
End Function


Private Function GeneracionEnvioMail(ByRef Rs As ADODB.Recordset) As Boolean
Dim Tipo As Integer
Dim TipoRec As String ' tipo de factura a la que rectifica
Dim Sql5 As String
Dim EsComplemen As Byte
Dim letraser As String

    On Error GoTo EGeneracionEnvioMail
    GeneracionEnvioMail = False
    
    cadSelect = "Select * from tmpinformes where codusu =" & vUsu.Codigo & " ORDER BY importe1,codigo1"
    Rs.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CodClien = ""
    While Not Rs.EOF
    
        InicializarVbles
        '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        
        '[Monica]28/01/2013: impresion con arrobas o sin arrobas para Montifrut
        If vParamAplic.Cooperativa = 12 Then
            If Check5.Value = 1 Then
                CadParam = CadParam & "pConArrobas=1|"
            Else
                CadParam = CadParam & "pConArrobas=0|"
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
                                CadParam = CadParam & "pDetalle=1|"
                            Else
                                CadParam = CadParam & "pDetalle=0|"
                            End If
                        Else
                            CadParam = CadParam & "pDetalle=1|"
                        End If
                        numParam = numParam + 1
                    End If
                End If
       End Select
        
       If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Function
       'Nombre fichero .rpt a Imprimir
        
        
       cadFormula = "({rfactsoc.codtipom}='" & Trim(Rs!Nombre1) & "') "
       cadFormula = cadFormula & " AND ({rfactsoc.numfactu}=" & Rs!importe1 & ") "
       cadFormula = cadFormula & " AND ({rfactsoc.fecfactu}= Date(" & Year(Rs!fecha1) & "," & Month(Rs!fecha1) & "," & Day(Rs!fecha1) & "))"

   
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
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
        
        
        If OpcionListado = 315 Then
   
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
Dim cad As String
Dim T1 As Single

    On Error GoTo eEsperaFichero
    
    T1 = Timer
    Do
        cad = Dir(App.Path & "\docum.pdf", vbArchive)
        If cad = "" Then
            If Timer - T1 > 2 Then cad = "SAL"
        End If
    Loop Until cad <> ""
    
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

Private Sub CargarComboTipoMov(Indice As Integer)
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo

'Lo cargamos con los valores de la tabla stipom que tengan tipo de documento=Albaranes (tipodocu=1)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim I As Byte

    On Error GoTo ECargaCombo
    
    '[Monica]18/02/2013: excluimos los movimientos de facturas varias
    SQL = "select codtipom, nomtipom from usuarios.stipom where (tipodocu <> 0) and tipodocu <> 12"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 0
    
    ListTipoMov(Indice).Clear
    
    'LOS TIKCETS NO LOS ENVIO POR MAIL
    While Not Rs.EOF
        ListTipoMov(Indice).AddItem Rs.Fields(0).Value & "-" & Rs.Fields(1).Value
        'ListTipoMov(indice).List (ListTipoMov(indice).NewIndex)
        ListTipoMov(Indice).Selected((ListTipoMov(Indice).NewIndex)) = True
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    
ECargaCombo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function SituacionesBloqueo() As String
Dim SQL As String
Dim cadena As String
Dim Rs As ADODB.Recordset

    cadena = ""

    SituacionesBloqueo = cadena

    SQL = "select codsitua from rsituacion where bloqueo = 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
Dim I As Integer
Dim SQL As String
Dim Rs As ADODB.Recordset

    Combo1(0).Clear
    
    SQL = "select distinct numfases from rsocios_pozos "

    Combo1(0).AddItem "Todas" 'campo del codigo
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        SQL = "Fase " & Rs.Fields(0).Value
        
        Combo1(0).AddItem SQL 'campo del codigo
        Combo1(0).ItemData(Combo1(0).NewIndex) = Rs.Fields(0).Value
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    
End Sub



