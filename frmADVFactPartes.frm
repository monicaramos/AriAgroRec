VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmADVFactPartes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7470
   Icon            =   "frmADVFactPartes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7800
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameFacturar 
      Height          =   6285
      Left            =   30
      TabIndex        =   25
      Top             =   -30
      Width           =   7395
      Begin VB.Frame FrameProgress 
         Height          =   1050
         Left            =   120
         TabIndex        =   51
         Top             =   4980
         Visible         =   0   'False
         Width           =   4605
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   345
            Left            =   120
            TabIndex        =   52
            Top             =   600
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess 
            Caption         =   "Iniciando el proceso ..."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   54
            Top             =   350
            Width           =   4335
         End
         Begin VB.Label lblProgess 
            Caption         =   "Facturando:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   135
            Width           =   4215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4065
         Left            =   120
         TabIndex        =   36
         Top             =   780
         Width           =   7125
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
            Index           =   34
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   56
            Top             =   300
            Width           =   1320
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
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
            Left            =   3255
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "Text5"
            Top             =   3480
            Width           =   3735
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
            Index           =   42
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   33
            Tag             =   "Forma Pago|N|N|0|999|scaalb|codforpa|000||"
            Top             =   3480
            Width           =   1095
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
            Index           =   41
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   32
            Tag             =   "Cod. Cliente|N|N|0|999999|scaalb|codclien|000000||"
            Top             =   3000
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
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
            Left            =   3255
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "Text5"
            Top             =   3000
            Width           =   3735
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
            Index           =   40
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   31
            Tag             =   "Cod. Cliente|N|N|0|999999|scaalb|codclien|000000||"
            Top             =   2595
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
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
            Left            =   3255
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "Text5"
            Top             =   2595
            Width           =   3735
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
            Index           =   38
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   29
            Top             =   1740
            Width           =   1320
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
            Index           =   39
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   30
            Top             =   2160
            Width           =   1320
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
            Index           =   36
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   27
            Top             =   765
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
            Index           =   37
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   28
            Top             =   1170
            Width           =   1350
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   14
            Left            =   1860
            Picture         =   "frmADVFactPartes.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label10 
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
            Index           =   5
            Left            =   240
            TabIndex        =   57
            Top             =   270
            Width           =   1440
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   22
            Left            =   1860
            ToolTipText     =   "Buscar forma pago"
            Top             =   3480
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Forma pago"
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
            Left            =   240
            TabIndex        =   49
            Top             =   3450
            Width           =   1155
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   20
            Left            =   1860
            ToolTipText     =   "Buscar socio"
            Top             =   2595
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   21
            Left            =   1860
            ToolTipText     =   "Buscar socio"
            Top             =   3000
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Index           =   50
            Left            =   1155
            TabIndex        =   47
            Top             =   3000
            Width           =   705
         End
         Begin VB.Label Label4 
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
            Index           =   51
            Left            =   1155
            TabIndex        =   46
            Top             =   2595
            Width           =   735
         End
         Begin VB.Label Label10 
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
            Index           =   2
            Left            =   240
            TabIndex        =   45
            Top             =   2445
            Width           =   540
         End
         Begin VB.Label Label4 
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
            Index           =   37
            Left            =   1170
            TabIndex        =   42
            Top             =   2160
            Width           =   615
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   12
            Left            =   1860
            Picture         =   "frmADVFactPartes.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   1755
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Parte"
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
            TabIndex        =   41
            Top             =   1485
            Width           =   1200
         End
         Begin VB.Label Label4 
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
            Index           =   46
            Left            =   1170
            TabIndex        =   40
            Top             =   1740
            Width           =   645
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   13
            Left            =   1860
            Picture         =   "frmADVFactPartes.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   2175
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Index           =   36
            Left            =   1200
            TabIndex        =   39
            Top             =   1170
            Width           =   570
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nº Parte"
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
            Left            =   240
            TabIndex        =   38
            Top             =   600
            Width           =   840
         End
         Begin VB.Label Label4 
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
            Index           =   45
            Left            =   1200
            TabIndex        =   37
            Top             =   765
            Width           =   600
         End
      End
      Begin VB.CommandButton cmdAceptarFac 
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
         Left            =   5085
         TabIndex        =   34
         Top             =   5670
         Width           =   1030
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   6210
         TabIndex        =   35
         Top             =   5670
         Width           =   1030
      End
      Begin VB.Label Label10 
         Caption         =   "Facturación de Partes ADV"
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
         Left            =   315
         TabIndex        =   26
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label10 
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
         Index           =   10
         Left            =   120
         TabIndex        =   55
         Top             =   3360
         Width           =   6855
      End
   End
   Begin VB.Frame FramePreFacturar 
      Height          =   6225
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7035
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "Text5"
         Top             =   4185
         Width           =   3615
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
         Index           =   1
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   10
         Top             =   4185
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Text5"
         Top             =   3780
         Width           =   3615
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
         Index           =   0
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   9
         Top             =   3780
         Width           =   1215
      End
      Begin VB.CheckBox chkSoloFacturar 
         Caption         =   "Solo Partes para facturar"
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
         Left            =   540
         TabIndex        =   13
         Top             =   5460
         Value           =   1  'Checked
         Width           =   3405
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2190
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarPreFac 
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
         Left            =   4410
         TabIndex        =   14
         Top             =   5490
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   5580
         TabIndex        =   16
         Top             =   5490
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2190
         Width           =   1215
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
         Index           =   30
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1155
         Width           =   1215
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
         Index           =   31
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
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
         Index           =   29
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   8
         Top             =   3210
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   3210
         Width           =   3615
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
         Index           =   28
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   7
         Top             =   2805
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text5"
         Top             =   2805
         Width           =   3615
      End
      Begin VB.Frame Frame1 
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
         Height          =   795
         Left            =   360
         TabIndex        =   64
         Top             =   4560
         Width           =   6255
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Por Socio/Artículo"
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
            Index           =   4
            Left            =   3585
            TabIndex        =   67
            Top             =   330
            Width           =   2355
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Previsión"
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
            Left            =   2175
            TabIndex        =   66
            Top             =   330
            Width           =   1575
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Tipo de Venta"
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
            Left            =   240
            TabIndex        =   65
            Top             =   330
            Width           =   1785
         End
      End
      Begin VB.Frame Frame7 
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
         Height          =   735
         Left            =   360
         TabIndex        =   50
         Top             =   4560
         Width           =   3135
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Trabajadores"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   12
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Artículos"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   11
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   1
         Left            =   1380
         Top             =   4185
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
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
         Left            =   420
         TabIndex        =   63
         Top             =   3480
         Width           =   1065
      End
      Begin VB.Label Label4 
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
         Left            =   675
         TabIndex        =   62
         Top             =   3810
         Width           =   690
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   0
         Left            =   1380
         Top             =   3810
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Left            =   675
         TabIndex        =   61
         Top             =   4215
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nº Parte"
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
         Left            =   420
         TabIndex        =   58
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label4 
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
         Index           =   44
         Left            =   3060
         TabIndex        =   24
         Top             =   2190
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1380
         Picture         =   "frmADVFactPartes.frx":01AD
         Top             =   2190
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Prefacturación de Albaranes"
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
         Left            =   420
         TabIndex        =   23
         Top             =   300
         Width           =   6165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Parte"
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
         Left            =   420
         TabIndex        =   22
         Top             =   1860
         Width           =   1200
      End
      Begin VB.Label Label4 
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
         Index           =   42
         Left            =   675
         TabIndex        =   21
         Top             =   2190
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   3540
         Picture         =   "frmADVFactPartes.frx":0238
         Top             =   2190
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   40
         Left            =   675
         TabIndex        =   20
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
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
         Index           =   39
         Left            =   675
         TabIndex        =   19
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label Label4 
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
         Index           =   35
         Left            =   675
         TabIndex        =   18
         Top             =   3210
         Width           =   705
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   15
         Left            =   1380
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   34
         Left            =   675
         TabIndex        =   17
         Top             =   2850
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
         Index           =   33
         Left            =   420
         TabIndex        =   15
         Top             =   2520
         Width           =   540
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   14
         Left            =   1380
         Top             =   2805
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmADVFactPartes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer

' 0 = prevision de facturacion
' 1 = facturacion
      
      
'Alguna opcion mas
'                   1000.-  Es cuando paso pedido a albaran y este a factura en el mismo proceso
'                   1001.-  Facturar un unico albaran
      
      
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir


Public CodClien As String 'Para seleccionar inicialmente las ofertas del Proveedor

'#Laura 14/11/2006 Recuperar facturas Alzira
Public EstaRecupFact As Boolean ' si esta recuperando facturas (para albaranes de mostrador)


'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

'Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Private WithEvents frmSoc As frmManSocios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmArt As frmADVArticulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta
Attribute frmFPa.VB_VarHelpID = -1
'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmTto As frmADVTrataMoi 'Tipo de venta
Attribute frmTto.VB_VarHelpID = -1
Private WithEvents frmTrab As frmManTraba 'Trabajadores
Attribute frmTrab.VB_VarHelpID = -1

'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------


Dim indCodigo As Integer 'indice para txtCodigo

Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub





Private Sub chkSoloFacturar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub cmdAceptarFac_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
'Facturacion de Albaranes
Dim campo As String, cad As String, Cad1 As String, Cad2 As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
Dim CambiamosConta As Boolean

Dim CadParam As String

    
    InicializarVbles
    cadFrom = ""
    CambiamosConta = False
    '--- Comprobar q los campos tienen valor
    If Trim(txtCodigo(34).Text) = "" Then 'Fecha factura
        MsgBox "El campo Fecha Factura debe tener valor.", vbExclamation
        PonerFoco txtCodigo(34)
        Exit Sub
    End If
    If Trim(txtCodigo(42).Text) = "" Then 'la forma de pago debe tener un valor
        MsgBox "El campo Forma de Pago debe tener un valor.", vbExclamation
        PonerFoco txtCodigo(42)
        Exit Sub
    End If
   
    
    '[Monica]20/06/2017: control de fechas que antes no estaba
    ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(34)))
    If ResultadoFechaContaOK > 0 Then
        If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
        Exit Sub
    End If
    
    
    '--- Seleccinar los Partes que cumplen los criterios introducidos
    'Desde/Hasta Nº PARTE
    '-------------------------
    cDesde = Trim(txtCodigo(36).Text)
    cHasta = Trim(txtCodigo(37).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advpartes.numparte}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDParte= """) Then Exit Sub
    End If

    'Desde/Hasta FECHA del PARTE
    '--------------------------------------------
    cDesde = Trim(txtCodigo(38).Text)
    cHasta = Trim(txtCodigo(39).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advpartes.fechapar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDFecha= """) Then Exit Sub
    End If

    'Cadena para seleccion D/H SOCIO
    '----------------------------------------
    cDesde = Trim(txtCodigo(40).Text)
    cHasta = Trim(txtCodigo(41).Text)
    nDesde = Trim(txtNombre(40).Text)
    nHasta = Trim(txtNombre(41).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advpartes.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pSocio= """) Then Exit Sub
    End If
    
    cadSQL = cadSelect
    'Seleccionar los Albaranes que tiene scaalb.factursn=1
    cad = " {advpartes.factursn=1} " 'and {advpartes.litrosrea<>0} "
    If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
    AnyadirAFormula cadFormula, cad
    
    '[Monica]21/03/2011 he quitado el not in de esta primera condicion
    Cad1 = " (( advpartes_lineas.codartic in (select codartic from advartic where advartic.tipoprod = 0) and "
    Cad1 = Cad1 & " advpartes.litrosrea <> 0) or "
    Cad1 = Cad1 & "(not exists (select l.codartic from advpartes_lineas l inner join advartic a on l.codartic=a.codartic  where a.tipoprod=0 and advpartes.numparte=l.numparte) and "
    Cad1 = Cad1 & " advpartes.litrosrea = 0 ))   "
    
'    cad2 = "( advpartes_lineas.codartic not in (select codartic from advartic where advartic.tipoprod = 0) and advpartes.litrosrea = 0 )"
    Cad2 = "((advpartes_lineas.codartic in (select codartic from advartic where advartic.tipoprod = 0)) and "
    Cad2 = Cad2 & " advpartes.litrosrea = 0 )"
    
    If Not AnyadirAFormula(cadSelect, Cad1) Then Exit Sub
    AnyadirAFormula cadFormula, Cad1
    
    
    '--- Comprobar q se han Seleccionados registros de Albaran con esos criterios
    cad = "Select count(*) " ' & NomTabla & " INNER JOIN " & nomTablaLin
    If cadFrom = "" Then cadFrom = " (advpartes inner join advpartes_lineas on advpartes.numparte = advpartes_lineas.numparte) inner join rsocios on advpartes.codsocio = rsocios.codsocio "
    cad = cad & " FROM " & cadFrom

    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    'Verificar si con los criterios seleccionados (PARA VENTAS)
    'seleccionar si quedan en el rango Albaranes que no se van a Facturar
    'y mostrar mensaje
    
    'Seleccionar los Albaranes que tiene scaalb.factursn=0
    campo = " (advpartes.factursn=0 or " & Cad2 & ")"  ' advpartes.litrosrea=0) "
    If Not AnyadirAFormula(cadSQL, campo) Then Exit Sub
    cadSQL = cad & " WHERE " & cadSQL
    If RegistrosAListar(cadSQL) > 0 Then
        'Mostrar los Albaranes que no se van a Facturar
        cadSQL = Replace(cadSQL, "count(*)", "advpartes.numparte,advpartes.fechapar,advpartes.codtrata, advpartes.codsocio,rsocios.nomsocio,advpartes.codcampo")
        frmMensajes.OpcionMensaje = 12
        frmMensajes.cadWHERE = cadSQL
        frmMensajes.Show vbModal
        If frmMensajes.vCampos = "0" Then Exit Sub
    End If
    
    cad = cad & " WHERE " & cadSelect
    'Pasar Albaranes a Facturas
    If InStr(cad, "rsocios") <> 0 Then 'hay JOIN con rsocios
        cad = Replace(cad, "count(*)", "*")
    Else
        cad = Replace(cad, "count(*)", "*")
    End If

    '[Monica]17/03/2011
    CadParam = cad
    If Not EstaParametrizado(CadParam) Then
        Exit Sub
    End If

    Me.Height = Me.Height + 300
    Me.FrameFacturar.Height = Me.FrameFacturar.Height + 300
    Me.FrameProgress.visible = True
'--monica
'    Me.FrameProgress.Top = 6250
    Me.ProgressBar1.Left = 200
    Me.ProgressBar1.Value = 0
    Me.lblProgess(1).Caption = "Inicializando el proceso..."
        
    'proceso normal
    Screen.MousePointer = vbHourglass
     
    '------------------------------------------------------------------------------
    '  LOG de acciones.
    Set LOG = New cLOG
    campo = "Facturacion de Partes: " & cadSelect
    LOG.Insertar 2, vUsu, campo
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    campo = "" ' txtCSB(0).Text & "|" & txtCSB(1).Text & "|" & txtCSB(2).Text & "|"
    TraspasoPartesFacturas cad, cadSelect, txtCodigo(34).Text, "", Me.ProgressBar1, Me.lblProgess(1), True, CodClien, campo, txtCodigo(42).Text

    Screen.MousePointer = vbDefault
    
    If CambiamosConta Then
       'Reestablecer la conexion con la antigua conta
'       AbrirConexionConta False
    End If
    Me.Height = Me.Height - 300
    Me.FrameFacturar.Height = Me.FrameFacturar.Height - 300
    Me.FrameProgress.visible = False
End Sub

Private Function EstaParametrizado(cad As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
    
    cad = Replace(cad, "*", "distinct advpartes.codsocio")
    SQL = "select count(*)  from rsocios where esfactadvinterna = 1 and codsocio in (" & cad & ")"
    
    EstaParametrizado = True
    
    If TotalRegistros(SQL) > 0 Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.SeccionADV) Then
            If vSeccion.AbrirConta Then
                ' codigo de iva de facturas internas de adv
                SQL = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                
                If SQL = "" Then
                    MsgBox "No está parametrizado el código de iva de socios con facturación interna o no existe en contabilidad. Revise.", vbExclamation
                    EstaParametrizado = False
                    Set vSeccion = Nothing
                    Exit Function
                End If
            End If
        Else
            MsgBox "No está parametrizada la sección de adv en parámetros. Revise.", vbExclamation
            EstaParametrizado = False
            Set vSeccion = Nothing
            Exit Function
        End If
        Set vSeccion = Nothing
    End If
End Function


'#### Laura 14/11/2006 Recuperar facturas ALZIRA
Private Function ComprobarCliente_RecuperarFac(cadSelAlb As String, FecFac As String, numFac As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim codMacta1 As String 'cliente factura ariges
Dim codMacta2 As String 'cliente factura conta
Dim LEtra As String

    On Error GoTo ErrCompCliente
    ComprobarCliente_RecuperarFac = False
    
    'codmacta del cliente del albaran a facturar en Ariges
    SQL = "select scaalb.codclien,sclien.codmacta"
    SQL = SQL & " from scaalb inner join sclien on scaalb.codclien=sclien.codclien "
    SQL = SQL & " Where " & cadSelAlb
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        codMacta1 = DBLet(Rs!Codmacta, "T")
    
    End If
    Set Rs = Nothing
    
    
    'codmacta en la contabilidad
    LEtra = ObtenerLetraSerie("FAV")
    SQL = "SELECT codmacta FROM cabfact "
    SQL = SQL & " WHERE numserie=" & DBSet(LEtra, "T") & " AND codfaccl=" & numFac & " AND anofaccl=" & Year(FecFac)
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        codMacta2 = DBLet(Rs!Codmacta, "T")
    End If
    Set Rs = Nothing
    
    If codMacta1 <> "" And codMacta2 <> "" Then
        If codMacta1 = codMacta2 Then
            ComprobarCliente_RecuperarFac = True
        Else
            ComprobarCliente_RecuperarFac = False
            MsgBox "La cuenta contable en la factura de Contabilidad no coincide con la del cliente del Albaran", vbExclamation
        End If
    Else
        ComprobarCliente_RecuperarFac = False
        MsgBox "No se ha podido leer la cuenta contable del cliente", vbExclamation
    End If
    
    Exit Function
    
ErrCompCliente:
    ComprobarCliente_RecuperarFac = False
    MuestraError Err.Number, "Comprobar cliente", Err.Description
End Function
'#####

Private Sub cmdAceptarPreFac_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
'Prevision de Facturacion de Albaranes
Dim campo As String, cad As String
Dim B As Boolean
Dim Indice As Integer
Dim Cad1 As String
Dim Cad2 As String
Dim cadTabla As String


    InicializarVbles
        
    'Pasar nombre de la Empresa como parametro
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    'Desde/Hasta NRO de PARTE
    '--------------------------------------------
    cDesde = Trim(txtCodigo(30).Text)
    cHasta = Trim(txtCodigo(31).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advpartes.numparte}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHParte= """) Then Exit Sub
    End If
    
    'Desde/Hasta FECHA del ALBARAN
    '--------------------------------------------
    cDesde = Trim(txtCodigo(26).Text)
    cHasta = Trim(txtCodigo(27).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advpartes.fechapar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= """) Then Exit Sub
    End If

    If OptDetalle(0).Value Then
        'Cadena para seleccion SOCIO
        '--------------------------------------------
        cDesde = Trim(txtCodigo(28).Text)
        cHasta = Trim(txtCodigo(29).Text)
        nDesde = Trim(txtNombre(28).Text)
        nHasta = Trim(txtNombre(29).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{advpartes.codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
        End If
    End If
    
    If OptDetalle(1).Value Then
        'Cadena para seleccion TRABAJADOR
        '--------------------------------------------
        cDesde = Trim(txtCodigo(0).Text)
        cHasta = Trim(txtCodigo(1).Text)
        nDesde = Trim(txtNombre(0).Text)
        nHasta = Trim(txtNombre(1).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{advpartes_trabajador.codtraba}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador= """) Then Exit Sub
        End If
    End If
    
    If vParamAplic.Cooperativa = 3 Then
        'Cadena para seleccion Tipo de venta
        '--------------------------------------------
        cDesde = Trim(txtCodigo(0).Text)
        cHasta = Trim(txtCodigo(1).Text)
        nDesde = Trim(txtNombre(0).Text)
        nHasta = Trim(txtNombre(1).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{advpartes.codtrata}"
            TipCod = "T"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrata= """) Then Exit Sub
        End If
    
    End If
    
    
    If Me.OptDetalle(0).Value Then
        'Seleccionar los que esten marcados para facturar
        'Seleccionar solo aquellos que el campo advpartes.factursn=1
        If Me.chkSoloFacturar.Value = 1 Then
            cad = " {advpartes.factursn}=1 "
            If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
        End If
        
        '[Monica]21/03/2011 he quitado el not in de esta primera condicion
        Cad1 = " (( advpartes_lineas.codartic in (select codartic from advartic where advartic.tipoprod = 0) and "
        Cad1 = Cad1 & " advpartes.litrosrea <> 0) or "
        Cad1 = Cad1 & "(not exists (select l.codartic from advpartes_lineas l inner join advartic a on l.codartic=a.codartic  where a.tipoprod=0 and advpartes.numparte=l.numparte) and "
        Cad1 = Cad1 & " advpartes.litrosrea = 0 ))   "
        
    '    cad2 = "( advpartes_lineas.codartic not in (select codartic from advartic where advartic.tipoprod = 0) and advpartes.litrosrea = 0 )"
        Cad2 = "((advpartes_lineas.codartic in (select codartic from advartic where advartic.tipoprod = 0)) and "
        Cad2 = Cad2 & " advpartes.litrosrea = 0 )"
    
        If Not AnyadirAFormula(cadSelect, Cad1) Then Exit Sub
        AnyadirAFormula cadFormula, Cad1
    
        cadTabla = "advpartes INNER JOIN advpartes_lineas ON advpartes.numparte = advpartes_lineas.numparte"
    Else
        cadTabla = "advpartes INNER JOIN advpartes_trabajador ON advpartes.numparte = advpartes_trabajador.numparte"
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If ProcesoCargaIntermedia(cadTabla, cadSelect, (OptDetalle(0).Value)) Then
        If HayRegParaInforme("tmpinformes", "codusu = " & vUsu.Codigo) Then
            If OptDetalle(0).Value Then
                '[Monica]18/05/2012
                If vParamAplic.Cooperativa = 3 Then
                    If OptDetalle(2).Value Then
                        Titulo = "Ventas por Destinos"
                        CadParam = CadParam & "pParte=0|"
                    End If
                    '[Monica]10/07/2013: antes estaba en else, he añadido nueva opcion para otro tipo de listado
                    If OptDetalle(3).Value Then
                        Titulo = "Previsión Facturación Albaranes"
                        CadParam = CadParam & "pParte=0|"
                    End If
                    If OptDetalle(4).Value Then
                        Titulo = "Ventas por Socio/Artículo"
                        CadParam = CadParam & "pParte=0|"
                    End If
                Else
                    Titulo = "Previsión Facturación Partes de ADV"
                    CadParam = CadParam & "pParte=1|"
                End If
                numParam = numParam + 1
                
                conSubRPT = True
                nomRPT = "rADVPrevFactParte.rpt"
                
                '[Monica]18/05/2012
                If vParamAplic.Cooperativa = 3 And OptDetalle(2).Value Then
                    nomRPT = "rADVPrevFactTto.rpt"
                End If
                '[Monica]10/07/2013: nuevo informe por socios/articulos
                If vParamAplic.Cooperativa = 3 And OptDetalle(4).Value Then
                    nomRPT = "rADVPrevFactSocArt.rpt"
                End If
                
            End If
            
            If OptDetalle(1).Value Then
                Titulo = "Previsión Trabajadores Partes de ADV"
            
                conSubRPT = True
                nomRPT = "rADVPrevFactTrab.rpt"
            End If
       
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            cad = "pTitulo=""" & Titulo & """|"
            CadParam = CadParam & cad
            numParam = numParam + 1
        
            LlamarImprimir
        End If
    End If
    
EPreFact:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Informe Prefacturación", Err.Description
    End If
End Sub



Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
     
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 0 '0: Prevision de Facturacion Partes (NO IMPRIME LISTADO)
                PonerFoco txtCodigo(30)
            Case 1 '1: Facturacion de Partes
                PonerFoco txtCodigo(34)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim I As Integer
Dim indFrame As Single

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    Me.FramePreFacturar.visible = False
    Me.FrameFacturar.visible = False
    
    
    For I = 0 To 1
        Me.imgBuscarOfer(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 14 To 15
        Me.imgBuscarOfer(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 20 To 22
        Me.imgBuscarOfer(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I

    CommitConexion
    
    ' Necesitamos la conexion a la contabilidad de la seccion de adv
    ' para sacar los porcentajes de iva de los articulos y calcular
    ' los datos de la factura
    ConexionConta
    
    
    Me.FramePreFacturar.visible = False
    Me.FrameFacturar.visible = False
    
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
        Case 0 '0: Prevision Facturacion de Partes (NO IMPRIME LISTADO)
            PonerFramePreFacVisible True, H, W
            indFrame = 5 'solo para el boton cancelar
        
            NomTabla = "advpartes"
            NomTablaLin = "advpartes_lineas"
            
            '[Monica]18/05/2012
            If vParamAplic.Cooperativa = 3 Then
                Label10(6).Caption = "Albarán"
                Label4(43).Caption = "Fecha Albarán"
                Frame7.visible = False
                Frame7.Enabled = False
                Frame1.visible = True
                Frame1.Enabled = True
'                Label4(0).visible = False
'                Label4(1).visible = False
'                Label4(2).visible = False
'                imgBuscarOfer(0).visible = False
'                imgBuscarOfer(0).Enabled = False
'                imgBuscarOfer(1).visible = False
'                imgBuscarOfer(1).Enabled = False
'                txtCodigo(0).visible = False
'                txtCodigo(0).Enabled = False
'                txtCodigo(1).visible = False
'                txtCodigo(1).Enabled = False
'                txtNombre(0).visible = False
'                txtNombre(0).Enabled = False
'                txtNombre(1).visible = False
'                txtNombre(1).Enabled = False

                Label4(2).Caption = "Tipo de Venta"
                imgBuscarOfer(0).ToolTipText = "Tipo de venta"
                imgBuscarOfer(1).ToolTipText = "Tipo de venta"
                

'                chkSoloFacturar.Left = 550
                chkSoloFacturar.Caption = "Solo Albaranes para facturar"
                chkSoloFacturar.Width = 3115
                Me.OptDetalle(2).Value = True
                OptDetalle_KeyDown 2, 1, 0
            Else
                Frame1.visible = False
                Frame1.Enabled = False
            End If
        
        
        Case 1 '1: Facturacion de Partes
            PonerFrameFacVisible True, H, W
            txtCodigo(34).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(39).Text = Format(CDate(txtCodigo(34).Text) - 1, "dd/mm/yyyy")
            indFrame = 6
            
            NomTabla = "advpartes"
            NomTablaLin = "advpartes_lineas"
            
            '[Monica]18/05/2012
            If vParamAplic.Cooperativa = 3 Then
                Label10(0).Caption = " Facturación de Albaranes"
                Label10(4).Caption = "Albarán"
                Label10(1).Caption = "Fecha Albarán"
            End If
            
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Pago
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTrab_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de trabajadores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTto_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de tipos de venta
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0, 1 ' trabajadores o tipo de venta
            indCodigo = Index
            If vParamAplic.Cooperativa = 3 Then
                'tipo de venta
                Set frmTto = New frmADVTrataMoi
                frmTto.DatosADevolverBusqueda = "0|1|"
                frmTto.Show vbModal
                Set frmTto = Nothing
            Else
                'trabajadores
                Set frmTrab = New frmManTraba
                frmTrab.DatosADevolverBusqueda = "0|2|"
                If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
                frmTrab.Show vbModal
                Set frmTrab = Nothing
            End If
            
        Case 14, 15, 20, 21 'Cod. Socio
            Select Case Index
                Case 11, 12: indCodigo = Index + 9
                Case 14, 15: indCodigo = Index + 14
                Case 20, 21: indCodigo = Index + 20
                Case 27, 28: indCodigo = Index + 21
                Case 32: indCodigo = 8
            End Select
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|2|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            
        Case 16, 17, 23  'Forma de PAGO

        
        
            Select Case Index
                Case 16, 17: indCodigo = Index + 14
                Case 22, 23: indCodigo = Index + 20
                Case 29, 30: indCodigo = Index + 21
            End Select
            
            Set frmFPa = New frmComercial
            
            AyudaFPagoCom frmFPa, txtCodigo(42).Text
            
            Set frmFPa = Nothing
            
       Case 22 'Forma de Pago de contabilidad
            indCodigo = Index + 20
            AbrirFrmForpaConta (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   
'++monica

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
        Case 10 'FramePreFacturar
            indCodigo = 26
        Case 11 'FramePreFacturar
            indCodigo = 27
        Case 12 'Frame Factura
            indCodigo = 38
        Case 13 'Frame Factura
            indCodigo = 39
        Case 14 'FrameFactura
            indCodigo = 34
   
   End Select
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)


'++
'
'
'
'
'
'   Screen.MousePointer = vbHourglass
'   Set frmF = New frmCal
'   frmF.Fecha = Now
'
'
'   Select Case Index
'        Case 10 'FramePreFacturar
'            indCodigo = 26
'        Case 11 'FramePreFacturar
'            indCodigo = 27
'        Case 12 'Frame Factura
'            indCodigo = 38
'        Case 13 'Frame Factura
'            indCodigo = 39
'        Case 14 'FrameFactura
'            indCodigo = 34
'   End Select
'
'   PonerFormatoFecha txtCodigo(indCodigo)
'   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)
'
'   Screen.MousePointer = vbDefault
'   frmF.Show vbModal
'   Set frmF = Nothing
'   PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub OptTipoInf_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptDetalle_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 1 Then
        chkSoloFacturar.Enabled = False
        chkSoloFacturar.Value = 0
        Label9(0).Caption = "Previsión Trabajadores Partes ADV"
    End If
    If Index = 0 Then
        chkSoloFacturar.Enabled = True
        Label9(0).Caption = "Previsión Facturación Partes ADV"
    End If
    '[Monica]18/05/2012
    If Index = 2 Then
        Label9(0).Caption = "Informe por Tipo de Venta"
    End If
    If Index = 3 Then
        Label9(0).Caption = "Previsión Facturación Albaranes"
    End If
End Sub

Private Sub OptDetalle_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptDetalle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 Then
        chkSoloFacturar.Enabled = False
        chkSoloFacturar.Value = 0
        Label9(0).Caption = "Previsión Trabajadores Partes ADV"
    End If
    If Index = 0 Then
        chkSoloFacturar.Enabled = True
        Label9(0).Caption = "Previsión Facturación Partes ADV"
    End If
    '[Monica]18/05/2012
    If Index = 2 Then
        Label9(0).Caption = "Informe por Tipo de Venta"
    End If
    If Index = 3 Then
        Label9(0).Caption = "Previsión Facturación Albaranes"
    End If
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim devuelve As String
Dim codcampo As String, nomCampo As String
Dim tabla As String
      
    Select Case Index
        Case 0, 1 'Codigo de trabajador
            If vParamAplic.Cooperativa <> 3 Then
                If PonerFormatoEntero(txtCodigo(Index)) Then
                    nomCampo = "nomtraba"
                    tabla = "straba"
                    codcampo = "codtraba"
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomCampo, codcampo, "N")
                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
                Else
                    txtNombre(Index).Text = ""
                End If
            Else
                nomCampo = "nomtrata"
                tabla = "advtrata"
                codcampo = "codtrata"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomCampo, codcampo, "T")
            End If
        
        'FECHA Desde Hasta
        Case 26, 27, 34, 38, 39
            PonerFormatoFecha txtCodigo(Index)
        
        Case 30, 31, 36, 37 'Nº de Parte
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If

        Case 28, 29, 40, 41 'Cod. Socio
            If PonerFormatoEntero(txtCodigo(Index)) Then
                nomCampo = "nomsocio"
                tabla = "rsocios"
                codcampo = "codsocio"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomCampo, codcampo, "N")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 42  'Cod. Formas de PAGO de comercial
            If PonerFormatoEntero(txtCodigo(Index)) Then
'                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "forpago", "nomforpa", "codforpa", "N")
'[Monica] 09/02/2010 no es de comercial sino de la contabilidad de adv
                If vParamAplic.ContabilidadNueva Then
                    txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(Index), "N")
                Else
                    txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(Index), "N")
                End If
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
        
    End Select
End Sub




Private Sub PonerFramePreFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame del Prevision Facturacion Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim B As Boolean
Dim cad As String

    H = 6105
'    If OpcionListado = 1 Then 'Inf. Incum. plazos entrega
'        h = 5300
'        Me.cmdAceptarPreFac.Top = 4600
'        Me.cmdCancel(5).Top = Me.cmdAceptarPreFac.Top
'    End If
    W = 7040
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    PonerFrameVisible Me.FramePreFacturar, visible, H, W
    If visible = True Then
        B = (OpcionListado = 0)
        Me.imgBuscarOfer(14).visible = B
        Me.imgBuscarOfer(15).visible = B
        Me.txtCodigo(30).visible = B
        Me.txtCodigo(31).visible = B
        'solo albaranes a facturar
        Me.chkSoloFacturar.visible = B
        Me.chkSoloFacturar.Value = 1
        
        'Detalle o resumen
        Me.Frame7.visible = B 'And CodClien = "ALV"
        Me.OptDetalle(0).Value = True
        
        If Not B Then
            Me.Label9(0).Caption = "Incum. plazos entrega"
        Else 'Prevision Facturacion
            Me.Label9(0).Caption = "Previsión facturación de Partes ADV"
        End If
    End If
End Sub


Private Sub PonerFrameFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim cad As String

    H = 6285
    W = 7395
    
    If visible = True Then
         Select Case CodClien 'aqui guardamos el tipo de movimiento
            Case "PAR": cad = "(ADV)"
                
        End Select
        
        Me.Label10(0).Caption = "Factura de Partes ADV " & cad
        Me.Caption = "Facturación"
    End If
    
    PonerFrameVisible Me.FrameFacturar, visible, H, W
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

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = Titulo
        .ConSubInforme = conSubRPT
        .NombreRPT = nomRPT  'nombre del informe
        .Show vbModal
    End With
End Sub

Private Sub txtCodigo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
           Case 15, 16 'ARTICULO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sartic", "nomartic", "codartic", "Articulo", "T")
            If txtNombre(Index).Text = "" And txtCodigo(Index) <> "" Then Cancel = True
    End Select
End Sub

Private Function ObtenerClientes(cadW As String, Importe As String) As String
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error GoTo EClientes
    
    cadW = Replace(cadW, "{", "")
    cadW = Replace(cadW, "}", "")
    
    SQL = "select codclien,nomclien,sum(baseimp1),sum(baseimp2),sum(baseimp3),sum(baseimp1)+ sum(if(isnull(baseimp2),0,baseimp2))+ sum(if(isnull(baseimp3),0,baseimp3)) as BaseImp"
    SQL = SQL & " From scafac "
    If cadW <> "" Then SQL = SQL & " where " & cadW
    SQL = SQL & " group by codclien "
    If Importe <> "" Then SQL = SQL & "having baseimp>" & Importe
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not Rs.EOF
'        If RS!BaseImp >= CCur(Importe) Then
            SQL = SQL & Rs!CodClien & ","
'        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    If SQL <> "" Then
        SQL = Mid(SQL, 1, Len(SQL) - 1)
        SQL = "( {scafac.codclien} IN [" & SQL & "] )"
    End If
    ObtenerClientes = SQL
    
EClientes:
   If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
End Function

Private Sub AbrirFrmForpaConta(Indice As Integer)
'    indCodigo = indice
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtCodigo(indCodigo)
'    frmFpa.Conexion = cContaFacSoc
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub


Private Sub ConexionConta()
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.SeccionADV) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.SeccionADV) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub


Private Function ProcesoCargaIntermedia(cTabla As String, cWhere As String, Partes As Boolean) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoCargaHoras
    
    Screen.MousePointer = vbHourglass
    
    ProcesoCargaIntermedia = False

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
    End If
    
    
    If Partes Then
        SQL = "select distinct " & vUsu.Codigo & ", advpartes.numparte from " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            SQL = SQL & " WHERE" & cWhere
        End If
        
        Sql3 = "insert into tmpinformes (codusu,importe1) " & SQL
        conn.Execute Sql3
    Else
        SQL = "Select advpartes_trabajador.codtraba, advpartes.fechapar, advpartes.numparte, 0 as tipo, sum(advpartes_trabajador.horas) horas, sum(advpartes_trabajador.importel) importe FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            SQL = SQL & " WHERE " & cWhere
        End If
        SQL = SQL & " group by 1, 2, 3, 4"
        SQL = SQL & " union "
        SQL = SQL & " Select advfacturas_trabajador.codtraba, advfacturas_partes.fechapar, advfacturas_partes.numparte, 1 as tipo, sum(advfacturas_trabajador.horas) horas, sum(advfacturas_trabajador.importel) importe  "
        SQL = SQL & " from " & Replace(Replace(QuitarCaracterACadena(cTabla, "_1"), "advpartes_trabajador", "advfacturas_trabajador"), "advpartes", "advfacturas_partes")
        If cWhere <> "" Then
            SQL = SQL & " WHERE " & Replace(Replace(cWhere, "advpartes_trabajador", "advfacturas_trabajador"), "advpartes", "advfacturas_partes")
        End If
        SQL = SQL & " group by 1, 2, 3, 4"
        SQL = SQL & " order by 1, 2, 3, 4"
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                                                'horas,  numparte, tipo:0=parte/1=factura, Importe
        SQL = "insert into tmpinformes (codusu, codigo1, fecha1, importe1, importe2, importe3, importe4) values "
            
        While Not Rs.EOF
            Sql2 = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
            Sql2 = Sql2 & " and codigo1 = " & DBSet(Rs.Fields(0).Value, "N")
            Sql2 = Sql2 & " and fecha1 = " & DBSet(Rs.Fields(1).Value, "F")
            Sql2 = Sql2 & " and importe2 = " & DBSet(Rs.Fields(2).Value, "N")
            Sql2 = Sql2 & " and importe3 = " & DBSet(Rs.Fields(3).Value, "N")
            
            If TotalRegistros(Sql2) = 0 Then
                Sql3 = "(" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & ","
                Sql3 = Sql3 & DBSet(Rs.Fields(1).Value, "F") & "," & DBSet(Rs.Fields(4).Value, "N")
                Sql3 = Sql3 & "," & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(3).Value, "N") & ","
                Sql3 = Sql3 & DBSet(Rs.Fields(5).Value, "N") & ")"
                
                conn.Execute SQL & Sql3
            Else
                Sql3 = "update tmpinformes set importe1 = imnporte1 + " & DBSet(Rs.Fields(4).Value, "N")
                Sql3 = Sql3 & ", importe4 = importe4 + " & DBSet(Rs.Fields(5).Value, "N")
                Sql3 = Sql3 & " where codusu = " & vUsu.Codigo
                Sql3 = Sql3 & " and codigo1 = " & DBSet(Rs.Fields(0).Value, "N")
                Sql3 = Sql3 & " and fecha1 = " & DBSet(Rs.Fields(1).Value, "F")
                Sql3 = Sql3 & " and importe2 = " & DBSet(Rs.Fields(2).Value, "N")
                Sql3 = Sql3 & " and importe3 = " & DBSet(Rs.Fields(3).Value, "N")
            
                conn.Execute Sql3
            End If
        
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
    End If
    
    Screen.MousePointer = vbDefault
    
    ProcesoCargaIntermedia = True
    Exit Function
    
eProcesoCargaHoras:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso Carga Tabla Intermedia", Err.Description
End Function





