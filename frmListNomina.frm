VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListNomina 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8145
   Icon            =   "frmListNomina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6750
      Top             =   4980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameListMensAsesoria 
      Height          =   4575
      Left            =   30
      TabIndex        =   234
      Top             =   60
      Width           =   6375
      Begin VB.CommandButton CmdAcepInfAse 
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
         Left            =   3975
         TabIndex        =   241
         Top             =   3900
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
         Index           =   12
         Left            =   5115
         TabIndex        =   242
         Top             =   3885
         Width           =   1065
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3105
         Left            =   360
         TabIndex        =   235
         Top             =   870
         Width           =   5805
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
            Index           =   78
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   240
            Top             =   2250
            Width           =   1350
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Exportar Cadena para Excel"
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
            Left            =   0
            TabIndex        =   251
            Top             =   2760
            Width           =   3645
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
            Left            =   1425
            TabIndex        =   239
            Text            =   "Combo2"
            Top             =   1770
            Width           =   1575
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
            Index           =   61
            Left            =   1425
            MaxLength       =   4
            TabIndex        =   238
            Top             =   1290
            Width           =   840
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
            Index           =   65
            Left            =   2265
            Locked          =   -1  'True
            TabIndex        =   244
            Text            =   "Text5"
            Top             =   810
            Width           =   3615
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
            Index           =   64
            Left            =   2265
            Locked          =   -1  'True
            TabIndex        =   243
            Text            =   "Text5"
            Top             =   405
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
            Index           =   64
            Left            =   1425
            MaxLength       =   6
            TabIndex        =   236
            Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
            Text            =   "000000"
            Top             =   405
            Width           =   810
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
            Index           =   65
            Left            =   1425
            MaxLength       =   6
            TabIndex        =   237
            Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   795
            Width           =   810
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   25
            Left            =   1125
            Picture         =   "frmListNomina.frx":000C
            Top             =   2250
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Baja"
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
            Index           =   105
            Left            =   0
            TabIndex        =   301
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
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
            Left            =   0
            TabIndex        =   250
            Top             =   1830
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "A�o"
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
            Left            =   0
            TabIndex        =   248
            Top             =   1320
            Width           =   375
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   28
            Left            =   1125
            MouseIcon       =   "frmListNomina.frx":0097
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   810
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   27
            Left            =   1125
            MouseIcon       =   "frmListNomina.frx":01E9
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label2 
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
            Index           =   84
            Left            =   0
            TabIndex        =   247
            Top             =   60
            Width           =   1065
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
            Index           =   83
            Left            =   210
            TabIndex        =   246
            Top             =   810
            Width           =   720
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
            Index           =   82
            Left            =   210
            TabIndex        =   245
            Top             =   420
            Width           =   765
         End
      End
      Begin VB.Label Label2 
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
         Index           =   107
         Left            =   360
         TabIndex        =   322
         Top             =   4050
         Width           =   3195
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
         Left            =   405
         TabIndex        =   249
         Top             =   360
         Width           =   5595
      End
   End
   Begin VB.Frame FrameInfDiasTrabajados 
      Height          =   4275
      Left            =   30
      TabIndex        =   257
      Top             =   30
      Width           =   6375
      Begin VB.CheckBox Check6 
         Caption         =   "Exportar Cadena para Excel"
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
         Left            =   570
         TabIndex        =   273
         Top             =   3285
         Width           =   3045
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
         Left            =   4905
         TabIndex        =   259
         Top             =   3600
         Width           =   1065
      End
      Begin VB.CommandButton CmdDiasTrabajados 
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
         Left            =   3750
         TabIndex        =   258
         Top             =   3600
         Width           =   1065
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2985
         Left            =   405
         TabIndex        =   260
         Top             =   900
         Width           =   5595
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
            Index           =   69
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   266
            Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   765
            Width           =   750
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
            Index           =   68
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   265
            Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   405
            Width           =   750
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
            Index           =   68
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   264
            Text            =   "Text5"
            Top             =   405
            Width           =   3420
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
            Index           =   69
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   263
            Text            =   "Text5"
            Top             =   780
            Width           =   3420
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
            Index           =   67
            Left            =   1380
            MaxLength       =   4
            TabIndex        =   262
            Top             =   1245
            Width           =   840
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
            Left            =   1380
            TabIndex        =   261
            Text            =   "Combo2"
            Top             =   1815
            Width           =   1575
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
            Index           =   92
            Left            =   390
            TabIndex        =   271
            Top             =   420
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
            Index           =   91
            Left            =   390
            TabIndex        =   270
            Top             =   780
            Width           =   645
         End
         Begin VB.Label Label2 
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
            Index           =   90
            Left            =   180
            TabIndex        =   269
            Top             =   60
            Width           =   1065
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   30
            Left            =   1080
            MouseIcon       =   "frmListNomina.frx":033B
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   810
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   29
            Left            =   1080
            MouseIcon       =   "frmListNomina.frx":048D
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "A�o"
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
            Left            =   180
            TabIndex        =   268
            Top             =   1275
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
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
            Left            =   180
            TabIndex        =   267
            Top             =   1875
            Width           =   390
         End
      End
      Begin VB.Label Label17 
         Caption         =   "Informe Mensual D�as Trabajados"
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
         TabIndex        =   272
         Top             =   390
         Width           =   5595
      End
   End
   Begin VB.Frame FrameBorradoMasivoETT 
      Height          =   3885
      Left            =   0
      TabIndex        =   96
      Top             =   -60
      Width           =   6585
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
         Left            =   4995
         TabIndex        =   106
         Top             =   3135
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepBorradoMasivo 
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
         Left            =   3825
         TabIndex        =   104
         Top             =   3135
         Width           =   1065
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
         Index           =   32
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   100
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   1695
         Width           =   900
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
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   99
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   1305
         Width           =   900
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
         Index           =   31
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   98
         Text            =   "Text5"
         Top             =   1305
         Width           =   3315
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
         Index           =   32
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   97
         Text            =   "Text5"
         Top             =   1710
         Width           =   3315
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
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   102
         Top             =   2790
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
         Index           =   29
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   101
         Top             =   2400
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
         Index           =   42
         Left            =   960
         TabIndex        =   111
         Top             =   1320
         Width           =   615
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
         Index           =   40
         Left            =   960
         TabIndex        =   110
         Top             =   1710
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
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
         Left            =   600
         TabIndex        =   109
         Top             =   990
         Width           =   810
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
         Left            =   600
         TabIndex        =   108
         Top             =   390
         Width           =   5505
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
         Index           =   37
         Left            =   960
         TabIndex        =   107
         Top             =   2400
         Width           =   615
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
         Index           =   35
         Left            =   960
         TabIndex        =   105
         Top             =   2775
         Width           =   570
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
         Index           =   34
         Left            =   600
         TabIndex        =   103
         Top             =   2130
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1605
         MouseIcon       =   "frmListNomina.frx":05DF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   1725
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1605
         MouseIcon       =   "frmListNomina.frx":0731
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   1335
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   1590
         Picture         =   "frmListNomina.frx":0883
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1605
         Picture         =   "frmListNomina.frx":090E
         Top             =   2400
         Width           =   240
      End
   End
   Begin VB.Frame FrameTrabajadoresCapataz 
      Height          =   5055
      Left            =   -60
      TabIndex        =   152
      Top             =   90
      Width           =   6375
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1395
         Left            =   300
         TabIndex        =   304
         Top             =   2610
         Width           =   5085
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
            Index           =   51
            Left            =   270
            MaxLength       =   13
            TabIndex        =   305
            Top             =   480
            Width           =   1680
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
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   60
            Left            =   270
            TabIndex        =   306
            Top             =   120
            Width           =   765
         End
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
         Index           =   45
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   303
         Tag             =   "C�digo|N|N|0|9999|straba|codtraba|0000|S|"
         Top             =   2085
         Width           =   780
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
         Index           =   45
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   302
         Text            =   "Text5"
         Top             =   2085
         Width           =   3240
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
         Index           =   47
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   155
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1110
         Width           =   810
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
         Index           =   47
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   154
         Text            =   "Text5"
         Top             =   1110
         Width           =   3240
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
         Index           =   46
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   156
         Top             =   1605
         Width           =   1350
      End
      Begin VB.CommandButton CmdAcepTrabajCapataz 
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
         Left            =   3630
         TabIndex        =   157
         Top             =   4230
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
         Index           =   7
         Left            =   4770
         TabIndex        =   153
         Top             =   4230
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         TabIndex        =   161
         Top             =   1140
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":0999
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1155
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   1545
         Picture         =   "frmListNomina.frx":0AEB
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":0B76
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   2115
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
         Index           =   62
         Left            =   570
         TabIndex        =   160
         Top             =   1620
         Width           =   600
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
         Left            =   585
         TabIndex        =   159
         Top             =   390
         Width           =   5595
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
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
         Left            =   570
         TabIndex        =   158
         Top             =   2100
         Width           =   810
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
         Left            =   420
         TabIndex        =   26
         Top             =   3360
         Width           =   3120
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
         Left            =   5715
         TabIndex        =   10
         Top             =   3735
         Width           =   1065
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
         Left            =   4500
         TabIndex        =   8
         Top             =   3735
         Width           =   1065
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
         Index           =   19
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1665
         Width           =   750
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
         Index           =   18
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1305
         Width           =   750
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
         Index           =   18
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   1305
         Width           =   4050
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
         Index           =   19
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text5"
         Top             =   1680
         Width           =   4050
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
         Index           =   17
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2745
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
         Index           =   16
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2370
         Width           =   1350
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
         Left            =   915
         TabIndex        =   15
         Top             =   1320
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
         Left            =   915
         TabIndex        =   14
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label2 
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
         Index           =   27
         Left            =   420
         TabIndex        =   13
         Top             =   990
         Width           =   1065
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
         Left            =   915
         TabIndex        =   11
         Top             =   2400
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
         Left            =   915
         TabIndex        =   9
         Top             =   2715
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
         Left            =   420
         TabIndex        =   7
         Top             =   2070
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1620
         MouseIcon       =   "frmListNomina.frx":0CC8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1620
         MouseIcon       =   "frmListNomina.frx":0E1A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1575
         Picture         =   "frmListNomina.frx":0F6C
         Top             =   2745
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1575
         Picture         =   "frmListNomina.frx":0FF7
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
         TabIndex        =   253
         Top             =   3090
         Width           =   4155
         Begin VB.CheckBox Check5 
            Caption         =   "Prevision de Pago de Partes"
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
            Left            =   45
            TabIndex        =   254
            Top             =   240
            Width           =   3120
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   3270
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   240
            Width           =   240
         End
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
         Index           =   15
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   33
         Top             =   2745
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
         Index           =   14
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   32
         Top             =   2340
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
         Index           =   1
         Left            =   1710
         MaxLength       =   7
         TabIndex        =   31
         Tag             =   "N� Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1665
         Width           =   960
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
         Left            =   1710
         MaxLength       =   7
         TabIndex        =   30
         Tag             =   "N� Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1260
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
         Index           =   0
         Left            =   3435
         TabIndex        =   35
         Top             =   3690
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
         Left            =   4605
         TabIndex        =   37
         Top             =   3690
         Width           =   1065
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1395
         Picture         =   "frmListNomina.frx":1082
         Top             =   2745
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1395
         Picture         =   "frmListNomina.frx":110D
         Top             =   2340
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
         Index           =   5
         Left            =   420
         TabIndex        =   42
         Top             =   2115
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
         Index           =   4
         Left            =   735
         TabIndex        =   41
         Top             =   2715
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
         Index           =   3
         Left            =   735
         TabIndex        =   40
         Top             =   2400
         Width           =   735
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
         Left            =   405
         TabIndex        =   39
         Top             =   450
         Width           =   4350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parte"
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
         TabIndex        =   38
         Top             =   1035
         Width           =   525
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
         Left            =   735
         TabIndex        =   36
         Top             =   1680
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
         Index           =   0
         Left            =   735
         TabIndex        =   34
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame FrameInfComprobacion 
      Height          =   5085
      Left            =   0
      TabIndex        =   162
      Top             =   0
      Width           =   6915
      Begin VB.CheckBox Check7 
         Caption         =   "Resumen por trabajador"
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
         Height          =   315
         Left            =   630
         TabIndex        =   300
         Top             =   4005
         Visible         =   0   'False
         Width           =   2835
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
         Index           =   77
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   296
         Text            =   "Text5"
         Top             =   3495
         Width           =   3735
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
         Index           =   76
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   295
         Text            =   "Text5"
         Top             =   3120
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
         Index           =   77
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   170
         Tag             =   "C�digo|N|N|0|9999|straba|codtraba|0000|S|"
         Top             =   3495
         Width           =   840
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
         Index           =   76
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   169
         Tag             =   "C�digo|N|N|0|9999|straba|codtraba|0000|S|"
         Top             =   3105
         Width           =   840
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
         Left            =   5400
         TabIndex        =   174
         Top             =   4350
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepInfComprob 
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
         Left            =   4185
         TabIndex        =   171
         Top             =   4350
         Width           =   1065
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
         Index           =   50
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   166
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1620
         Width           =   840
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
         Index           =   49
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   165
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1215
         Width           =   840
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
         Index           =   49
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   164
         Text            =   "Text5"
         Top             =   1215
         Width           =   3735
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
         Index           =   50
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   163
         Text            =   "Text5"
         Top             =   1635
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
         Index           =   44
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   167
         Top             =   2130
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
         Index           =   48
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   168
         Top             =   2520
         Width           =   1350
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   34
         Left            =   1560
         MouseIcon       =   "frmListNomina.frx":1198
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   3495
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   33
         Left            =   1560
         MouseIcon       =   "frmListNomina.frx":12EA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   3105
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
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
         Left            =   600
         TabIndex        =   299
         Top             =   2850
         Width           =   810
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
         Index           =   103
         Left            =   870
         TabIndex        =   298
         Top             =   3495
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
         Index           =   102
         Left            =   870
         TabIndex        =   297
         Top             =   3135
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
         Index           =   66
         Left            =   870
         TabIndex        =   179
         Top             =   1230
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
         Left            =   870
         TabIndex        =   178
         Top             =   1590
         Width           =   645
      End
      Begin VB.Label Label2 
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
         Index           =   64
         Left            =   600
         TabIndex        =   177
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label Label11 
         Caption         =   "Informe de Comprobaci�n"
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
         TabIndex        =   176
         Top             =   390
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
         Index           =   59
         Left            =   870
         TabIndex        =   175
         Top             =   2145
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
         Index           =   58
         Left            =   870
         TabIndex        =   173
         Top             =   2505
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
         Index           =   57
         Left            =   600
         TabIndex        =   172
         Top             =   1905
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":143C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":158E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   1560
         Picture         =   "frmListNomina.frx":16E0
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1560
         Picture         =   "frmListNomina.frx":176B
         Top             =   2115
         Width           =   240
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
         Height          =   3405
         Left            =   300
         TabIndex        =   140
         Top             =   1440
         Width           =   5955
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
            Left            =   1530
            MaxLength       =   6
            TabIndex        =   135
            Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   2085
            Width           =   870
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
            Left            =   1530
            MaxLength       =   6
            TabIndex        =   134
            Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
            Top             =   1695
            Width           =   870
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
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   148
            Text            =   "Text5"
            Top             =   1695
            Width           =   3375
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
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   147
            Text            =   "Text5"
            Top             =   2100
            Width           =   3375
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
            Left            =   270
            MaxLength       =   13
            TabIndex        =   136
            Top             =   3000
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
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   132
            Top             =   540
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
            Index           =   33
            Left            =   1545
            MaxLength       =   10
            TabIndex        =   133
            Top             =   930
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
            Index           =   54
            Left            =   540
            TabIndex        =   151
            Top             =   1710
            Width           =   615
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
            Index           =   53
            Left            =   540
            TabIndex        =   150
            Top             =   2100
            Width           =   630
         End
         Begin VB.Label Label2 
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
            Index           =   52
            Left            =   240
            TabIndex        =   149
            Top             =   1410
            Width           =   1065
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   13
            Left            =   1215
            MouseIcon       =   "frmListNomina.frx":17F6
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   2085
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   1215
            MouseIcon       =   "frmListNomina.frx":1948
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   1680
            Width           =   240
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
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   50
            Left            =   270
            TabIndex        =   144
            Top             =   2700
            Width           =   765
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
            Index           =   44
            Left            =   270
            TabIndex        =   143
            Top             =   180
            Width           =   600
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   12
            Left            =   1215
            Picture         =   "frmListNomina.frx":1A9A
            Top             =   540
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   11
            Left            =   1200
            Picture         =   "frmListNomina.frx":1B25
            Top             =   930
            Width           =   240
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
            Index           =   43
            Left            =   510
            TabIndex        =   142
            Top             =   960
            Width           =   630
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
            Index           =   33
            Left            =   510
            TabIndex        =   141
            Top             =   570
            Width           =   645
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
         Index           =   6
         Left            =   5010
         TabIndex        =   138
         Top             =   4875
         Width           =   1065
      End
      Begin VB.CommandButton CmdEventuales 
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
         Left            =   3870
         TabIndex        =   137
         Top             =   4875
         Width           =   1065
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
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   139
         Text            =   "Text5"
         Top             =   1110
         Width           =   3315
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
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   131
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1110
         Width           =   870
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
         Left            =   570
         TabIndex        =   146
         Top             =   390
         Width           =   5595
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":1BB0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Left            =   570
         TabIndex        =   145
         Top             =   1140
         Width           =   855
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
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   2190
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
         Index           =   24
         Left            =   1650
         MaxLength       =   2
         TabIndex        =   20
         Top             =   2190
         Width           =   960
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
         Index           =   27
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1290
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
         Index           =   25
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   19
         Top             =   1740
         Width           =   990
      End
      Begin VB.CommandButton CmdAcepCalHProd 
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
         Left            =   3390
         TabIndex        =   21
         Top             =   2790
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancelCalHProd 
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
         Left            =   4545
         TabIndex        =   22
         Top             =   2790
         Width           =   1065
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1350
         MouseIcon       =   "frmListNomina.frx":1D02
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar almac�n"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Almac�n"
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
         Left            =   450
         TabIndex        =   28
         Top             =   2250
         Width           =   825
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1350
         Picture         =   "frmListNomina.frx":1E54
         Top             =   1290
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
         Index           =   41
         Left            =   450
         TabIndex        =   25
         Top             =   1290
         Width           =   600
      End
      Begin VB.Label Label8 
         Caption         =   "C�lculo de Horas Productivas"
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
         Index           =   38
         Left            =   450
         TabIndex        =   23
         Top             =   1800
         Width           =   1065
      End
   End
   Begin VB.Frame FrameBorradoAsesoria 
      Height          =   4215
      Left            =   60
      TabIndex        =   196
      Top             =   0
      Width           =   6705
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
         Index           =   57
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   200
         Top             =   2880
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
         Index           =   56
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   199
         Top             =   2475
         Width           =   1350
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
         Index           =   55
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   203
         Text            =   "Text5"
         Top             =   1740
         Width           =   3525
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
         Index           =   54
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   201
         Text            =   "Text5"
         Top             =   1305
         Width           =   3525
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
         Index           =   55
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   198
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1725
         Width           =   840
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
         Index           =   54
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   197
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Text            =   "000000"
         Top             =   1305
         Width           =   840
      End
      Begin VB.CommandButton CmdAcepBorrAse 
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
         TabIndex        =   202
         Top             =   3435
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
         Index           =   10
         Left            =   5220
         TabIndex        =   204
         Top             =   3435
         Width           =   1065
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   1605
         Picture         =   "frmListNomina.frx":1EDF
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   1605
         Picture         =   "frmListNomina.frx":1F6A
         Top             =   2475
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   1620
         MouseIcon       =   "frmListNomina.frx":1FF5
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1725
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   22
         Left            =   1620
         MouseIcon       =   "frmListNomina.frx":2147
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
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
         Index           =   75
         Left            =   360
         TabIndex        =   211
         Top             =   2190
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
         Index           =   74
         Left            =   720
         TabIndex        =   210
         Top             =   2895
         Width           =   630
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
         Index           =   73
         Left            =   720
         TabIndex        =   209
         Top             =   2520
         Width           =   645
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
         Left            =   360
         TabIndex        =   208
         Top             =   405
         Width           =   5925
      End
      Begin VB.Label Label2 
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
         Index           =   72
         Left            =   360
         TabIndex        =   207
         Top             =   1020
         Width           =   1065
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
         Index           =   71
         Left            =   720
         TabIndex        =   206
         Top             =   1770
         Width           =   630
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
         Index           =   70
         Left            =   720
         TabIndex        =   205
         Top             =   1350
         Width           =   675
      End
   End
   Begin VB.Frame FrameCapatazServicios 
      Height          =   3135
      Left            =   0
      TabIndex        =   307
      Top             =   0
      Width           =   6375
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
         Left            =   4770
         TabIndex        =   313
         Top             =   2490
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepCapat 
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
         Left            =   3630
         TabIndex        =   311
         Top             =   2490
         Width           =   1065
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
         Index           =   82
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   309
         Top             =   1215
         Width           =   1350
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
         Index           =   80
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   308
         Text            =   "Text5"
         Top             =   1665
         Width           =   3195
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
         Index           =   80
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   310
         Tag             =   "C�digo|N|N|0|9999|straba|codtraba|0000|S|"
         Top             =   1665
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
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
         Left            =   570
         TabIndex        =   315
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label Label19 
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
         Left            =   585
         TabIndex        =   314
         Top             =   390
         Width           =   5595
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
         Index           =   108
         Left            =   570
         TabIndex        =   312
         Top             =   1230
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   36
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":2299
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   1695
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   26
         Left            =   1545
         Picture         =   "frmListNomina.frx":23EB
         Top             =   1245
         Width           =   240
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
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   114
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Text            =   "000000"
         Top             =   1110
         Width           =   840
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
         Index           =   36
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text5"
         Top             =   1110
         Width           =   3495
      End
      Begin VB.CommandButton CmdAltaRapida 
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
         Left            =   3885
         TabIndex        =   119
         Top             =   4245
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
         Index           =   4
         Left            =   5070
         TabIndex        =   120
         Top             =   4245
         Width           =   1065
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2685
         Left            =   300
         TabIndex        =   121
         Top             =   1500
         Width           =   5955
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
            Index           =   26
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   116
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
            Index           =   35
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   115
            Top             =   480
            Width           =   1350
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
            Index           =   34
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   125
            Text            =   "Text5"
            Top             =   1440
            Width           =   3435
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
            Index           =   34
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   117
            Tag             =   "C�digo|N|N|0|999999|straba|codtraba|0000|S|"
            Top             =   1440
            Width           =   870
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
            Left            =   270
            MaxLength       =   13
            TabIndex        =   118
            Top             =   2250
            Width           =   1740
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
            Index           =   32
            Left            =   510
            TabIndex        =   129
            Top             =   510
            Width           =   705
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
            Index           =   31
            Left            =   510
            TabIndex        =   128
            Top             =   930
            Width           =   660
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   10
            Left            =   1200
            Picture         =   "frmListNomina.frx":2476
            Top             =   900
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   9
            Left            =   1215
            Picture         =   "frmListNomina.frx":2501
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1215
            MouseIcon       =   "frmListNomina.frx":258C
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   1470
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
            Index           =   48
            Left            =   270
            TabIndex        =   127
            Top             =   180
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Capataz"
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
            Index           =   47
            Left            =   270
            TabIndex        =   126
            Top             =   1485
            Width           =   810
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
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   46
            Left            =   270
            TabIndex        =   122
            Top             =   1950
            Width           =   765
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Left            =   570
         TabIndex        =   124
         Top             =   1140
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":26DE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Alta R�pida"
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
         Left            =   570
         TabIndex        =   123
         Top             =   390
         Width           =   5595
      End
   End
   Begin VB.Frame FrameImpresionParte 
      Height          =   5445
      Left            =   0
      TabIndex        =   274
      Top             =   0
      Width           =   6285
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
         Index           =   75
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   285
         Top             =   1620
         Width           =   1005
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
         Index           =   74
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   284
         Top             =   1245
         Width           =   1005
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
         Left            =   4710
         TabIndex        =   291
         Top             =   4545
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepImpPartes 
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
         Left            =   3585
         TabIndex        =   290
         Top             =   4545
         Width           =   1065
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
         Index           =   73
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   289
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   3855
         Width           =   840
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
         Index           =   72
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   288
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   3450
         Width           =   840
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
         Index           =   72
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   276
         Text            =   "Text5"
         Top             =   3450
         Width           =   3195
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
         Index           =   73
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   275
         Text            =   "Text5"
         Top             =   3870
         Width           =   3195
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
         Index           =   71
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   287
         Top             =   2670
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
         Index           =   70
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   286
         Top             =   2295
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
         Index           =   101
         Left            =   900
         TabIndex        =   294
         Top             =   1275
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
         Index           =   100
         Left            =   900
         TabIndex        =   293
         Top             =   1635
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parte"
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
         Index           =   99
         Left            =   630
         TabIndex        =   292
         Top             =   960
         Width           =   525
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
         Left            =   840
         TabIndex        =   283
         Top             =   3435
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
         Index           =   97
         Left            =   840
         TabIndex        =   282
         Top             =   3840
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
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
         Left            =   570
         TabIndex        =   281
         Top             =   3105
         Width           =   810
      End
      Begin VB.Label Label18 
         Caption         =   "Impresi�n de Partes"
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
         TabIndex        =   280
         Top             =   390
         Width           =   5505
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
         Index           =   95
         Left            =   870
         TabIndex        =   279
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
         Index           =   94
         Left            =   870
         TabIndex        =   278
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
         Index           =   93
         Left            =   600
         TabIndex        =   277
         Top             =   1995
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   32
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":2830
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   3855
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   31
         Left            =   1560
         MouseIcon       =   "frmListNomina.frx":2982
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   3450
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   1560
         Picture         =   "frmListNomina.frx":2AD4
         Top             =   2670
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   1575
         Picture         =   "frmListNomina.frx":2B5F
         Top             =   2295
         Width           =   240
      End
   End
   Begin VB.Frame FrameCalculoETT 
      Height          =   5055
      Left            =   0
      TabIndex        =   67
      Top             =   30
      Width           =   6375
      Begin VB.Frame FrameBonificacion 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1395
         Left            =   300
         TabIndex        =   93
         Top             =   2670
         Width           =   5805
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
            Index           =   23
            Left            =   270
            MaxLength       =   13
            TabIndex        =   94
            Top             =   480
            Width           =   1620
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
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   30
            Left            =   270
            TabIndex        =   95
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.Frame FramePenalizacion 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1395
         Left            =   210
         TabIndex        =   86
         Top             =   2550
         Width           =   6015
         Begin VB.TextBox txtCodigo 
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
            Index           =   22
            Left            =   360
            MaxLength       =   6
            TabIndex        =   89
            Top             =   720
            Width           =   1140
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
            Index           =   21
            Left            =   2190
            MaxLength       =   6
            TabIndex        =   88
            Top             =   720
            Width           =   1140
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   20
            Left            =   4020
            MaxLength       =   11
            TabIndex        =   87
            Top             =   720
            Width           =   1560
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kilos Entrados"
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
            Index           =   23
            Left            =   360
            TabIndex        =   92
            Top             =   390
            Width           =   1380
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "% Penalizacion"
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
            Left            =   2190
            TabIndex        =   91
            Top             =   390
            Width           =   1470
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
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   20
            Left            =   4020
            TabIndex        =   90
            Top             =   390
            Width           =   765
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
         Index           =   2
         Left            =   5010
         TabIndex        =   75
         Top             =   4230
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepCalculoETT 
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
         Left            =   3870
         TabIndex        =   74
         Top             =   4230
         Width           =   1065
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
         Index           =   12
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   70
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   2070
         Width           =   750
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
         Index           =   12
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Text5"
         Top             =   2070
         Width           =   3435
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
         Index           =   11
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   69
         Top             =   1605
         Width           =   1350
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
         Index           =   9
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text5"
         Top             =   1110
         Width           =   3375
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
         Index           =   9
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   68
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Text            =   "000000"
         Top             =   1110
         Width           =   840
      End
      Begin VB.Frame FrameDestajo 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1395
         Left            =   270
         TabIndex        =   82
         Top             =   2700
         Width           =   5685
         Begin VB.TextBox txtCodigo 
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
            Index           =   13
            Left            =   3450
            MaxLength       =   13
            TabIndex        =   73
            Top             =   720
            Width           =   1470
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   8
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   72
            Top             =   720
            Width           =   1140
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   10
            Left            =   330
            MaxLength       =   6
            TabIndex        =   71
            Top             =   720
            Width           =   1140
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
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   19
            Left            =   3450
            TabIndex        =   85
            Top             =   390
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   16
            Left            =   2040
            TabIndex        =   84
            Top             =   390
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kilos Entrados"
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
            Left            =   330
            TabIndex        =   83
            Top             =   390
            Width           =   1380
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
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
         Left            =   570
         TabIndex        =   81
         Top             =   2100
         Width           =   810
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
         Left            =   570
         TabIndex        =   80
         Top             =   390
         Width           =   5595
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
         Index           =   18
         Left            =   570
         TabIndex        =   79
         Top             =   1620
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1515
         MouseIcon       =   "frmListNomina.frx":2BEA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   2085
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1515
         Picture         =   "frmListNomina.frx":2D3C
         Top             =   1605
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1545
         MouseIcon       =   "frmListNomina.frx":2DC7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   17
         Left            =   570
         TabIndex        =   78
         Top             =   1140
         Width           =   855
      End
   End
   Begin VB.Frame FrameTrabajadoresActivos 
      Height          =   2715
      Left            =   0
      TabIndex        =   316
      Top             =   0
      Width           =   6285
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
         Left            =   4770
         TabIndex        =   319
         Top             =   1905
         Width           =   1065
      End
      Begin VB.CommandButton CmdTrabajadoresActivos 
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
         Left            =   3600
         TabIndex        =   318
         Top             =   1905
         Width           =   1065
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
         Index           =   81
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   317
         Top             =   1215
         Width           =   1350
      End
      Begin VB.Label Label20 
         Caption         =   "Trabajadores en Activo"
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
         Top             =   390
         Width           =   5505
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
         Index           =   106
         Left            =   630
         TabIndex        =   320
         Top             =   1215
         Width           =   600
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   28
         Left            =   1575
         Picture         =   "frmListNomina.frx":2F19
         Top             =   1215
         Width           =   240
      End
   End
   Begin VB.Frame FrameEntradasCapataz 
      Height          =   3885
      Left            =   0
      TabIndex        =   180
      Top             =   0
      Width           =   6285
      Begin VB.CheckBox Check4 
         Caption         =   "Imprimir resumen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   405
         TabIndex        =   252
         Top             =   3270
         Width           =   1995
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
         Index           =   53
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   184
         Top             =   2715
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
         Index           =   52
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   183
         Top             =   2340
         Width           =   1350
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
         Index           =   38
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   188
         Text            =   "Text5"
         Top             =   1290
         Width           =   3285
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
         Index           =   43
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   187
         Text            =   "Text5"
         Top             =   1665
         Width           =   3285
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
         Index           =   43
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   182
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   1665
         Width           =   930
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
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   181
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|0000|S|"
         Top             =   1275
         Width           =   930
      End
      Begin VB.CommandButton CmdAcepEntCapataz 
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
         Left            =   3600
         TabIndex        =   185
         Top             =   3255
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
         Index           =   9
         Left            =   4770
         TabIndex        =   186
         Top             =   3255
         Width           =   1065
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1320
         Picture         =   "frmListNomina.frx":2FA4
         Top             =   2715
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1335
         Picture         =   "frmListNomina.frx":302F
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1335
         MouseIcon       =   "frmListNomina.frx":30BA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1350
         MouseIcon       =   "frmListNomina.frx":320C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar capataz"
         Top             =   1320
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
         Index           =   69
         Left            =   375
         TabIndex        =   195
         Top             =   2040
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
         Index           =   68
         Left            =   645
         TabIndex        =   194
         Top             =   2730
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
         Index           =   67
         Left            =   645
         TabIndex        =   193
         Top             =   2370
         Width           =   645
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
         Left            =   405
         TabIndex        =   192
         Top             =   390
         Width           =   5505
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Capataz"
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
         Left            =   375
         TabIndex        =   191
         Top             =   990
         Width           =   810
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
         Index           =   55
         Left            =   645
         TabIndex        =   190
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
         Index           =   45
         Left            =   645
         TabIndex        =   189
         Top             =   1320
         Width           =   645
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
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   47
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   2730
         Width           =   870
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
         Index           =   6
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   46
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   2355
         Width           =   870
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
         Index           =   6
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text5"
         Top             =   2370
         Width           =   3825
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
         Index           =   7
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text5"
         Top             =   2745
         Width           =   3825
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Informe para el Trabajador"
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
         Left            =   630
         TabIndex        =   61
         Top             =   4320
         Width           =   3075
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
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   49
         Top             =   3780
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
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   48
         Top             =   3390
         Width           =   1350
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
         Index           =   3
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text5"
         Top             =   1680
         Width           =   3825
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
         Index           =   2
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Text5"
         Top             =   1305
         Width           =   3825
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
         Index           =   2
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   44
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1305
         Width           =   870
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
         Index           =   3
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   45
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1665
         Width           =   870
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
         Left            =   4230
         TabIndex        =   50
         Top             =   4665
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
         Index           =   1
         Left            =   5445
         TabIndex        =   51
         Top             =   4665
         Width           =   1065
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
         Left            =   870
         TabIndex        =   66
         Top             =   2385
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
         Index           =   13
         Left            =   870
         TabIndex        =   65
         Top             =   2745
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Left            =   600
         TabIndex        =   64
         Top             =   2055
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":335E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2730
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":34B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2370
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1575
         Picture         =   "frmListNomina.frx":3602
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   1575
         Picture         =   "frmListNomina.frx":368D
         Top             =   3390
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":3718
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1575
         MouseIcon       =   "frmListNomina.frx":386A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1665
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
         Index           =   11
         Left            =   600
         TabIndex        =   60
         Top             =   3075
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
         Index           =   10
         Left            =   870
         TabIndex        =   59
         Top             =   3765
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
         Index           =   9
         Left            =   870
         TabIndex        =   58
         Top             =   3405
         Width           =   690
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
         Index           =   8
         Left            =   600
         TabIndex        =   56
         Top             =   990
         Width           =   1065
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
         Index           =   7
         Left            =   870
         TabIndex        =   55
         Top             =   1680
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
         Index           =   6
         Left            =   870
         TabIndex        =   54
         Top             =   1320
         Width           =   690
      End
   End
   Begin VB.Frame FramePaseABanco 
      Height          =   5990
      Left            =   60
      TabIndex        =   212
      Top             =   30
      Width           =   6435
      Begin VB.Frame FrameConcep 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   540
         TabIndex        =   255
         Top             =   4410
         Width           =   5715
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
            Index           =   66
            Left            =   1290
            MaxLength       =   30
            TabIndex        =   225
            Top             =   120
            Width           =   4290
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripci�n"
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
            Left            =   0
            TabIndex        =   256
            Top             =   90
            Width           =   1125
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
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   224
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   4080
         Width           =   1665
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
         Index           =   62
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   231
         Text            =   "Text5"
         Top             =   1230
         Width           =   3405
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
         Index           =   63
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   232
         Text            =   "Text5"
         Top             =   1635
         Width           =   3405
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
         Index           =   63
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   220
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1635
         Width           =   840
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
         Index           =   62
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   219
         Tag             =   "C�digo|N|N|0|999999|straba|codtraba|000000|S|"
         Top             =   1215
         Width           =   840
      End
      Begin VB.CommandButton CmdAcepPaseBanco 
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
         Left            =   3945
         TabIndex        =   226
         Top             =   5370
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
         Index           =   11
         Left            =   5070
         TabIndex        =   227
         Top             =   5370
         Width           =   1065
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
         Index           =   60
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   222
         Top             =   2910
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
         Index           =   59
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   221
         Top             =   2310
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
         Index           =   58
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   223
         Tag             =   "C�digo|N|N|0|9999|rcapataz|codcapat|0000|S|"
         Top             =   3390
         Width           =   840
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
         Index           =   58
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   233
         Text            =   "Text5"
         Top             =   3390
         Width           =   3405
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   240
         Left            =   480
         TabIndex        =   213
         Top             =   5070
         Visible         =   0   'False
         Width           =   5670
         _ExtentX        =   10001
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   5190
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   1530
         MouseIcon       =   "frmListNomina.frx":39BC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1530
         MouseIcon       =   "frmListNomina.frx":3B0E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   81
         Left            =   510
         TabIndex        =   230
         Top             =   930
         Width           =   1065
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
         Index           =   80
         Left            =   870
         TabIndex        =   229
         Top             =   1650
         Width           =   630
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
         Index           =   79
         Left            =   870
         TabIndex        =   228
         Top             =   1230
         Width           =   675
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
         TabIndex        =   218
         Top             =   405
         Width           =   5835
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1515
         Picture         =   "frmListNomina.frx":3C60
         ToolTipText     =   "Buscar fecha"
         Top             =   2910
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   78
         Left            =   540
         TabIndex        =   217
         Top             =   2610
         Width           =   1155
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
         Index           =   77
         Left            =   540
         TabIndex        =   216
         Top             =   2040
         Width           =   1320
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   1530
         Picture         =   "frmListNomina.frx":3CEB
         ToolTipText     =   "Buscar fecha"
         Top             =   2310
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Banco "
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
         Left            =   540
         TabIndex        =   215
         Top             =   3360
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1515
         MouseIcon       =   "frmListNomina.frx":3D76
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar banco"
         Top             =   3390
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "Concepto Transferencia "
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
         Left            =   540
         TabIndex        =   214
         Top             =   3810
         Width           =   1875
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
    
    ' 41 .- Trabajadores de un capataz servicios especiales
    
    
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
    ' 40 .- Impresion de partes de trabajo
    
    ' 42 .- Listado de trabajadores en activo en una fecha dada
    
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmTra As frmManTraba 'mantenimiento de trabajadores
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmAlm As frmBasico2 'mantenimiento de almacenes propios de comercial
Attribute frmAlm.VB_VarHelpID = -1
 
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCap As frmManCapataz 'capataces
Attribute frmCap.VB_VarHelpID = -1

Private WithEvents frmBan As frmBasico2 'Banco propio
Attribute frmBan.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub CmdAcepBorradoMasivo_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim cWhere As String
Dim SQL As String

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
        SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            cWhere = QuitarCaracterACadena(cWhere, "{")
            cWhere = QuitarCaracterACadena(cWhere, "}")
            cWhere = QuitarCaracterACadena(cWhere, "_1")
            SQL = SQL & " WHERE " & cWhere
        End If
    
        Dim NumF As Long
        NumF = TotalRegistros(SQL)
        If NumF <> 0 Then
            If MsgBox("Va a eliminar " & NumF & " registros. � Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            If ProcesoBorradoMasivo(cTabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
                Exit Sub
            Else
                MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "No hay registros entre esos l�mites.", vbExclamation
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
Dim SQL As String


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
        SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            cWhere = QuitarCaracterACadena(cWhere, "{")
            cWhere = QuitarCaracterACadena(cWhere, "}")
            cWhere = QuitarCaracterACadena(cWhere, "_1")
            SQL = SQL & " WHERE " & cWhere
        End If
    
        If RegistrosAListar(SQL) <> 0 Then
            If ProcesoBorradoMasivo(cTabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
                Exit Sub
            Else
                MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "No hay registros entre esos l�mites.", vbExclamation
        End If

        

End Sub

Private Sub CmdAcepCalculoETT_Click()
Dim SQL As String
Dim CodigoETT As String

    If txtCodigo(9).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo Variedad.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(11).Text = "" Then
        MsgBox "Debe introducir una Fecha para realizar el c�lculo.", vbExclamation
        Exit Sub
    End If

    If txtCodigo(12).Text = "" Then
        MsgBox "Debe introducir el capataz para realizar el c�lculo.", vbExclamation
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
            SQL = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
            
            CodigoETT = DevuelveValor(SQL)
        
            SQL = "select count(*) from horasett where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            SQL = SQL & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            SQL = SQL & " and codigoett = " & DBSet(CodigoETT, "N")
            If TotalRegistros(SQL) = 0 Then
                MsgBox "No existe registro para realizar la penalizaci�n. Revise.", vbExclamation
            Else
                If CalculoPenalizacionETT(True) Then
                     MsgBox "Proceso realizado correctamente.", vbExclamation
                    
                     cmdCancel_Click (2)
                End If
            End If
        
        Case 31 'horas: calculo de penalizacion
            SQL = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            SQL = SQL & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            If TotalRegistros(SQL) = 0 Then
                MsgBox "No existen registros para realizar la penalizaci�n. Revise.", vbExclamation
            Else
                If CalculoPenalizacion(True) Then
                     MsgBox "Proceso realizado correctamente.", vbExclamation
                    
                     cmdCancel_Click (2)
                End If
            End If
        
        Case 22 ' horasett: calculo de bonificacion
            SQL = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
            
            CodigoETT = DevuelveValor(SQL)
        
            SQL = "select count(*) from horasett where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            SQL = SQL & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            SQL = SQL & " and codigoett = " & DBSet(CodigoETT, "N")
            If TotalRegistros(SQL) = 0 Then
                MsgBox "No existen registros para realizar la bonificaci�n. Revise.", vbExclamation
            Else
                If CalculoBonificacionETT(True) Then
                     MsgBox "Proceso realizado correctamente.", vbExclamation
                    
                     cmdCancel_Click (2)
                End If
            End If
        
        Case 32 ' horas: calculo de bonificacion
            SQL = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            SQL = SQL & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            If TotalRegistros(SQL) = 0 Then
                MsgBox "No existen registros para realizar la bonificaci�n. Revise.", vbExclamation
            Else
                If CalculoBonificacion(True) Then
                     MsgBox "Proceso realizado correctamente.", vbExclamation
                    
                     cmdCancel_Click (2)
                End If
            End If
        
    End Select
End Sub

Private Sub CmdAcepCalHProd_Click()
Dim SQL As String

    If txtCodigo(27).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo Fecha para realizar el c�lculo.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(25).Text = "" Then
        MsgBox "Debe introducir un porcentaje para realizar el c�lculo.", vbExclamation
        Exit Sub
    End If

    If txtCodigo(24).Text = "" Then
        MsgBox "Debe introducir el almac�n para realizar el c�lculo.", vbExclamation
        Exit Sub
    End If
    
    SQL = "select * from horas where fechahora = " & DBSet(txtCodigo(27).Text, "F")
    SQL = SQL & " and codalmac = " & DBSet(txtCodigo(24), "N")
    SQL = SQL & " and codtraba in (select codtraba from straba where codsecci = 1)"

    If TotalRegistros(SQL) = 0 Then
        MsgBox "No existen registros para esa fecha en el almac�n introducido. Revise.", vbExclamation
        PonerFoco txtCodigo(27)
    Else
        If CalculoHorasProductivas Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
           
            cmdCancelCalHProd_Click
        End If
    End If
End Sub

Private Sub CmdAcepCapat_Click()
Dim SQL As String
Dim CodigoETT As String

    If txtCodigo(82).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en la Fecha.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(80).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo capataz.", vbExclamation
        Exit Sub
    End If
    
    If CalculoCapatazServicios() Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
       
        vCadBusqueda = "horas.fechahora = " & DBSet(txtCodigo(82).Text, "F") & " and horas.codcapat = " & DBSet(txtCodigo(80).Text, "N") & " and horas.codvarie = 0"
       
        cmdCancel_Click (2)
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
Dim SQL As String

       InicializarVbles
       
        'A�adir el parametro de Empresa
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
        SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            cWhere = QuitarCaracterACadena(cWhere, "{")
            cWhere = QuitarCaracterACadena(cWhere, "}")
            cWhere = QuitarCaracterACadena(cWhere, "_1")
            SQL = SQL & " WHERE " & cWhere
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
                    MsgBox "No hay registros entre esos l�mites.", vbExclamation
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
                    MsgBox "No hay registros entre esos l�mites.", vbExclamation
                End If
            End If
            
        
        End If


End Sub

Private Sub CmdAcepImpPartes_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim cWhere As String
Dim SQL As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

       InicializarVbles
       
        'A�adir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

        'D/H Capataz
        cDesde = Trim(txtCodigo(72).Text)
        cHasta = Trim(txtCodigo(73).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rcuadrilla.codcapat}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCapataz=""") Then Exit Sub
        End If
        
        'D/H Fecha
        cDesde = Trim(txtCodigo(70).Text)
        cHasta = Trim(txtCodigo(71).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fechapar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If

        'D/H Parte
        cDesde = Trim(txtCodigo(74).Text)
        cHasta = Trim(txtCodigo(75).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".nroparte}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHParte=""") Then Exit Sub
        End If

        cTabla = Tabla & " inner join rcuadrilla on rpartes.codcuadrilla = rcuadrilla.codcuadrilla "
        cWhere = cadSelect
        
        cTabla = QuitarCaracterACadena(cTabla, "{")
        cTabla = QuitarCaracterACadena(cTabla, "}")
        SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            cWhere = QuitarCaracterACadena(cWhere, "{")
            cWhere = QuitarCaracterACadena(cWhere, "}")
            cWhere = QuitarCaracterACadena(cWhere, "_1")
            SQL = SQL & " WHERE " & cWhere
        End If
    
        If HayRegParaInforme(cTabla, cWhere) Then
        
            If vParamAplic.Cooperativa = 16 Then
                CargarTemporalNotas cTabla, cWhere
                CadParam = CadParam & "pUsu=" & vUsu.Codigo & "|"
                numParam = numParam + 1
            End If
        
        
            'Nombre fichero .rpt a Imprimir
            indRPT = 111 ' impresion de partes de trabajo
            
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = nomDocu '"rParteTrabajo.rpt"
            cadTitulo = "Impresi�n de Partes"
            ConSubInforme = True
            LlamarImprimir
        End If


End Sub

Private Sub CargarTemporalNotas(cTabla As String, cWhere As String)
Dim SQL As String
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
                                           'nroparte, numnota,  codsocio
    SQL = "insert into tmpinformes (codusu, importe1, importe2, importe3) "
    SQL = SQL & " select " & vUsu.Codigo & ", rpartes_variedad.nroparte, rhisfruta_entradas.numnotac, rhisfruta.codsocio from rpartes_variedad, rhisfruta, rhisfruta_entradas "
    SQL = SQL & " where rhisfruta.numalbar = rhisfruta_entradas.numalbar and rhisfruta_entradas.numnotac = rpartes_variedad.numnotac and rpartes_variedad.nroparte in "
    SQL = SQL & "(select rpartes.nroparte from " & cTabla
    If cWhere <> "" Then SQL = SQL & " where " & cWhere
    SQL = SQL & ") "
    SQL = SQL & " union "
    SQL = SQL & " select " & vUsu.Codigo & ", rpartes_variedad.nroparte, rclasifica.numnotac, rclasifica.codsocio from rpartes_variedad, rclasifica "
    SQL = SQL & " where  rclasifica.numnotac = rpartes_variedad.numnotac and rpartes_variedad.nroparte in "
    SQL = SQL & "(select rpartes.nroparte from " & cTabla
    If cWhere <> "" Then SQL = SQL & " where " & cWhere
    SQL = SQL & ") "
    
    conn.Execute SQL
    
    


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
    
    'A�adir el parametro de Empresa
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
                                    '[Monica]29/01/2018: tb catadau
    If vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
        
        cadNombreRPT = nomDocu '"rInfAsesoriaNomiMes.rpt"
        cadTitulo = "Informe de Generacion de N�mina"
'        If Me.Check2.Value = 1 Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "1.rpt") '"rInfAsesoriaNomiMes1.rpt"
        
        If CargarTemporalListNominaCoopic(cadSelect, Fdesde, Fhasta, txtCodigo(78).Text) Then
            Tabla = "{tmpinformes}"
            cadSelect = "{tmpinformes.codusu} = " & vUsu.Codigo
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            CadParam = CadParam & "pDias=" & Day(Fhasta) & "|"
            numParam = numParam + 1
        Else
            Exit Sub
        End If

    Else
        If CargarTemporalListAsesoria(cadSelect, Fdesde, Fhasta) Then
            Tabla = "{tmpinformes}"
            cadSelect = "{tmpinformes.codusu} = " & vUsu.Codigo
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            CadParam = CadParam & "pDias=" & Day(Fhasta) & "|"
            numParam = numParam + 1
        Else
            Exit Sub
        End If
    End If
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(Tabla, cadSelect) Then
        If (vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19) And Me.Check2.Value = 1 Then
            
            If vParamAplic.Cooperativa = 4 Then  ' Alzira
                Shell App.Path & "\nomina.exe /E|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
            Else
                '[Monica]29/01/2018: para el caso de Catadau
                If vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
                    '[Monica]07/02/2017: modificacion para los dados de baja
                    Dim Fec1 As Date
                    Dim mes As Integer
                    Dim Anyo As Integer
                    
                    mes = Me.Combo1(1).ListIndex + 2
                    Anyo = txtCodigo(61).Text
                    If mes > 12 Then
                        mes = 1
                        Anyo = Anyo + 1
                    End If
                    Fec1 = DateAdd("d", -1, CDate("01/" & Format(mes, "00") & "/" & Format(Anyo, "0000")))
                    If txtCodigo(78).Text <> "" Then Fec1 = CDate(txtCodigo(78).Text)
                    
                    If GeneraNominaA3(Fec1) Then
                        '[Monica]04/05/2018: generamos 2 ficheros
                        If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
                            ' cargamos las hojas excel
                            
                            If Dir(App.Path & "\nomina.z") <> "" Then Kill App.Path & "\nomina.z"
            
                            Shell App.Path & "\nomina.exe /R|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
                            
                            While Dir(App.Path & "\nomina.z") = ""
                                Me.Label2(107).Caption = "Procesando Fichero "
                                DoEvents
            
                                espera 1
                            Wend
                            
                            If Dir(App.Path & "\nomina.z") <> "" Then Kill App.Path & "\nomina.z"
                            
                            Shell App.Path & "\nomina.exe /S|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
                            
                            Label2(107).Caption = ""
                            DoEvents
                            Me.Refresh
                        Else
                            If CopiarFicheroA3("NominaA3.txt", CStr(Fec1)) Then
                                MsgBox "Proceso realizado correctamente", vbExclamation
                            End If
                        End If
                    End If
                Else
                    ' Picassent
                    Shell App.Path & "\nomina.exe /P|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
                End If
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
    
    'A�adir el parametro de Empresa
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
    
    'D/H CAPATAZ
    cDesde = Trim(txtCodigo(76).Text)
    cHasta = Trim(txtCodigo(77).Text)
    nDesde = txtNombre(76).Text
    nHasta = txtNombre(77).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{horas.codcapat}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCapataz=""") Then Exit Sub
    End If
    
    
    
    
    Select Case OpcionListado
        Case 28 ' informe de comprobacion
            ConSubInforme = False
        
            cadNombreRPT = "rInfComprobNomi.rpt"
        
            indRPT = 84 ' personalizamos el informe de comprobacion
            
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    
            cadTitulo = "Informe de Comprobaci�n N�minas"
            
            If vParamAplic.Cooperativa = 16 Then
                CadParam = CadParam & "pResumen=" & Check7.Value & "|"
                numParam = numParam + 1
            End If
            
            '[Monica]07/06/2018: para el caso de catadau el informe resumido por dias es otro rpt
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
                If Check7.Value = 1 Then nomDocu = Replace(nomDocu, ".rpt", "1.rpt")
            End If
            
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
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
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
        
    SQL = "select count(distinct horas.codtraba) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Rs.Close
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = "select horas.codtraba,  sum(horasdia), sum(compleme), sum(penaliza), sum(importe) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    SQL = SQL & " group by horas.codtraba "
    SQL = SQL & " order by 1 "
        
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
Dim SQL As String

    If Not DatosOk Then Exit Sub
    
    
    InicializarVbles
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '[Monica]25/05/2018: A�ado Catadau
    If vParamAplic.Cooperativa = 9 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
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
        SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cadSelect <> "" Then
            cadSelect = QuitarCaracterACadena(cadSelect, "{")
            cadSelect = QuitarCaracterACadena(cadSelect, "}")
            cadSelect = QuitarCaracterACadena(cadSelect, "_1")
            SQL = SQL & " WHERE " & cadSelect
        End If
        
        If RegistrosAListar(SQL) = 0 Then
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
        SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cadSelect <> "" Then
            cadSelect = QuitarCaracterACadena(cadSelect, "{")
            cadSelect = QuitarCaracterACadena(cadSelect, "}")
            cadSelect = QuitarCaracterACadena(cadSelect, "_1")
            SQL = SQL & " WHERE " & cadSelect
        End If
        
        If RegistrosAListar(SQL) = 0 Then
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
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Select Case Index
        Case 0 ' Proceso de pago de partes de campo
            NomAlmac = ""
            NomAlmac = DevuelveDesdeBDNew(cAgro, "salmpr", "nomalmac", "codalmac", vParamAplic.AlmacenNOMI, "N")
            If NomAlmac = "" Then
                MsgBox "Debe introducir un c�digo de almac�n de N�minas en par�metros. Revise.", vbExclamation
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
                    '[Monica]23/12/2011: s�lo en el caso de que queramos la prevision
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
                    If vParamAplic.Cooperativa = 2 Then  ' Picassent
                        If ProcesoCargaHorasPicassent(cTabla, cadSelect) Then
                            MsgBox "Proceso realizado correctamente.", vbExclamation
                            cmdCancel_Click (0)
                            Exit Sub
                        Else
                            MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
                            Exit Sub
                        End If
                    Else
                        If vParamAplic.Cooperativa = 16 Then
                            If ProcesoCargaHorasCoopic(cTabla, cadSelect) Then
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
                            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then ' catadau
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
Dim SQL As String
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
Dim SQL As String
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
        MsgBox "El c�digo desde no puede ser superior al c�digo hasta.", vbExclamation
        Exit Sub
    End If
    
    If CalculoEventuales() Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
       
        cmdCancel_Click (2)
    End If

End Sub



Private Sub CmdTrabajadoresActivos_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim cWhere As String
Dim SQL As String

       InicializarVbles
       
        'A�adir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

        CadParam = CadParam & "pFecha=""" & txtCodigo(81).Text & """|"
        numParam = numParam + 1


        cTabla = Tabla
        cWhere = cadSelect
        
        cTabla = QuitarCaracterACadena(cTabla, "{")
        cTabla = QuitarCaracterACadena(cTabla, "}")
        SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            cWhere = QuitarCaracterACadena(cWhere, "{")
            cWhere = QuitarCaracterACadena(cWhere, "}")
            cWhere = QuitarCaracterACadena(cWhere, "_1")
            SQL = SQL & " WHERE " & cWhere
        End If
    
        ' trabajadores en activo
        If TrabajadoresEnActivo(txtCodigo(81).Text) Then
            If HayRegParaInforme("tmpinformes", "{tmpinformes.codusu} = " & vUsu.Codigo) Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                
                cadNombreRPT = "rInfTrabajadoresActivos.rpt"
                cadTitulo = "Trabajadores en Activo"
                ConSubInforme = True
                LlamarImprimir
            Else
                MsgBox "No hay registros entre esos l�mites.", vbExclamation
            End If
        End If

End Sub

Private Function TrabajadoresEnActivo(Fecha As String) As Boolean
Dim SQL As String

    On Error GoTo eTrabajadoresEnActivo

    TrabajadoresEnActivo = False

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL


    SQL = "insert into tmpinformes (codusu, codigo1, nombre1, nombre2) "
    SQL = SQL & "select " & vUsu.Codigo & ", codtraba, nomtraba, niftraba "
    SQL = SQL & " from straba where fechaalta <= " & DBSet(Fecha, "F")
    SQL = SQL & " and fechabaja is null "
    conn.Execute SQL
    
    SQL = "delete from tmpinformes2 where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    ' en tmpinformes2 metemos en que cuadrilla estan
    SQL = "insert into tmpinformes2 (codusu, codigo1, importe1) "
    SQL = SQL & " select codusu, codigo1, rcuadrilla_trabajador.codcuadrilla "
    SQL = SQL & " from  tmpinformes left join rcuadrilla_trabajador on tmpinformes.codigo1 = rcuadrilla_trabajador.codtraba "
    SQL = SQL & " where tmpinformes.codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = "update tmpinformes2,  rcuadrilla, rcapataz      "
    SQL = SQL & " set tmpinformes2.nombre1 = rcapataz.nomcapat "
    SQL = SQL & " where tmpinformes2.codusu = " & vUsu.Codigo
    SQL = SQL & " and tmpinformes2.importe1 = rcuadrilla.codcuadrilla "
    SQL = SQL & " and rcuadrilla.codcapat = rcapataz.codcapat "
    conn.Execute SQL
    
    TrabajadoresEnActivo = True
    
    Exit Function
    
    
eTrabajadoresEnActivo:
    MuestraError Err.Number, "Carga Trabajadores en Activo", Err.Description
End Function


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
    
    'A�adir el parametro de Empresa
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
    cadTitulo = "Informe Mensual D�as Trabajados"
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
            If vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 16 Then ' Alzira o Coopic
                Shell App.Path & "\nomina.exe /E|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
            Else ' Picassent
                Shell App.Path & "\nomina.exe /P|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
            End If
        Else
            LlamarImprimir
        End If
    End If


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
                
                If OpcionListado = 19 Then Label3.Caption = "Actualizaci�n Entradas de Destajo"
                
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
                        Label4.Caption = "Calculo Penalizaci�n"
                        
                    Case 22, 32
                        Me.FrameDestajo.visible = False
                        Me.FramePenalizacion.visible = False
                        Me.FrameBonificacion.visible = True
                        Me.FrameDestajo.Enabled = False
                        Me.FramePenalizacion.Enabled = False
                        Me.FrameBonificacion.Enabled = True
                        Label4.Caption = "Calculo Bonificaci�n"
                            
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
                If vParamAplic.Cooperativa = 16 Then Check7.Value = 1
            
            
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
                
            Case 40 ' Impresion de parte de trabajo
                PonerFoco txtCodigo(74)
                
            Case 41 ' creacion de horas de capataz de servicios
                PonerFoco txtCodigo(82)
                
            Case 42 ' listado de trabajadores activos
                PonerFoco txtCodigo(81)
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
    Me.imgBuscar(36).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    
   
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
    Me.FrameImpresionParte.visible = False
    Me.FrameCapatazServicios.visible = False
    Me.FrameTrabajadoresActivos.visible = False
    
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
    
        '[Monica]07/06/2018: Catadau tb quiere un resumen por trabajador
        Check7.visible = (vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
        Check7.Enabled = (vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
    
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
    
        If vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then Label15.Caption = "Pago N�mina"
    
        FechaBajaVisible vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19
    
        CargaCombo
    
        FrameListMensAsesoriaVisible True, H, W
        indFrame = 0
        Tabla = "rrecasesoria"
        
        If vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then Check2.Caption = "Generar Fichero A3"
    
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

    Case 40 ' Impresion de partes de trabajo
    
        FrameImpresionParteVisible True, H, W
        indFrame = 0
        Tabla = "rpartes"

    Case 41 ' capataz servicios especiales
        FrameCapatazServiciosVisible True, H, W
        indFrame = 0
        Tabla = "rpartes"
        
    Case 42 ' trabajadores activos
        FrameTrabajadoresActivosVisible True, H, W
        indFrame = 0
        Tabla = "straba"
        
        txtCodigo(81).Text = Format(Now, "dd/mm/yyyy")
        
    End Select
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.CmdCancel(0).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub FechaBajaVisible(Mostrar As Boolean)
    Label2(105).visible = Mostrar
    imgFecha(25).visible = Mostrar
    txtCodigo(78).visible = Mostrar
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
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
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
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
            
        Case 31, 32 ' capataz
            AbrirFrmCapataces Index + 41
        
        Case 33, 34 ' capataz
            AbrirFrmCapataces Index + 43

        Case 36 ' capataz
            AbrirFrmCapataces 80
        
        
    End Select
    PonerFoco txtCodigo(indCodigo)
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
    
    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40
    
    Select Case Index
        Case 0, 1, 13, 2, 3
            Indice = Index + 14
        Case 4, 5
            Indice = Index
        Case 7
            Indice = 11
        Case 6
            Indice = 29
        Case 8
            Indice = 30
        Case 9
            Indice = 35
        Case 10
            Indice = 26
        Case 12
            Indice = 37
        Case 11
            Indice = 33
        Case 14
            Indice = 46
        Case 15
            Indice = 44
        Case 16
            Indice = 48
        Case 17, 18
            Indice = Index + 35
        Case 19, 20
            Indice = Index + 37
        Case 21, 22
            Indice = Index + 38
        Case 25
            Indice = 78
        Case 26
            Indice = 82
        Case 28
            Indice = 81
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
        
        
            Case 72: KEYBusqueda KeyAscii, 31 'capataz desde
            Case 73: KEYBusqueda KeyAscii, 32 'capataz hasta
            Case 70: KEYFecha KeyAscii, 23 'fecha desde
            Case 71: KEYFecha KeyAscii, 24 'fecha hasta
        
            Case 76: KEYBusqueda KeyAscii, 33 'capataz desde
            Case 77: KEYBusqueda KeyAscii, 34 'capataz hasta
        
            Case 78: KEYFecha KeyAscii, 25 'fecha de baja del trabajador (coopic)
        
            Case 82: KEYFecha KeyAscii, 26 'fecha de creacion
            Case 80: KEYBusqueda KeyAscii, 36 'capataz hasta
        
            Case 81: KEYFecha KeyAscii, 28 'fecha de trabajadores activos
        
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
            
        Case 35, 26, 33, 37, 46, 44, 48, 52, 53, 56, 57, 59, 60, 70, 71, 78, 82, 81 'FECHAS
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
            
        Case 31, 32, 34, 45, 38, 43, 72, 73, 76, 77, 80 'CAPATAZ
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
        
        Case 74, 75 ' Nro de parte
            PonerFormatoEntero txtCodigo(Index)
        
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
        Me.FrameInfComprobacion.Height = 5085
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
        Me.FrameBorradoMasivoETT.Width = 6585
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
        Me.FramePaseABanco.Height = 5990 '5130
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
        Me.FrameListMensAsesoria.Height = 4575
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


Private Sub FrameImpresionParteVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el pase a banco
    Me.FrameImpresionParte.visible = visible
    If visible = True Then
        Me.FrameImpresionParte.Top = -90
        Me.FrameImpresionParte.Left = 0
        Me.FrameImpresionParte.Height = 5445
        Me.FrameImpresionParte.Width = 6285
        W = Me.FrameImpresionParte.Width
        H = Me.FrameImpresionParte.Height
    End If
End Sub


Private Sub FrameCapatazServiciosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el pase a banco
    Me.FrameCapatazServicios.visible = visible
    If visible = True Then
        Me.FrameCapatazServicios.Top = -90
        Me.FrameCapatazServicios.Left = 0
        Me.FrameCapatazServicios.Height = 3135
        Me.FrameCapatazServicios.Width = 6375
        W = Me.FrameCapatazServicios.Width
        H = Me.FrameCapatazServicios.Height
    End If
End Sub


Private Sub FrameTrabajadoresActivosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el pase a banco
    Me.FrameTrabajadoresActivos.visible = visible
    If visible = True Then
        Me.FrameTrabajadoresActivos.Top = -90
        Me.FrameTrabajadoresActivos.Left = 0
        Me.FrameTrabajadoresActivos.Height = 3135
        Me.FrameTrabajadoresActivos.Width = 6375
        W = Me.FrameTrabajadoresActivos.Width
        H = Me.FrameTrabajadoresActivos.Height
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
        .ConSubInforme = ConSubInforme
        .Opcion = OpcionListado
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmManTraba(Indice As Integer)
    indCodigo = Indice
    Set frmTra = New frmManTraba
    frmTra.DatosADevolverBusqueda = "0|2|"
    frmTra.Show vbModal
    Set frmTra = Nothing
End Sub

Private Sub AbrirFrmManCapataz(Indice As Integer)
    indCodigo = Indice
    Set frmCap = New frmManCapataz
    frmCap.DatosADevolverBusqueda = "0|1|"
    frmCap.Show vbModal
    Set frmCap = Nothing
End Sub

Private Sub AbrirFrmManBanco(Indice As Integer)
    indCodigo = Indice
    
    Set frmBan = New frmBasico2
    
    AyudaBancosCom frmBan, txtCodigo(indCodigo)
    
    Set frmBan = Nothing
    
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub AbrirFrmManAlmac(Indice As Integer)
    indCodigo = Indice + 4
    
    Set frmAlm = New frmBasico2
    
    AyudaAlmacenCom frmAlm, txtCodigo(indCodigo).Text
    
    Set frmAlm = Nothing
    
    PonerFoco txtCodigo(indCodigo)

End Sub


Private Function CargarTablaTemporal() As Boolean
Dim SQL As String
Dim SQL1 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarTablaTemporal
    
    CargarTablaTemporal = False

    SQL = "delete from tmpenvasesret where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL

'select albaran_envase.codartic, albaran_envase.fechamov
'from (albaran_envase inner join sartic on albaran_envase.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar
'Where stipar.esretornable = 1
'Union
'select smoval.codartic, smoval.fechamov
'from (smoval inner join  sartic on smoval.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar
'Where stipar.esretornable = 1

    SQL = "select " & vUsu.Codigo & ", albaran_envase.codartic, albaran_envase.fechamov, albaran_envase.cantidad, albaran_envase.tipomovi, albaran_envase.numalbar, "
    SQL = SQL & " albaran_envase.codclien, clientes.nomclien, " & DBSet("ALV", "T")
    SQL = SQL & " from ((albaran_envase inner join sartic on albaran_envase.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar) "
    SQL = SQL & " inner join clientes on albaran_envase.codclien = clientes.codclien "
    SQL = SQL & " where stipar.esretornable = 1 "
    
    If txtCodigo(12).Text <> "" Then SQL = SQL & " and stipar.codtipar >= " & DBSet(txtCodigo(12).Text, "N")
    If txtCodigo(13).Text <> "" Then SQL = SQL & " and stipar.codtipar <= " & DBSet(txtCodigo(13).Text, "N")
    
    If txtCodigo(20).Text <> "" Then SQL = SQL & " and albaran_envase.codartic >= " & DBSet(txtCodigo(20).Text, "T")
    If txtCodigo(21).Text <> "" Then SQL = SQL & " and albaran_envase.codartic <= " & DBSet(txtCodigo(21).Text, "T")
    
    If txtCodigo(22).Text <> "" Then SQL = SQL & " and albaran_envase.codclien >= " & DBSet(txtCodigo(22).Text, "N")
    If txtCodigo(23).Text <> "" Then SQL = SQL & " and albaran_envase.codclien <= " & DBSet(txtCodigo(23).Text, "N")
    
    If txtCodigo(14).Text <> "" Then SQL = SQL & " and albaran_envase.fechamov >= " & DBSet(txtCodigo(14).Text, "F")
    If txtCodigo(15).Text <> "" Then SQL = SQL & " and albaran_envase.fechamov <= " & DBSet(txtCodigo(15).Text, "F")
    
    SQL = SQL & " union "
    
    SQL = SQL & "select " & vUsu.Codigo & ", smoval.codartic, smoval.fechamov, smoval.cantidad, smoval.tipomovi, smoval.document, "
    SQL = SQL & " smoval.codigope, proveedor.nomprove, " & DBSet("ALC", "T")
    SQL = SQL & " from ((smoval inner join sartic on smoval.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar) "
    SQL = SQL & " inner join proveedor on smoval.codigope = proveedor.codprove "
    SQL = SQL & " where stipar.esretornable = 1 "
    
    If txtCodigo(12).Text <> "" Then SQL = SQL & " and stipar.codtipar >= " & DBSet(txtCodigo(12).Text, "N")
    If txtCodigo(13).Text <> "" Then SQL = SQL & " and stipar.codtipar <= " & DBSet(txtCodigo(13).Text, "N")
    
    If txtCodigo(20).Text <> "" Then SQL = SQL & " and smoval.codartic >= " & DBSet(txtCodigo(20).Text, "T")
    If txtCodigo(21).Text <> "" Then SQL = SQL & " and smoval.codartic <= " & DBSet(txtCodigo(21).Text, "T")
    
    If txtCodigo(14).Text <> "" Then SQL = SQL & " and smoval.fechamov >= " & DBSet(txtCodigo(14).Text, "F")
    If txtCodigo(15).Text <> "" Then SQL = SQL & " and smoval.fechamov <= " & DBSet(txtCodigo(15).Text, "F")

    SQL1 = "insert into tmpenvasesret " & SQL
    conn.Execute SQL1
    
    CargarTablaTemporal = True
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Carga Tabla Temporal"
End Function

Private Function CalculoHorasProductivas() As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim SQL1 As String

    On Error GoTo eCalculoHorasProductivas

    CalculoHorasProductivas = False

    SQL = "fechahora = " & DBSet(txtCodigo(27).Text, "F") & " and codalmac = " & DBSet(txtCodigo(24), "N")
    SQL = SQL & " and codtraba in (select codtraba from straba where codsecci = 1)"


    If BloqueaRegistro("horas", SQL) Then
        SQL1 = "update horas set horasproduc = round(horasdia * (1 + (" & DBSet(txtCodigo(25), "N") & "/ 100)),2) "
        SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(27).Text, "F")
        SQL1 = SQL1 & " and codalmac = " & DBSet(txtCodigo(24), "N")
        SQL1 = SQL1 & " and codtraba in (select codtraba from straba where codsecci = 1) "
        
        conn.Execute SQL1
    
        CalculoHorasProductivas = True
    End If

    TerminaBloquear
    Exit Function

eCalculoHorasProductivas:
    MuestraError Err.Number, "Calculo Horas Productivas", Err.Description
    TerminaBloquear
End Function


Private Function ProcesoCargaHoras(cTabla As String, cWhere As String, EsPrevision As Boolean) As Boolean
Dim SQL As String
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
        SQL = "CARNOM" 'carga de nominas
        'Bloquear para que nadie mas pueda contabilizar
        DesBloqueoManual (SQL)
        If Not BloqueoManual(SQL, "1") Then
            MsgBox "No se puede realizar el proceso de Carga de N�minas. Hay otro usuario realiz�ndolo.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    ProcesoCargaHoras = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    If Not EsPrevision Then
        SQL = "Select rpartes_trabajador.nroparte, rpartes.fechapar, rpartes_trabajador.codtraba, sum(if(rpartes_trabajador.importe is null,0,rpartes_trabajador.importe)) FROM " & QuitarCaracterACadena(cTabla, "_1")
    Else
        SQL = "Select rpartes_trabajador.codtraba, rpartes.fechapar, sum(if(rpartes_trabajador.importe is null,0,rpartes_trabajador.importe)) FROM " & QuitarCaracterACadena(cTabla, "_1")
    End If
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    If Not EsPrevision Then
        SQL = SQL & " group by 1, 2, 3"
        SQL = SQL & " order by 1, 2, 3"
    Else
        SQL = SQL & " group by 1, 2"
        SQL = SQL & " order by 1, 2"
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    
    If EsPrevision Then
        SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
        conn.Execute SQL
        
        '                                       codtraba,fecha,  importe
        SQL = "insert into tmpinformes (codusu, codigo1, fecha1, importe1) values "
    Else
        SQL = "insert into horas (codtraba, fechahora, horasdia, horasproduc, compleme,"
        SQL = SQL & "intconta, pasaridoc, codalmac, nroparte) values "
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
        SQL = SQL & Sql3
        
        conn.Execute SQL
    End If
    
    If Not EsPrevision Then
        DesBloqueoManual ("CARNOM") 'carga de nominas
        
    Else
        
        SQL = "select codigo1, sum(importe1) from tmpinformes where codusu = " & vUsu.Codigo
        SQL = SQL & " group by 1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            
            Sql2 = "select salarios.impsalar, salarios.imphorae, straba.dtosirpf, straba.dtosegso, straba.porc_antig from salarios, straba where straba.codtraba = " & DBSet(Rs!Codigo1, "N")
            Sql2 = Sql2 & " and salarios.codcateg = straba.codcateg "
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            ImpBruto = Round2(DBLet(Rs.Fields(1).Value, "N"), 2)
            
    '        [Monica]23/03/2010: incrementamos el bruto el porcentaje de antig�edad si lo tiene, si no 0
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
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoCargaHorasPicassent
    
    Screen.MousePointer = vbHourglass
    
    SQL = "CARNOM" 'carga de nominas
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se puede realizar el proceso de Carga de N�minas. Hay otro usuario realiz�ndolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    ProcesoCargaHorasPicassent = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = cTabla & " INNER JOIN rcuadrilla ON rpartes.codcuadrilla = rcuadrilla.codcuadrilla "
    SQL = "Select rpartes_trabajador.nroparte, rpartes.fechapar, rpartes_trabajador.codtraba, rpartes_trabajador.codvarie, rcuadrilla.codcapat, sum(rpartes_trabajador.importe), sum(rpartes_trabajador.kilosrec) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5"
    SQL = SQL & " order by 1, 2, 3, 4, 5"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    SQL = "insert into horas (codtraba, fechahora, horasdia, horasproduc, importe,"
    SQL = SQL & "intconta, pasaridoc, codalmac, nroparte, codvarie, codcapat, kilos) values "
        
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
        SQL = SQL & Sql3
        
        conn.Execute SQL
    End If
    
    DesBloqueoManual ("CARNOM") 'carga de nominas
    
    Screen.MousePointer = vbDefault
    
    ProcesoCargaHorasPicassent = True
    Exit Function
    
eProcesoCargaHorasPicassent:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Carga de Horas", Err.Description
End Function

Private Function ProcesoCargaHorasCoopic(cTabla As String, cWhere As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoCargaHorasCoopic
    
    Screen.MousePointer = vbHourglass
    
    SQL = "CARNOM" 'carga de nominas
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se puede realizar el proceso de Carga de N�minas. Hay otro usuario realiz�ndolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    ProcesoCargaHorasCoopic = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = cTabla & " INNER JOIN rcuadrilla ON rpartes.codcuadrilla = rcuadrilla.codcuadrilla "
    SQL = "Select rpartes_trabajador.nroparte, rpartes.fechapar, rpartes_trabajador.codtraba, rpartes_trabajador.codvarie, rcuadrilla.codcapat, rpartes_trabajador.codgasto, sum(rpartes_trabajador.importe), sum(rpartes_trabajador.kilosrec) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6"
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    SQL = "insert into horas (codtraba, fechahora, horasdia, horasproduc, importe,"
    SQL = SQL & "intconta, pasaridoc, codalmac, nroparte, codvarie, codcapat, kilos, codforfait) values "
        
    Sql3 = ""
    While Not Rs.EOF
        Sql2 = "select count(*) from horas where fechahora = " & DBSet(Rs.Fields(1).Value, "F")
        Sql2 = Sql2 & " and codtraba = " & DBSet(Rs.Fields(2).Value, "N")
        Sql2 = Sql2 & " and codalmac = " & DBSet(vParamAplic.AlmacenNOMI, "N")
        Sql2 = Sql2 & " and codvarie = " & DBLet(Rs.Fields(3).Value, "N")
        Sql2 = Sql2 & " and codcapat = " & DBSet(Rs.Fields(4).Value, "N")
        If IsNull(Rs.Fields(5).Value) Then
            Sql2 = Sql2 & " and codforfait = ''"
        Else
            Sql2 = Sql2 & " and codforfait = " & DBLet(Rs.Fields(5).Value, "T")
        End If
        
        
        If TotalRegistros(Sql2) = 0 Then
            Sql3 = Sql3 & "(" & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "F") & ",0,0,"
            Sql3 = Sql3 & DBSet(Rs.Fields(6).Value, "N") & ",0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(3).Value, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(4).Value, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(7).Value, "N") & ","
            
            If IsNull(Rs.Fields(5).Value) Then
                Sql3 = Sql3 & "''),"
            Else
                Sql3 = Sql3 & DBSet(Rs.Fields(5).Value, "N") & "),"
            End If
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If Sql3 <> "" Then
        ' quitamos la ultima coma
        Sql3 = Mid(Sql3, 1, Len(Sql3) - 1)
        SQL = SQL & Sql3
        
        conn.Execute SQL
    End If
    
    DesBloqueoManual ("CARNOM") 'carga de nominas
    
    Screen.MousePointer = vbDefault
    
    ProcesoCargaHorasCoopic = True
    Exit Function
    
eProcesoCargaHorasCoopic:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Carga de Horas", Err.Description
End Function





Private Function ProcesoCargaHorasCatadau(cTabla As String, cWhere As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Almacen As Integer
Dim Sql5 As String
Dim Nregs As Long

Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoCargaHorasCatadau
    
    Screen.MousePointer = vbHourglass
    
    SQL = "CARNOM" 'carga de nominas
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se puede realizar el proceso de Carga de N�minas. Hay otro usuario realiz�ndolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    ProcesoCargaHorasCatadau = False

    Sql5 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql5


    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    cTabla = cTabla & " INNER JOIN rcuadrilla ON rpartes.codcuadrilla = rcuadrilla.codcuadrilla "
    SQL = "Select rpartes_trabajador.nroparte, rpartes.fechapar, rpartes_trabajador.codtraba, rpartes_trabajador.codvarie, rcuadrilla.codcapat, sum(rpartes_trabajador.importe), sum(rpartes_trabajador.kilosrec), sum(rpartes_trabajador.horastra) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5"
    SQL = SQL & " order by 1, 2, 3, 4, 5"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    SQL = "insert into horas (codtraba, fechahora, horasdia, horasproduc, importe,"
    SQL = SQL & "intconta, pasaridoc, codalmac, nroparte, codvarie, codcapat, kilos) values "
        
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
            '[Monica]18/06/2013: solo voy a dejar que el trabajador trabaje ma�ana y tarde
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
        SQL = SQL & Sql3
        
        conn.Execute SQL
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
Dim SQL As String
Dim SQL1 As String
Dim Sql2 As String
Dim cadMen As String
Dim i As Long
Dim B As Boolean
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
    
    B = True
    
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
    SQL = "insert into horas (codtraba, fechahora, horasdia, horasproduc, compleme, horasini, horasfin, "
    SQL = SQL & "anticipo, horasext, fecharec, intconta, pasaridoc, codalmac, nroparte, codvarie, codforfait, "
    SQL = SQL & " numcajon, kilos) "
    SQL = SQL & Sql2
    
    conn.Execute SQL
    
    ' borramos de horasdestajo
    SQL = "delete from horasdestajo "
    If cWhere <> "" Then
        SQL = SQL & " WHERE " & cWhere
    End If
    
    conn.Execute SQL
    
eActualizarTabla:
    If Err.Number Then
        B = False
        MuestraError Err.Number, "Actualizando Horas Destajo", Err.Description & cadMen
    End If
    If B Then
        conn.CommitTrans
        ActualizarTabla = True
    Else
        conn.RollbackTrans
        ActualizarTabla = False
    End If
End Function


Private Sub AbrirFrmVariedades(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmCapataces(Indice As Integer)
    indCodigo = Indice
    Set frmCap = New frmManCapataz
    frmCap.DatosADevolverBusqueda = "0|1|"
    frmCap.CodigoActual = txtCodigo(indCodigo)
    frmCap.Show vbModal
    Set frmCap = Nothing
End Sub



Private Function CalculoDestajoETT(actualiza As Boolean) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim SQL1 As String
Dim Kilos As Long
Dim CodigoETT As Long
Dim Categoria As Long

Dim Precio As Currency
Dim Importe As Currency


    On Error GoTo eCalculoDestajoETT

    CalculoDestajoETT = False

    SQL = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    CodigoETT = DevuelveValor(SQL)

    SQL = "select codcateg from rcapataz left join straba on rcapataz.codtraba = straba.codtraba where rcapataz.codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    Categoria = DevuelveValor(SQL)


    SQL = "select sum(kilostra) from rclasifica where fechaent = " & DBSet(txtCodigo(11).Text, "F") & " and (codvarie = " & DBSet(txtCodigo(9), "N")
    
    '[Monica]11/09/2017: tenemos que traer los kilos de las variedades relacionadas, a�adimos el or
    'SQL = SQL & " or codvarie in (select codvarie1 from variedades_rel where codvarie = " & DBSet(txtCodigo(9), "N") & ")) "
    
    '[Monica]22/12/2017: ahora en las relacionadas hemos de ver si tienen o no el mismo precio de recoleccion
    Dim VRel As String
    VRel = VariedadesRelacionadas(txtCodigo(9).Text)
    If VRel <> "" Then
        SQL = SQL & " or codvarie in ( " & VRel & " )) "
    Else
        SQL = SQL & " ) "
    End If
    SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")

    Kilos = DevuelveValor(SQL)
    
    SQL = "select precio from rtarifaett where codvarie = " & DBSet(txtCodigo(9).Text, "N")
    SQL = SQL & " and codigoett = " & DBSet(CodigoETT, "N")
    
    Precio = DevuelveValor(SQL)
    
    Importe = Round2(Kilos * Precio, 2)
    
    txtCodigo(10).Text = Format(Kilos, "###,###,##0")
    txtCodigo(8).Text = Format(Precio, "###,##0.0000")
    txtCodigo(13).Text = Format(Importe, "###,###,##0.00")

    If Not actualiza Then
        CalculoDestajoETT = True
        Exit Function
    Else
        SQL = "select count(*) from horasett where fechahora = " & DBSet(txtCodigo(11).Text, "F")
        SQL = SQL & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
        SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
        SQL = SQL & " and codigoett = " & DBSet(CodigoETT, "N")
        
        If TotalRegistros(SQL) = 0 Then
            SQL1 = "insert into horasett (fechahora,codvarie,codigoett,codcapat,complemento,codcateg,importe,penaliza,"
            SQL1 = SQL1 & "complcapataz , kilosalicatados, kilostiron, fecharec, intconta, pasaridoc) values ("
            SQL1 = SQL1 & DBSet(txtCodigo(11).Text, "F") & ","
            SQL1 = SQL1 & DBSet(txtCodigo(9).Text, "N") & ","
            SQL1 = SQL1 & DBSet(CodigoETT, "N") & ","
            SQL1 = SQL1 & DBSet(txtCodigo(12).Text, "N") & ","
            SQL1 = SQL1 & "0,"
            SQL1 = SQL1 & DBSet(Categoria, "N") & ","
            SQL1 = SQL1 & DBSet(Importe, "N") & ","
            SQL1 = SQL1 & "0,0,"
            SQL1 = SQL1 & DBSet(Kilos, "N") & ","
            SQL1 = SQL1 & "0,null,0,0) "
            
            conn.Execute SQL1
        Else
            SQL1 = "update horasett set importe = " & DBSet(Importe, "N")
            SQL1 = SQL1 & ", kilosalicatados = " & DBSet(Kilos, "N")
            SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            SQL1 = SQL1 & " and codigoett = " & DBSet(CodigoETT, "N")
            
            conn.Execute SQL1
        End If
        
        CalculoDestajoETT = True
        Exit Function
    End If
    
eCalculoDestajoETT:
    MuestraError Err.Number, "Calculo Destajo ETT", Err.Description
    TerminaBloquear
End Function

Private Function VariedadesRelacionadas(vVarie As String) As String
Dim SQL As String
Dim Precio As Currency
Dim Rs As ADODB.Recordset

    On Error GoTo eVariedadesRelacionadas
    
    
    VariedadesRelacionadas = ""
    
    Precio = DevuelveValor("select eurdesta from variedades where codvarie = " & DBSet(vVarie, "N"))
    SQL = "select codvarie1 from variedades_rel inner join variedades on variedades_rel.codvarie1 = variedades.codvarie where variedades_rel.codvarie = " & DBSet(vVarie, "N") & " and eurdesta= " & DBSet(Precio, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = ""
    
    While Not Rs.EOF
        SQL = SQL & "," & DBSet(Rs!codvarie1, "N")
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If SQL <> "" Then
        VariedadesRelacionadas = Mid(SQL, 2, Len(SQL))
    End If
    Exit Function
    
eVariedadesRelacionadas:
    MuestraError Err.Number, "Variedades Relacionadas", Err.Description
End Function

Private Function CalculoPenalizacionETT(actualiza As Boolean) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim SQL1 As String
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

    SQL = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    CodigoETT = DevuelveValor(SQL)

    SQL = "select sum(kilostra) from rclasifica where fechaent = " & DBSet(txtCodigo(11).Text, "F") & " and (codvarie = " & DBSet(txtCodigo(9), "N")
    '[Monica]11/09/2017: tenemos que traer los kilos de las variedades relacionadas, a�adimos el or
'    Sql = Sql & " or codvarie in (select codvarie1 from variedades_rel where codvarie = " & DBSet(txtCodigo(9), "N") & ")) "
    
    '[Monica]22/12/2017: ahora en las relacionadas hemos de ver si tienen o no el mismo precio de recoleccion
    Dim VRel As String
    VRel = VariedadesRelacionadas(txtCodigo(9).Text)
    If VRel <> "" Then
        SQL = SQL & " or codvarie in ( " & VRel & ")) "
    Else
        SQL = SQL & " ) "
    End If
    
    SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")

    Porcentaje = 0
    If txtCodigo(21).Text <> "" Then Porcentaje = CCur(ImporteSinFormato(txtCodigo(21).Text))


    Kilos = DevuelveValor(SQL)
    KilosTiron = Round2(Kilos * Porcentaje * 0.01, 0)
    
    SQL = "select precio from rtarifaett where codvarie = " & DBSet(txtCodigo(9).Text, "N")
    SQL = SQL & " and codigoett = " & DBSet(CodigoETT, "N")
    
    Precio = DevuelveValor(SQL)
    
    ImporteAlicatado = Round2((Kilos - KilosTiron) * Precio, 2)
    ImporteTotal = Round2(Kilos * Precio, 2)
    Penalizacion = ImporteTotal - ImporteAlicatado
    
    txtCodigo(22).Text = Format(Kilos, "###,###,##0")
    txtCodigo(20).Text = Format(Penalizacion, "###,###,##0.00")

    If Not actualiza Then
        CalculoPenalizacionETT = True
        Exit Function
    Else
        
        SQL1 = "update horasett set  penaliza = " & DBSet(Penalizacion, "N")
        SQL1 = SQL1 & ", kilosalicatados = " & DBSet(Kilos - KilosTiron, "N")
        SQL1 = SQL1 & ", kilostiron = " & DBSet(KilosTiron, "N")
        SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
        SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
        SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
        SQL1 = SQL1 & " and codigoett = " & DBSet(CodigoETT, "N")
        
        conn.Execute SQL1
        
        CalculoPenalizacionETT = True
        Exit Function
    End If
    
eCalculoPenalizacionETT:
    MuestraError Err.Number, "Calculo Penalizacion ETT", Err.Description
End Function
                               

Private Function CalculoBonificacionETT(actualiza As Boolean) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim SQL1 As String
Dim CodigoETT As Long

    On Error GoTo eCalculoBonificacionETT

    CalculoBonificacionETT = False

    SQL = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    CodigoETT = DevuelveValor(SQL)

    SQL1 = "update horasett set  complemento = " & DBSet(txtCodigo(23).Text, "N")
    SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
    SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
    SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
    SQL1 = SQL1 & " and codigoett = " & DBSet(CodigoETT, "N")
    
    conn.Execute SQL1
        
    CalculoBonificacionETT = True
    Exit Function
    
eCalculoBonificacionETT:
    MuestraError Err.Number, "Calculo Bonificacion ETT", Err.Description
End Function
                               


Private Function ProcesoBorradoMasivo(cTabla As String, cWhere As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoBorradoMasivo
    
    Screen.MousePointer = vbHourglass
    
    SQL = "BORMAS" 'BORrado MASivo
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se puede realizar el proceso de Borrado Masivo. Hay otro usuario realiz�ndolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    ProcesoBorradoMasivo = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "delete FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    conn.Execute SQL
        
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
Dim SQL As String
Dim SQL1 As String
Dim Kilos As Long
Dim CodigoETT As Long
Dim Categoria As Long

Dim Precio As Currency
Dim Importe As Currency
Dim i As Integer

Dim Fdesde As Date
Dim Fhasta As Date
Dim Fecha As Date

Dim Trabajador As Long
Dim Dias As Long

    On Error GoTo eCalculoAltaRapida

    CalculoAltaRapida = False

    SQL = "select codtraba from rcapataz where rcapataz.codcapat = " & DBSet(txtCodigo(34).Text, "N")
    
    Trabajador = DevuelveValor(SQL)

    SQL = "select codcateg from straba where codtraba = " & DBSet(Trabajador, "N")

    Categoria = DevuelveValor(SQL)

    Fdesde = CDate(txtCodigo(35).Text)
    Fhasta = CDate(txtCodigo(26).Text)

    Dias = Fhasta - Fdesde

    Importe = 0
    If txtCodigo(40).Text <> "" Then
        Importe = ImporteSinFormato(txtCodigo(40).Text)
    End If

    For i = 0 To Dias
        Fecha = DateAdd("y", i, Fdesde)

        SQL = "select count(*) from horas where fechahora = " & DBSet(Fecha, "F")
        SQL = SQL & " and codvarie = " & DBSet(txtCodigo(36).Text, "N")
        SQL = SQL & " and codcapat = " & DBSet(txtCodigo(34).Text, "N")
        SQL = SQL & " and codtraba = " & DBSet(Trabajador, "N")
        
        If TotalRegistros(SQL) = 0 Then
            SQL1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,importe,fecharec,intconta,pasaridoc,codalmac) values ("
            SQL1 = SQL1 & DBSet(Fecha, "F") & ","
            SQL1 = SQL1 & DBSet(txtCodigo(36).Text, "N") & ","
            SQL1 = SQL1 & DBSet(Trabajador, "N") & ","
            SQL1 = SQL1 & DBSet(txtCodigo(34).Text, "N") & ","
            SQL1 = SQL1 & DBSet(Importe, "N") & ","
            SQL1 = SQL1 & "null,0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ") "
            
            conn.Execute SQL1
        End If
        
    Next i
    
    CalculoAltaRapida = True
    Exit Function
    
eCalculoAltaRapida:
    MuestraError Err.Number, "Calculo Alta R�pida", Err.Description
    TerminaBloquear
End Function



Private Function CalculoEventuales() As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim SQL1 As String
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
        '[Monica]29/10/2014: a�adimos la condicion de que el trabajador que vamos a introducir no tenga fecha de baja
        If TotalRegistros("select count(*) from straba where codtraba = " & J & " and (fechabaja is null or fechabaja = '')") <> 0 Then
    
            For i = 0 To Dias
                Fecha = DateAdd("y", i, Fdesde)
        
                SQL = "select count(*) from horas where fechahora = " & DBSet(Fecha, "F")
                SQL = SQL & " and codvarie = " & DBSet(txtCodigo(28).Text, "N")
                SQL = SQL & " and codcapat = " & DBSet(0, "N")
                SQL = SQL & " and codtraba = " & DBSet(J, "N")
                
                If TotalRegistros(SQL) = 0 Then
                    SQL1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,importe,fecharec,intconta,pasaridoc,codalmac) values ("
                    SQL1 = SQL1 & DBSet(Fecha, "F") & ","
                    SQL1 = SQL1 & DBSet(txtCodigo(28).Text, "N") & ","
                    SQL1 = SQL1 & DBSet(J, "N") & ","
                    SQL1 = SQL1 & "0,"
                    SQL1 = SQL1 & DBSet(Importe, "N") & ","
                    SQL1 = SQL1 & "null,0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ") "
                    
                    conn.Execute SQL1
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
Dim SQL As String
Dim SQL1 As String
Dim Importe As Currency

    On Error GoTo eCalculoTrabajCapataz

    CalculoTrabajCapataz = False
        
    conn.BeginTrans
        
    Importe = 0
    If txtCodigo(51).Text <> "" Then
        Importe = ImporteSinFormato(txtCodigo(51).Text)
    End If

    SQL = "select * from rcuadrilla INNER JOIN rcuadrilla_trabajador ON rcuadrilla.codcuadrilla = rcuadrilla_trabajador.codcuadrilla "
    SQL = SQL & " where rcuadrilla.codcapat = " & DBSet(txtCodigo(45).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        SQL = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(46).Text, "F")
        SQL = SQL & " and codvarie = " & DBSet(txtCodigo(47).Text, "N")
        SQL = SQL & " and codtraba = " & DBSet(Rs!CodTraba, "N")
        SQL = SQL & " and codcapat = " & DBSet(txtCodigo(45).Text, "N")
        
        If TotalRegistros(SQL) = 0 Then
            SQL1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,importe,compleme, fecharec,intconta,pasaridoc,codalmac) values ("
            SQL1 = SQL1 & DBSet(txtCodigo(46).Text, "F") & ","
            SQL1 = SQL1 & DBSet(txtCodigo(47).Text, "N") & ","
            SQL1 = SQL1 & DBSet(Rs!CodTraba, "N") & ","
            SQL1 = SQL1 & DBSet(txtCodigo(45).Text, "N") & ",null, "
            SQL1 = SQL1 & DBSet(Importe, "N") & ","
            SQL1 = SQL1 & "null,0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ") "
            
            conn.Execute SQL1
        Else
            SQL1 = "update horas set compleme = if(compleme is null,0,compleme) + " & DBSet(Importe, "N")
            SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(46).Text, "F")
            SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(47).Text, "N")
            SQL1 = SQL1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
            SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(45).Text, "N")
        
            conn.Execute SQL1
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
Dim SQL As String
Dim SQL1 As String
Dim Importe As Currency

    On Error GoTo eCalculoTrabajCapatazNew

    CalculoTrabajCapatazNew = False
        
    conn.BeginTrans
        
    Importe = 0
    If txtCodigo(51).Text <> "" Then
        Importe = ImporteSinFormato(txtCodigo(51).Text)
    End If

    SQL = "select * from horas "
    SQL = SQL & " where horas.codcapat = " & DBSet(txtCodigo(45).Text, "N")
    SQL = SQL & " and horas.fechahora = " & DBSet(txtCodigo(46).Text, "F")
    SQL = SQL & " and horas.codvarie = " & DBSet(txtCodigo(47).Text, "N")
    
    If TotalRegistrosConsulta(SQL) <> 0 Then
    
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            SQL = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(46).Text, "F")
            SQL = SQL & " and codvarie = " & DBSet(txtCodigo(47).Text, "N")
            SQL = SQL & " and codtraba = " & DBSet(Rs!CodTraba, "N")
            SQL = SQL & " and codcapat = " & DBSet(txtCodigo(45).Text, "N")
            
            If TotalRegistros(SQL) = 0 Then
                SQL1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,importe,compleme, fecharec,intconta,pasaridoc,codalmac) values ("
                SQL1 = SQL1 & DBSet(txtCodigo(46).Text, "F") & ","
                SQL1 = SQL1 & DBSet(txtCodigo(47).Text, "N") & ","
                SQL1 = SQL1 & DBSet(Rs!CodTraba, "N") & ","
                SQL1 = SQL1 & DBSet(txtCodigo(45).Text, "N") & ",null, "
                SQL1 = SQL1 & DBSet(Importe, "N") & ","
                SQL1 = SQL1 & "null,0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ") "
                
                conn.Execute SQL1
            Else
                SQL1 = "update horas set compleme = if(compleme is null,0,compleme) + " & DBSet(Importe, "N")
                SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(46).Text, "F")
                SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(47).Text, "N")
                SQL1 = SQL1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
                SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(45).Text, "N")
            
                conn.Execute SQL1
            
                SQL1 = "update horas set compleme = if(compleme=0,null,compleme) "
                SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(46).Text, "F")
                SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(47).Text, "N")
                SQL1 = SQL1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
                SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(45).Text, "N")
            
                conn.Execute SQL1
            
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
Dim SQL As String
Dim SQL1 As String
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

    SQL = "select codcuadrilla from rcuadrilla where codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    Cuadrilla = DevuelveValor(SQL)

    SQL = "select count(*) from rcuadrilla_trabajador, rcuadrilla where rcuadrilla.codcapat = " & DBSet(txtCodigo(12).Text, "N")
    SQL = SQL & " and rcuadrilla.codcuadrilla = rcuadrilla_trabajador.codcuadrilla"
    
    Nregs = DevuelveValor(SQL)
    
    If Nregs <> 0 Then
        SQL = "select sum(kilostra) from rclasifica where fechaent = " & DBSet(txtCodigo(11).Text, "F") & " and (codvarie = " & DBSet(txtCodigo(9), "N")
        '[Monica]11/09/2017: tenemos que traer los kilos de las variedades relacionadas, a�adimos el or
'        SQL = SQL & " or codvarie in (select codvarie1 from variedades_rel where codvarie = " & DBSet(txtCodigo(9).Text, "N") & ")) "
        
        '[Monica]29/12/2017: ahora en las relacionadas hemos de ver si tienen o no el mismo precio de recoleccion
        Dim VRel As String
        VRel = VariedadesRelacionadas(txtCodigo(9).Text)
        If VRel <> "" Then
            SQL = SQL & " or codvarie in ( " & VRel & " )) "
        Else
            SQL = SQL & " ) "
        End If
        
        SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
        Kilos = DevuelveValor(SQL)
        
        SQL = "select eurdesta from variedades where codvarie = " & DBSet(txtCodigo(9).Text, "N")
        
        Precio = DevuelveValor(SQL)
        
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
            
            SQL = "select codtraba from rcuadrilla_trabajador , rcuadrilla where rcuadrilla.codcapat = " & DBSet(txtCodigo(12).Text, "N")
            SQL = SQL & " and rcuadrilla.codcuadrilla = rcuadrilla_trabajador.codcuadrilla"
            
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            While Not Rs.EOF
                SQL = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(11).Text, "F")
                SQL = SQL & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
                SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
                SQL = SQL & " and codtraba = " & DBSet(Rs!CodTraba, "N")
                
                If TotalRegistros(SQL) = 0 Then
                    SQL1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,compleme,importe,penaliza,"
                    SQL1 = SQL1 & "kilos, fecharec, intconta, pasaridoc,codalmac) values ("
                    SQL1 = SQL1 & DBSet(txtCodigo(11).Text, "F") & ","
                    SQL1 = SQL1 & DBSet(txtCodigo(9).Text, "N") & ","
                    SQL1 = SQL1 & DBSet(Rs!CodTraba, "N") & ","
                    SQL1 = SQL1 & DBSet(txtCodigo(12).Text, "N") & ","
                    SQL1 = SQL1 & "0,"
                    SQL1 = SQL1 & DBSet(ImporteTrab, "N") & ","
                    SQL1 = SQL1 & "0,"
                    SQL1 = SQL1 & DBSet(KilosTrab, "N") & ","
                    SQL1 = SQL1 & "null,0,0, "
                    SQL1 = SQL1 & vParamAplic.AlmacenNOMI & ") "
                    
                    conn.Execute SQL1
                Else
                    SQL1 = "update horas set importe = " & DBSet(ImporteTrab, "N")
                    SQL1 = SQL1 & ", kilos = " & DBSet(KilosTrab, "N")
                    SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
                    SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
                    SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
                    SQL1 = SQL1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
                    
                    conn.Execute SQL1
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
Dim SQL As String
Dim SQL1 As String
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

    SQL = "select codigoett from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    CodigoETT = DevuelveValor(SQL)

    SQL = "select sum(kilostra) from rclasifica where fechaent = " & DBSet(txtCodigo(11).Text, "F") & " and (codvarie = " & DBSet(txtCodigo(9), "N")
    '[Monica]11/09/2017: tenemos que traer los kilos de las variedades relacionadas, a�adimos el or
'    SQL = SQL & " or codvarie in (select codvarie1 from variedades_rel where codvarie = " & DBSet(txtCodigo(9), "N") & ")) "
    
    '[Monica]29/12/2017: ahora en las relacionadas hemos de ver si tienen o no el mismo precio de recoleccion
    Dim VRel As String
    VRel = VariedadesRelacionadas(txtCodigo(9).Text)
    If VRel <> "" Then
        SQL = SQL & " or codvarie in ( " & VRel & " )) "
    Else
        SQL = SQL & " ) "
    End If
    
    SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")

    Porcentaje = 0
    If txtCodigo(21).Text <> "" Then Porcentaje = CCur(ImporteSinFormato(txtCodigo(21).Text))


    Kilos = DevuelveValor(SQL)
    KilosTiron = Round2(Kilos * Porcentaje * 0.01, 0)
    
    '[Monica]06/10/2011: antes era eurhaneg
    SQL = "select eurdesta from variedades where codvarie = " & DBSet(txtCodigo(9).Text, "N")
    
    Precio = DevuelveValor(SQL)
    
    ImporteAlicatado = Round2((Kilos - KilosTiron) * Precio, 2)
    ImporteTotal = Round2(Kilos * Precio, 2)
    Penalizacion = ImporteTotal - ImporteAlicatado
    
    SQL = "select codtraba from horas where fechahora = " & DBSet(txtCodigo(11).Text, "F")
    SQL = SQL & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
    SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    NumTrab = TotalRegistrosConsulta(SQL)
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
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
        
            SQL1 = "update horas set  penaliza = " & DBSet(PenalizacionTrab, "N")
            SQL1 = SQL1 & ", kilos = " & DBSet(KilosTrab, "N")
            SQL1 = SQL1 & ", kilostiron = " & DBSet(KilosTironTrab, "N")
            SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            SQL1 = SQL1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
            
            conn.Execute SQL1
        
            Rs.MoveNext
        
        Wend
        
        If PenalizacionDif <> 0 Or KilosDif <> 0 Or KilosTironDif <> 0 Then
            TrabCapataz = DevuelveValor("select codtraba from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N"))
            
            SQL1 = "update horas set penaliza = penaliza + " & DBSet(PenalizacionDif, "N")
            SQL1 = SQL1 & ", kilos = kilos + " & DBSet(KilosDif, "N")
            SQL1 = SQL1 & ", kilostiron = kilostiron + " & DBSet(KilosTironDif, "N")
            SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
            SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
            SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
            SQL1 = SQL1 & " and codtraba = " & DBSet(TrabCapataz, "N")
            
            conn.Execute SQL1
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
Dim SQL As String
Dim SQL1 As String
Dim Bonif As Currency
Dim NumTrab As Long

Dim BonifTrab As Currency
Dim BonifDif As Currency
Dim TrabCapataz As Long

    On Error GoTo eCalculoBonificacion

    CalculoBonificacion = False

    SQL = "select codtraba from horas where fechahora = " & DBSet(txtCodigo(11).Text, "F")
    SQL = SQL & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
    SQL = SQL & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
    
    NumTrab = TotalRegistrosConsulta(SQL)
    
    Bonif = CCur(ImporteSinFormato(txtCodigo(23).Text))
    BonifTrab = 0
    If NumTrab <> 0 Then BonifTrab = Round2(Bonif / NumTrab, 2)
    
    BonifDif = Bonif - Round2(BonifTrab * NumTrab, 2)
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        SQL1 = "update horas set  compleme = " & DBSet(BonifTrab, "N")
        SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
        SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
        SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
        SQL1 = SQL1 & " and codtraba = " & DBSet(Rs!CodTraba, "N")
        
        conn.Execute SQL1
        
        Rs.MoveNext
    Wend
    
    If BonifDif <> 0 Then
        TrabCapataz = DevuelveValor("select codtraba from rcapataz where codcapat = " & DBSet(txtCodigo(12).Text, "N"))
    
        SQL1 = "update horas set  complemen = " & DBSet(BonifDif, "N")
        SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(11).Text, "F")
        SQL1 = SQL1 & " and codvarie = " & DBSet(txtCodigo(9).Text, "N")
        SQL1 = SQL1 & " and codcapat = " & DBSet(txtCodigo(12).Text, "N")
        SQL1 = SQL1 & " and codtraba = " & DBSet(TrabCapataz, "N")
        
        conn.Execute SQL1
    
    End If
        
    Set Rs = Nothing
    
    CalculoBonificacion = True
    Exit Function
    
eCalculoBonificacion:
    MuestraError Err.Number, "C�lculo Bonificacion", Err.Description
End Function


Private Function ProcesoEntradasCapataz(cTabla As String, cWhere As String) As Boolean
Dim SQL As String
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

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "select rentradas.codcapat, rentradas.fechaent, rentradas.codvarie, sum(rentradas.numcajo1) as cajon, sum(rentradas.kilostra) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rentradas")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE  " & Replace(Replace(cWhere, "horas", "rentradas"), "fechahora", "fechaent")
    End If
    SQL = SQL & " group by 1,2,3 "
    SQL = SQL & " union "
    
    SQL = SQL & "select rclasifica.codcapat, rclasifica.fechaent, rclasifica.codvarie, sum(rclasifica.numcajon) as cajon, sum(rclasifica.kilostra) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rclasifica")
    If cWhere <> "" Then
        SQL = SQL & " WHERE  " & Replace(Replace(cWhere, "horas", "rclasifica"), "fechahora", "fechaent")
    Else
        SQL = SQL & " WHERE (1=1) "
    End If
    '[Monica]11/09/2017
    SQL = SQL & " and not rclasifica.codvarie in (select codvarie1 from variedades_rel)"
    SQL = SQL & " group by 1, 2, 3 "
    
    
    '[Monica]11/09/2017
    SQL = SQL & " union "
    SQL = SQL & "select rclasifica.codcapat, rclasifica.fechaent, "
    '[Monica]21/12/2017: para el caso de las variedades de distinto precio de recoleccion (caso de picassent)
    SQL = SQL & " if(vrel.eurdesta = vvar.eurdesta, variedades_rel.codvarie, variedades_rel.codvarie1) codvarie, "
    SQL = SQL & " sum(rclasifica.numcajon) as cajon, sum(rclasifica.kilostra) as kilos from (((" & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rclasifica")
    SQL = SQL & ") inner join variedades_rel on rclasifica.codvarie = variedades_rel.codvarie1) inner join variedades vrel on vrel.codvarie = variedades_rel.codvarie1) inner join variedades vvar on vvar.codvarie = variedades_rel.codvarie  "
    
    If cWhere <> "" Then
        SQL = SQL & " WHERE " & Replace(Replace(cWhere, "horas", "rclasifica"), "fechahora", "fechaent")
    End If
    SQL = SQL & " group by 1, 2, 3"
    
    
'    Sql = Sql & " union "
'
'    Sql = Sql & "select rhisfruta_entradas.codcapat, rhisfruta_entradas.fechaent, rhisfruta.codvarie, sum(rhisfruta_entradas.numcajon) as cajon, sum(rhisfruta_entradas.kilostra) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rhisfruta_entradas")
'    Sql = Sql & " INNER JOIN rhisfruta ON rhisfruta_entradas.numalbar = rhisfruta.numalbar "
'    If cWhere <> "" Then
'        Sql = Sql & " WHERE " & Replace(Replace(cWhere, "horas", "rhisfruta_entradas"), "fechahora", "fechaent")
'    End If
'    Sql = Sql & " group by 1, 2, 3 "
    
    
    SQL = SQL & " order by 1, 2, 3 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        VarieAnt = DBLet(Rs!Codvarie, "N")
        CapatAnt = DBLet(Rs!codcapat, "N")
        FechaAnt = DBLet(Rs!FechaEnt, "F")
        
        TotCajon = 0
        TotKilos = 0
    End If
    Sql2 = ""
    Nregs = 0
                                        '   capataz,fecha,  variedad, numcajon, kilos
    SQL = "insert into tmpinformes (codusu, campo1, fecha1, importe1, importe2, importe3) values  "
    While Not Rs.EOF
        If DBLet(Rs!codcapat, "N") <> CapatAnt Or DBLet(Rs!FechaEnt, "F") <> FechaAnt Or DBLet(Rs!Codvarie, "N") <> VarieAnt Then
            Sql2 = Sql2 & "( " & vUsu.Codigo & "," & DBSet(CapatAnt, "N") & "," & DBSet(FechaAnt, "F") & "," & DBSet(VarieAnt, "N") & ","
            Sql2 = Sql2 & DBSet(TotCajon, "N") & "," & DBSet(TotKilos, "N") & "),"
        
            VarieAnt = DBLet(Rs!Codvarie, "N")
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
    
        conn.Execute SQL & Sql2
    End If
    
  
                'capataz, fecha,  variedad
    SQL = "select campo1, fecha1, importe1 from tmpinformes where codusu = " & vUsu.Codigo & " order by 1,2,3"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        SQL = "select if(isnull(sum(importe)),0,sum(importe)) + if(isnull(sum(compleme)),0,sum(compleme)) - if(isnull(sum(penaliza)),0,sum(penaliza)) as importe "
        SQL = SQL & " from horas where codcapat = " & DBSet(Rs!campo1, "N")
        SQL = SQL & " and fechahora = " & DBSet(Rs!fecha1, "F")
        SQL = SQL & " and codvarie = " & DBSet(Rs!importe1, "N")
    
        Importe = DevuelveValor(SQL)
        ImporteTot = Importe
        
        CodigoETT = DevuelveValor("select codigoett from rcapataz where codcapat = " & DBSet(Rs!campo1, "N"))
         
        ' si es ett tendr� registros en horasett
        SQL = "select if(isnull(sum(importe)),0,sum(importe)) + if(isnull(sum(complemento)),0,sum(complemento)) - if(isnull(sum(penaliza)),0,sum(penaliza)) "
        SQL = SQL & " from horasett where codcapat = " & DBSet(Rs!campo1, "N")
        SQL = SQL & " and fechahora = " & DBSet(Rs!fecha1, "F")
        SQL = SQL & " and codvarie = " & DBSet(Rs!importe1, "N")
        SQL = SQL & " and codigoett = " & DBSet(CodigoETT, "N")
        
        Importe = DevuelveValor(SQL)
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
Dim SQL As String
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

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    '[Monica]05/02/2014: solo lo cambio para Picassent, para el resto lo dejo como estaba
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        SQL = "select rentradas.codcapat, rentradas.fechaent, rentradas.codvarie, sum("
        If vParamAplic.EsCaja1 Then SQL = SQL & "+coalesce(rentradas.numcajo1,0)"
        If vParamAplic.EsCaja2 Then SQL = SQL & "+coalesce(rentradas.numcajo2,0)"
        If vParamAplic.EsCaja3 Then SQL = SQL & "+coalesce(rentradas.numcajo3,0)"
        If vParamAplic.EsCaja4 Then SQL = SQL & "+coalesce(rentradas.numcajo4,0)"
        If vParamAplic.EsCaja5 Then SQL = SQL & "+coalesce(rentradas.numcajo5,0)"
        
        SQL = SQL & ") as cajon, sum(rentradas.kilosnet) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rentradas")
    Else
        SQL = "select rentradas.codcapat, rentradas.fechaent, rentradas.codvarie, sum(rentradas.numcajo1) as cajon, sum(rentradas.kilosnet) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rentradas")
    End If
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & Replace(Replace(cWhere, "horas", "rentradas"), "fechahora", "fechaent")
    End If
    SQL = SQL & " group by 1,2,3 "
    SQL = SQL & " union "
    
    SQL = SQL & "select rclasifica.codcapat, rclasifica.fechaent, rclasifica.codvarie, sum(rclasifica.numcajon) as cajon, sum(rclasifica.kilosnet) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rclasifica")
    If cWhere <> "" Then
        SQL = SQL & " WHERE " & Replace(Replace(cWhere, "horas", "rclasifica"), "fechahora", "fechaent")
    End If
    SQL = SQL & " group by 1, 2, 3 "
    SQL = SQL & " union "

    SQL = SQL & "select rhisfruta_entradas.codcapat, rhisfruta_entradas.fechaent, rhisfruta.codvarie, sum(rhisfruta_entradas.numcajon) as cajon, sum(rhisfruta_entradas.kilosnet) as kilos from " & Replace(QuitarCaracterACadena(cTabla, "_1"), "horas", "rhisfruta_entradas")
    SQL = SQL & " INNER JOIN rhisfruta ON rhisfruta_entradas.numalbar = rhisfruta.numalbar "
    If cWhere <> "" Then
        SQL = SQL & " WHERE " & Replace(Replace(cWhere, "horas", "rhisfruta_entradas"), "fechahora", "fechaent")
    End If
    SQL = SQL & " group by 1, 2, 3 "
    
    SQL = SQL & " order by 1, 2, 3 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        VarieAnt = DBLet(Rs!Codvarie, "N")
        CapatAnt = DBLet(Rs!codcapat, "N")
        FechaAnt = DBLet(Rs!FechaEnt, "F")
        
        TotCajon = 0
        TotKilos = 0
    End If
    Sql2 = ""
    Nregs = 0
                                        '   capataz,fecha,  variedad, numcajon, kilos
    SQL = "insert into tmpinformes (codusu, campo1, fecha1, importe1, importe2, importe3) values  "
    While Not Rs.EOF
        If DBLet(Rs!codcapat, "N") <> CapatAnt Or DBLet(Rs!FechaEnt, "F") <> FechaAnt Or DBLet(Rs!Codvarie, "N") <> VarieAnt Then
            Sql2 = Sql2 & "( " & vUsu.Codigo & "," & DBSet(CapatAnt, "N") & "," & DBSet(FechaAnt, "F") & "," & DBSet(VarieAnt, "N") & ","
            Sql2 = Sql2 & DBSet(TotCajon, "N") & "," & DBSet(TotKilos, "N") & "),"
        
            VarieAnt = DBLet(Rs!Codvarie, "N")
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
    
        conn.Execute SQL & Sql2
    End If
    
  
                'capataz, fecha,  variedad
    SQL = "select campo1, fecha1, importe1 from tmpinformes where codusu = " & vUsu.Codigo & " order by 1,2,3"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        SQL = "select if(isnull(sum(importe)),0,sum(importe)) + if(isnull(sum(compleme)),0,sum(compleme)) - if(isnull(sum(penaliza)),0,sum(penaliza)) as importe "
        SQL = SQL & " from horas where codcapat = " & DBSet(Rs!campo1, "N")
        SQL = SQL & " and fechahora = " & DBSet(Rs!fecha1, "F")
        SQL = SQL & " and codvarie = " & DBSet(Rs!importe1, "N")
    
        Importe = DevuelveValor(SQL)
        ImporteTot = Importe
        
'        CodigoETT = DevuelveValor("select codigoett from rcapataz where codcapat = " & DBSet(Rs!campo1, "N"))
'
'        ' si es ett tendr� registros en horasett
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
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
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
    
    SQL = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
    
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    SQL = "select count(distinct rrecasesoria.codtraba) from (rrecasesoria inner join straba on rrecasesoria.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb1.visible = True
    CargarProgres Pb1, Rs.Fields(0).Value
    
    Rs.Close
    
    SQL = "select rrecasesoria.codtraba, sum(importe) importe from (rrecasesoria inner join straba on rrecasesoria.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    SQL = SQL & " group by rrecasesoria.codtraba "
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
    SQL = "select codbanco, codsucur, digcontr, cuentaba, codorden34, sufijoem, iban from banpropi where codbanpr = " & DBSet(txtCodigo(58).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    Extra = ""
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            cad = ""
        Else
            '[Monica]22/11/2013: iban
            cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
        Extra = DBLet(Rs!sufijoem, "T") & "|" & vParam.NombreEmpresa & "|"
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = cad
    
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
            B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", "Pago N�mina", Combo1(0).ListIndex, CodigoOrden34)
        Else
            B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", "Pago N�mina", Combo1(0).ListIndex, CodigoOrden34)
        End If
    Else
        B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, "Pago N�mina", CodigoOrden34, Combo1(0).ListIndex)
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
     
'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, "Pago N�mina", CodigoOrden34, Combo1(0).ListIndex)
    If B Then
        B = CopiarFichero
        If B Then
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
            
            If MsgBox("�Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                SQL = "update rrecasesoria, straba, forpago set rrecasesoria.idconta = 1 where rrecasesoria.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWHERE
                conn.Execute SQL
            End If
        End If
    End If

eProcesoPaseABanco:
    If Err.Number <> 0 Then
        Mens = Err.Description
        B = False
    End If
    If B Then
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
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPs = False
    
    SQL = "CREATE TEMPORARY TABLE tmpImpor ( "
    SQL = SQL & "codtraba int(6) unsigned NOT NULL default '0',"
    SQL = SQL & "importe decimal(12,2)  NOT NULL default '0')"
    
    conn.Execute SQL
    
    SQL = "CREATE TEMPORARY TABLE tmpImporNeg ( "
    SQL = SQL & "codtraba int(6) unsigned NOT NULL default '0',"
    SQL = SQL & "concepto varchar(30),"
    SQL = SQL & "importe decimal(12,2)  NOT NULL default '0')"
    
    conn.Execute SQL
     
    CrearTMPs = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPs = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpImpor;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS tmpImporNeg;"
        conn.Execute SQL
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
Dim B As Boolean
Dim SQL As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    B = True
    
    If txtCodigo(59).Text = "" Or txtCodigo(60).Text = "" Then
        SQL = "Debe introducir obligatoriamente un valor en los campos de fecha. Reintroduzca. " & vbCrLf & vbCrLf
        MsgBox SQL, vbExclamation
        B = False
        PonerFoco txtCodigo(59)
    End If
    If B Then
        If txtCodigo(58).Text = "" Then
            SQL = "Debe introducir obligatoriamente un valor en el banco. Reintroduzca. " & vbCrLf & vbCrLf
            MsgBox SQL, vbExclamation
            B = False
            PonerFoco txtCodigo(58)
        End If
    End If
    '[Monica]18/09/2013: debe introducir el concepto
    If B And vParamAplic.Cooperativa = 9 Then
        If txtCodigo(66).Text = "" Then
            SQL = "Debe introducir obligatoriamente una descripci�n. Reintroduzca. " & vbCrLf & vbCrLf
            MsgBox SQL, vbExclamation
            B = False
            PonerFoco txtCodigo(66)
        End If
    End If
        
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

' ********* si n'hi han combos a la cap�alera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I
    Combo1(0).Clear
    
    Combo1(0).AddItem "N�mina"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Pensi�n"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Otros Conceptos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    Combo1(1).Clear
    
    Combo1(1).AddItem "Enero"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Febrero"
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
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
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
        
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = "select distinct horas.codtraba, fechahora, sum(if(importe is null,0,importe) + if(compleme is null,0,compleme) - if(penaliza is null,0,penaliza)) importe from horas where " & cadWHERE
    SQL = SQL & " group by 1, 2 "
    SQL = SQL & " having sum(if(importe is null,0,importe) + if(compleme is null,0,compleme) - if(penaliza is null,0,penaliza)) <> 0 "
    SQL = SQL & " order by 1, 2 "
        
    Set Rs = New ADODB.Recordset
        
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
            SQL = "select sum(neto34) from rrecibosnomina where codtraba = " & DBSet(AntTraba, "N")
            SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
            SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
            
            '[Monica]04/11/2016: y que no haya sido embargado
            SQL = SQL & " and hayembargo = 0 "
                                                
            Anticipado = DevuelveValor(SQL)
                                                
            SQL = "select sum(importe) from rrecibosnomina where codtraba = " & DBSet(AntTraba, "N")
            SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
            SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
                                                
            Bruto = DevuelveValor(SQL)
                                                
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
'            v_cadena = Mid(v_cadena, 1, I - 1) & Replace(v_cadena, "N", "S", I, 1)
            v_cadena = Mid(v_cadena, 1, i - 1) & "S" & Mid(v_cadena, i + 1)
        End If
        Dias = Dias + 1
        
        Anticipado = Anticipado + DBLet(Rs!Importe, "N")
        
        Rs.MoveNext
    Wend
    If HayReg = 1 Then
        ' calculamos el importe anticipado de lo que tenemos guardado en rrecibosnomina
        SQL = "select sum(neto34) from rrecibosnomina where codtraba = " & DBSet(ActTraba, "N")
        SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
        SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
        '[Monica]04/11/2016: y que no haya sido embargado
        SQL = SQL & " and hayembargo = 0 "
                                            
        Anticipado = DevuelveValor(SQL)
                                            
        SQL = "select sum(importe) from rrecibosnomina where codtraba = " & DBSet(ActTraba, "N")
        SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
        SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
                                            
        Bruto = DevuelveValor(SQL)
        
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
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
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
    
    SQL = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
    
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    SQL = "select count(distinct horasanticipos.codtraba) from (horasanticipos inner join straba on horasanticipos.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb1.visible = True
    CargarProgres Pb1, Rs.Fields(0).Value
    
    Rs.Close
    
    SQL = "select horasanticipos.codtraba, sum(importe) importe from (horasanticipos inner join straba on horasanticipos.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    SQL = SQL & " group by horasanticipos.codtraba "
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
    SQL = "select codbanco, codsucur, digcontr, cuentaba, codorden34, sufijoem, iban from banpropi where codbanpr = " & DBSet(txtCodigo(58).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    Extra = ""
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            cad = ""
        Else
            '[Monica]22/11/2013: iban
            cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
        Extra = DBLet(Rs!sufijoem, "T") & "|" & vParam.NombreEmpresa & "|"
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = cad
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
            B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", txtCodigo(66).Text, Combo1(0).ListIndex, CodigoOrden34)
        Else
            B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", txtCodigo(66).Text, Combo1(0).ListIndex, CodigoOrden34)
        End If
    Else
        B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, txtCodigo(66).Text, CodigoOrden34, Combo1(0).ListIndex)
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
     
'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, txtCodigo(66).Text, CodigoOrden34, Combo1(0).ListIndex)
    If B Then
        B = CopiarFichero
        If B Then
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
            
            If MsgBox("�Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                SQL = "update horasanticipos, straba, forpago set horasanticipos.fechapago = " & DBSet(txtCodigo(60).Text, "F")
                SQL = SQL & ", concepto = " & DBSet(Trim(txtCodigo(66).Text), "T")
                SQL = SQL & " where horasanticipos.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWHERE
                conn.Execute SQL
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
        B = False
    End If
    If B Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (0)
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub



Private Function CargarTemporalListDiasTrabajados(cadWHERE As String, Fdesde As Date, Fhasta As Date, cadWHERE2 As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
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
    
    
        
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = "select codtraba, fecha from ("
    
    SQL = SQL & "select distinct rpartes_trabajador.codtraba, rpartes.fecentrada fecha, sum(coalesce(rpartes_trabajador.importe,0)) from rpartes inner join rpartes_trabajador on rpartes.nroparte = rpartes_trabajador.nroparte where " & cadWHERE
    SQL = SQL & " group by 1, 2 "
    SQL = SQL & " having sum(coalesce(rpartes_trabajador.importe,0)) <> 0 "
    
    SQL = SQL & " union "
    SQL = SQL & "select distinct horas.codtraba, horas.fechahora fecha , sum(if(importe is null,0,importe) + if(compleme is null,0,compleme) - if(penaliza is null,0,penaliza)) importe from  horas where " & cadWHERE2
    SQL = SQL & " group by 1, 2 "
    SQL = SQL & " having  sum(if(importe is null,0,importe) + if(compleme is null,0,compleme) - if(penaliza is null,0,penaliza)) <> 0 "
    SQL = SQL & " order by 1, 2 "
    
    SQL = SQL & ") aaaaaa "
    
    SQL = SQL & " order by 1, 2 "
        
    Set Rs = New ADODB.Recordset
        
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
            SQL = "select sum(neto34) from rrecibosnomina where codtraba = " & DBSet(AntTraba, "N")
            SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
            SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
                                                
            Anticipado = 0 'DevuelveValor(Sql)
                                                
            SQL = "select sum(importe) from rrecibosnomina where codtraba = " & DBSet(AntTraba, "N")
            SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
            SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
                                                
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
            '[Monica]01/06/2018
            v_cadena = Mid(v_cadena, 1, i - 1) & "S" & Mid(v_cadena, i + 1) '& Replace(v_cadena, "N", "S", I, 1)
        End If
        Dias = Dias + 1
        
'        Anticipado = Anticipado + DBLet(Rs!Importe, "N")
        
        Rs.MoveNext
    Wend
    If HayReg = 1 Then
        ' calculamos el importe anticipado de lo que tenemos guardado en rrecibosnomina
        SQL = "select sum(neto34) from rrecibosnomina where codtraba = " & DBSet(ActTraba, "N")
        SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
        SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
                                            
        Anticipado = 0 'DevuelveValor(Sql)
                                            
        SQL = "select sum(importe) from rrecibosnomina where codtraba = " & DBSet(ActTraba, "N")
        SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
        SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
                                            
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



Private Function CargarTemporalListNominaCoopic(cadWHERE As String, Fdesde As Date, Fhasta As Date, FBaja As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
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

Dim ImpBrutoAnticipado As Currency
Dim HorasPlusCapataz As Currency
Dim PlusCapataz As Currency
Dim Kilometros As Currency

On Error GoTo eCargarTemporalListNominaCoopic
    
    CargarTemporalListNominaCoopic = False
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = "select distinct horas.codtraba, fechahora, straba.pluscapataz, horas.escapataz, sum(coalesce(horasdia, 0)) horas, sum(if(importe is null,0,importe) + if(compleme is null,0,compleme) - if(penaliza is null,0,penaliza)) importe,"
    '[Monica]07/06/2018: a�adimos los kilometros ( importe en euros )
    SQL = SQL & " sum(coalesce(importekms,0)) kilometros "
    SQL = SQL & " from horas, straba where " & cadWHERE
    SQL = SQL & " and horas.codtraba = straba.codtraba " 'and straba.hayembargo = 0"
    
    '[Monica]07/02/2017: si pone fecha de baja solo cogemos los trabajadores con esa fecha de baja en caso contrario
    If FBaja <> "" Then
        SQL = SQL & " and straba.fechabaja = " & DBSet(FBaja, "F")
    Else
        SQL = SQL & " and (straba.fechabaja is null or straba.fechabaja = '')"
    End If
     
    SQL = SQL & " group by 1, 2, 3, 4 "
    SQL = SQL & " having sum(if(importe is null,0,importe) + if(compleme is null,0,compleme) - if(penaliza is null,0,penaliza)) <> 0 "
    SQL = SQL & " order by 1, 2, 3, 4 "
        
    Set Rs = New ADODB.Recordset
        
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        ActTraba = DBLet(Rs!CodTraba, "N")
        AntTraba = DBLet(Rs!CodTraba, "N")
    End If
    v_cadena = String(Day(Fhasta), "N")
    
    Anticipado = 0
    Dias = 0
    HayReg = 0
    ImpBrutoAnticipado = 0
    Kilometros = 0
    
    '[Monica]23/05/2018: inicializamos variables
    HorasPlusCapataz = 0
    PlusCapataz = 0
    Bruto = 0
    
    B = True
    
    While Not Rs.EOF And B
        HayReg = 1
        Mens = "Calculando Dias" & vbCrLf & vbCrLf & "Trabajador: " & ActTraba & vbCrLf
        ActTraba = DBLet(Rs!CodTraba, "N")
        If ActTraba <> AntTraba Then
                                                
            ImpBrutoAnticipado = BrutoAnticipado(AntTraba)
            
            '[Monica]23/05/2018: para el caso de catadau hay que indicar el plus de capataz
'            Bruto = Round2(HorasPlusCapataz * PlusCapataz, 2)
            
                                                
            Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, nombre1, importe1, importe2, importe3, importe4, importe5) values ("
            Sql3 = Sql3 & vUsu.Codigo & ","
            Sql3 = Sql3 & DBSet(AntTraba, "N") & ","
            Sql3 = Sql3 & DBSet(Fhasta, "F") & ","
            Sql3 = Sql3 & DBSet(v_cadena, "T") & ","
            Sql3 = Sql3 & DBSet(Anticipado, "N") & ","
            Sql3 = Sql3 & DBSet(Dias, "N") & ","
            Sql3 = Sql3 & DBSet(Bruto, "N") & ","
            Sql3 = Sql3 & DBSet(ImpBrutoAnticipado, "N") & ","
            '[Monica]07/06/2018: columna para kilometros
            Sql3 = Sql3 & DBSet(Kilometros, "N") & ")"
            
            conn.Execute Sql3
            
            v_cadena = String(Day(Fhasta), "N")
            
            AntTraba = ActTraba
            Anticipado = 0
            Dias = 0
            
            '[Monica]23/05/2018: plus capataz para Catadau
            HorasPlusCapataz = 0
            PlusCapataz = 0
            Bruto = 0
            Kilometros = 0
        End If
        
        i = Day(DBLet(Rs.Fields(1).Value, "F"))
        If i = 1 Then
            v_cadena = "S" & Mid(v_cadena, 2, Len(v_cadena)) ' Replace(v_cadena, "N", "S", I, 1)
        Else
'            v_cadena = Mid(v_cadena, 1, I - 1) & Replace(v_cadena, "N", "S", I, 1)
            v_cadena = Mid(v_cadena, 1, i - 1) & "S" & Mid(v_cadena, i + 1)
        End If
        Dias = Dias + 1
        
        Anticipado = Anticipado + DBLet(Rs!Importe, "N")
        
        '[Monica]23/05/2018: catadau tiene plus de capataz en la columna I
        HorasPlusCapataz = DBLet(Rs!Horas, "N")
        
        PlusCapataz = DBLet(Rs!PlusCapataz, "N")
        If DBLet(Rs!escapataz, "N") = 1 Then
            Bruto = Bruto + Round2(HorasPlusCapataz * PlusCapataz, 2)
        End If
        
        '[Monica]07/06/2018: kilometros
        Kilometros = Kilometros + DBLet(Rs!Kilometros, "N")
        
        
        Rs.MoveNext
    Wend
    If HayReg = 1 Then
        
        ImpBrutoAnticipado = BrutoAnticipado(AntTraba)
        
        '[Monica]23/05/2018: para el caso de catadau hay que indicar el plus de capataz
        'Bruto =  Round2(HorasPlusCapataz * PlusCapataz, 2)
        
        Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, nombre1, importe1, importe2, importe3, importe4, importe5) values ("
        Sql3 = Sql3 & vUsu.Codigo & ","
        Sql3 = Sql3 & DBSet(ActTraba, "N") & ","
        Sql3 = Sql3 & DBSet(Fhasta, "F") & ","
        Sql3 = Sql3 & DBSet(v_cadena, "T") & ","
        Sql3 = Sql3 & DBSet(Anticipado, "N") & ","
        Sql3 = Sql3 & DBSet(Dias, "N") & ","
        Sql3 = Sql3 & DBSet(Bruto, "N") & ","
        Sql3 = Sql3 & DBSet(ImpBrutoAnticipado, "N") & ","
        '[Monica]07/06/2018: importe de kilometros
        Sql3 = Sql3 & DBSet(Kilometros, "N") & ")"
        
        conn.Execute Sql3
    End If
    Set Rs = Nothing
    
    CargarTemporalListNominaCoopic = True
    Exit Function
    
eCargarTemporalListNominaCoopic:
    If Err.Number <> 0 Then
        Mens = Err.Description
        MsgBox "Error " & Mens, vbExclamation
    End If
End Function


Private Function BrutoAnticipado(vTrabajador As String) As Currency
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim BrutoAnt As Currency

    On Error GoTo eBrutoAnticipado
    
    BrutoAnticipado = 0
    
    SQL = "select sum(coalesce(neto34,0)) from rrecibosnomina where codtraba = " & DBSet(vTrabajador, "N")
    SQL = SQL & " and month(fechahora) = " & Combo1(1).ItemData(Combo1(1).ListIndex) & " and year(fechahora) = " & DBSet(txtCodigo(61).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    BrutoAnt = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then BrutoAnt = Rs.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing
    
    '[Monica]25/05/2018: anticipos descontados en ese momento
    SQL = "select idcontador from rrecibosnomina where codtraba = " & DBSet(vTrabajador, "N")
    SQL = SQL & " and month(fechahora) = " & Combo1(1).ItemData(Combo1(1).ListIndex) & " and year(fechahora) = " & DBSet(txtCodigo(61).Text, "N")
        
    SQL = "select sum(importe) from horasanticipos where idcontador in (" & SQL & ") "
        
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then BrutoAnt = BrutoAnt + Rs.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    '[Monica]04/10/2018: por si me debe algo lo pongo aqui para que me lo descuente como si fuera anticipado
    If vParamAplic.Cooperativa = 0 Then
        Dim vFecha As Date
        
        vFecha = CDate("01/" & Format(Combo1(1).ItemData(Combo1(1).ListIndex), "00") & "/" & txtCodigo(61).Text)
        vFecha = DateAdd("m", 1, vFecha)
        vFecha = DateAdd("d", -1, vFecha)
        
        BrutoAnt = BrutoAnt + AnticiposPendientes(vTrabajador)
        
'        '[Monica]04/10/2018: lo que quiero a�adir como anticipado
'        SQL = "update horasanticipos set descontado = 1, fechahora = " & DBSet(vFecha, "F")
'        SQL = SQL & " where codtraba = " & DBSet(vTrabajador, "N") & " and descontado = 0 "
'        conn.Execute SQL
    End If
    
    
    BrutoAnticipado = BrutoAnt
    
    Exit Function
    
eBrutoAnticipado:
    MuestraError Err.Number, "Error en el c�lculo de Bruto Anticipado", Err.Description
End Function


Private Function AnticiposPendientes(CodTraba As String) As Currency
Dim SQL As String

    SQL = "select sum(importe) from horasanticipos where codtraba = " & DBSet(CodTraba, "N")
    SQL = SQL & " and descontado = 0 "
    
    AnticiposPendientes = DevuelveValor(SQL)
    
End Function



Private Function CalculoCapatazServicios() As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim SQL1 As String
Dim Importe As Currency

    On Error GoTo eCalculoCapatazServicios

    CalculoCapatazServicios = False
        
    conn.BeginTrans
        
    Importe = 0

    SQL = "select straba.* from (rcuadrilla inner join rcuadrilla_trabajador on rcuadrilla.codcuadrilla = rcuadrilla_trabajador.codcuadrilla) "
    SQL = SQL & " inner join straba on rcuadrilla_trabajador.codtraba = straba.codtraba "
    SQL = SQL & " where rcuadrilla.codcapat = " & DBSet(txtCodigo(80).Text, "N")
    SQL = SQL & " and (straba.fechabaja is null or straba.fechabaja = '')"
    
    If TotalRegistrosConsulta(SQL) <> 0 Then
    
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            SQL = "select count(*) from horas where fechahora = " & DBSet(txtCodigo(82).Text, "F")
            SQL = SQL & " and codvarie = 0 "
            SQL = SQL & " and codtraba = " & DBSet(Rs!CodTraba, "N")
            SQL = SQL & " and codcapat = " & DBSet(txtCodigo(80).Text, "N")
            
            If TotalRegistros(SQL) = 0 Then
                SQL1 = "insert into horas (fechahora,codvarie,codtraba,codcapat,importe,compleme, fecharec,intconta,pasaridoc,codalmac) values ("
                SQL1 = SQL1 & DBSet(txtCodigo(82).Text, "F") & ","
                SQL1 = SQL1 & "0,"
                SQL1 = SQL1 & DBSet(Rs!CodTraba, "N") & ","
                SQL1 = SQL1 & DBSet(txtCodigo(80).Text, "N") & ",null, null,"
                SQL1 = SQL1 & "null,0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ") "
                
                conn.Execute SQL1
            End If
            
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    
    Else
    
        MsgBox "No hay trabajadores de esa cuadrilla. Revise.", vbExclamation
        conn.CommitTrans
        Exit Function
    End If
    
    
    conn.CommitTrans
                
    CalculoCapatazServicios = True
    Exit Function
    
eCalculoCapatazServicios:
    MuestraError Err.Number, "Calculo Trabajadores para un Capataz Servicios", Err.Description
    conn.RollbackTrans
    TerminaBloquear
End Function

