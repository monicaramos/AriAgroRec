VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBodListAnticipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6660
   Icon            =   "frmBodListAnticipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6030
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameDesFacturacion 
      Height          =   5055
      Left            =   -45
      TabIndex        =   29
      Top             =   45
      Width           =   6555
      Begin VB.Frame FrameTipoFactura 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   390
         TabIndex        =   45
         Top             =   1680
         Width           =   3615
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BeginProperty Font 
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
            ItemData        =   "frmBodListAnticipos.frx":000C
            Left            =   1695
            List            =   "frmBodListAnticipos.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Tag             =   "Recolección|N|N|0|3|rhisfruta|recolect|||"
            Top             =   105
            Width           =   1425
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
            Index           =   4
            Left            =   45
            TabIndex        =   46
            Top             =   105
            Width           =   1905
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
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   2475
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   40
         Tag             =   "admon"
         Top             =   1170
         Width           =   1545
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
         Left            =   2085
         MaxLength       =   7
         TabIndex        =   32
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2910
         Width           =   1260
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
         Index           =   9
         Left            =   2085
         MaxLength       =   7
         TabIndex        =   31
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2505
         Width           =   1260
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
         Index           =   11
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   33
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3585
         Width           =   1350
      End
      Begin VB.CommandButton cmdCancelDesF 
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
         Left            =   5175
         TabIndex        =   35
         Top             =   4440
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepDesF 
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
         Left            =   3960
         TabIndex        =   34
         Top             =   4440
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   255
         Left            =   420
         TabIndex        =   44
         Top             =   4095
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Este proceso borra facturas correlativas "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   210
         TabIndex        =   43
         Top             =   450
         Width           =   5820
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Actualiza contadores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   42
         Top             =   780
         Width           =   5595
      End
      Begin VB.Label Label6 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   1260
         TabIndex        =   41
         Top             =   1170
         Width           =   2235
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
         Index           =   17
         Left            =   1215
         TabIndex        =   39
         Top             =   2910
         Width           =   735
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
         Index           =   10
         Left            =   1215
         TabIndex        =   38
         Top             =   2550
         Width           =   780
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
         Index           =   9
         Left            =   450
         TabIndex        =   37
         Top             =   2280
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
         Index           =   8
         Left            =   465
         TabIndex        =   36
         Top             =   3270
         Width           =   1815
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1785
         Picture         =   "frmBodListAnticipos.frx":0010
         ToolTipText     =   "Buscar fecha"
         Top             =   3585
         Width           =   240
      End
   End
   Begin VB.Frame FrameAnticipos 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check1 
         Caption         =   "Complementaria"
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
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   59
         Top             =   4680
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton CmdAcepLiqAlmzCas 
         Caption         =   "LiqAlmzCas"
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
         Height          =   375
         Left            =   3960
         TabIndex        =   58
         Top             =   5490
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Frame FrameAgrupado 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   3120
         TabIndex        =   55
         Top             =   3360
         Width           =   3375
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
            Index           =   3
            Left            =   1515
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Tag             =   "Recolectado|N|N|0|1|rhisfruta|recolect||N|"
            Top             =   60
            Width           =   1740
         End
         Begin VB.Label Label11 
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
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   57
            Top             =   90
            Width           =   1395
         End
      End
      Begin VB.Frame FrameRecolectado 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   3240
         TabIndex        =   50
         Top             =   2910
         Width           =   3135
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
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Tag             =   "Recolectado|N|N|0|1|rhisfruta|recolect||N|"
            Top             =   120
            Width           =   1740
         End
         Begin VB.Label Label11 
            Caption         =   "Recolectado"
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
            Index           =   3
            Left            =   15
            TabIndex        =   52
            Top             =   150
            Width           =   1260
         End
      End
      Begin VB.Frame FrameOpciones 
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   3120
         TabIndex        =   47
         Top             =   3870
         Width           =   2520
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
            Left            =   135
            TabIndex        =   49
            Top             =   150
            Width           =   2280
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Factura"
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
            Left            =   135
            TabIndex        =   48
            Top             =   450
            Width           =   1995
         End
      End
      Begin VB.Frame FrameFechaAnt 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   600
         Left            =   390
         TabIndex        =   26
         Top             =   3900
         Width           =   3045
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
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   7
            Top             =   240
            Width           =   1350
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   930
            Picture         =   "frmBodListAnticipos.frx":009B
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Anticipo"
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
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   1470
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
         Index           =   21
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   2490
         Width           =   3735
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
         Index           =   20
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   2085
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
         Index           =   21
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2490
         Width           =   915
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
         Index           =   20
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2085
         Width           =   915
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListAnticipos.frx":0126
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmBodListAnticipos.frx":0430
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
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
         Index           =   13
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1500
         Width           =   3735
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
         Index           =   12
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   1095
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
         Index           =   13
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1500
         Width           =   930
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
         Index           =   12
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1095
         Width           =   930
      End
      Begin VB.CommandButton cmdAceptarAnt 
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
         Left            =   4110
         TabIndex        =   8
         Top             =   5310
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancelAnt 
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
         Left            =   5235
         TabIndex        =   9
         Top             =   5295
         Width           =   1065
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
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3420
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
         Index           =   6
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3015
         Width           =   1350
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   4980
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   12
         Left            =   390
         TabIndex        =   54
         Top             =   5520
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   10
         Left            =   390
         TabIndex        =   53
         Top             =   5310
         Width           =   3525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1350
         MouseIcon       =   "frmBodListAnticipos.frx":073A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1350
         MouseIcon       =   "frmBodListAnticipos.frx":088C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2085
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
         Index           =   28
         Left            =   720
         TabIndex        =   25
         Top             =   2475
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
         Index           =   18
         Left            =   720
         TabIndex        =   24
         Top             =   2085
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
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
         Left            =   390
         TabIndex        =   23
         Top             =   1830
         Width           =   525
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1350
         Picture         =   "frmBodListAnticipos.frx":09DE
         ToolTipText     =   "Buscar fecha"
         Top             =   3015
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1350
         Picture         =   "frmBodListAnticipos.frx":0A69
         ToolTipText     =   "Buscar fecha"
         Top             =   3420
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1350
         MouseIcon       =   "frmBodListAnticipos.frx":0AF4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1530
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1350
         MouseIcon       =   "frmBodListAnticipos.frx":0C46
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1095
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
         Index           =   27
         Left            =   390
         TabIndex        =   20
         Top             =   855
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Listado de Anticipos"
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
         TabIndex        =   19
         Top             =   345
         Width           =   5925
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
         Left            =   675
         TabIndex        =   18
         Top             =   1500
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
         Index           =   22
         Left            =   675
         TabIndex        =   17
         Top             =   1095
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
         Index           =   21
         Left            =   690
         TabIndex        =   16
         Top             =   3420
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
         Index           =   20
         Left            =   690
         TabIndex        =   15
         Top             =   3075
         Width           =   690
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
         Index           =   19
         Left            =   405
         TabIndex        =   14
         Top             =   2835
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmBodListAnticipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-+

Option Explicit

Public OpcionListado As Byte
' BODEGA
    '==== Listados / Procesos ANTICIPOS BODEGA ====
    '=============================
    ' 2 .- Prevision de Pagos de Anticipos
    ' 3 .- Facturación de Anticipos
    ' 5 .- Deshacer proceso de Facturación Anticipos
    
    
    '==== Listados / Procesos LIQUIDACIONES BODEGA====
    '================================
    ' 13 .- Prevision de Pagos de Liquidacion
    ' 14 .- Facturación de Liquidacion
    ' 15 .- Deshacer proceso de Facturación Anticipos
    
' ALMAZARA
    '==== Listados / Procesos ANTICIPOS ALMAZARA ====
    '=============================
    ' 20 .- Prevision de Pagos de Anticipos
    ' 30 .- Facturación de Anticipos
    ' 50 .- Deshacer proceso de Facturación Anticipos
    
    
    '==== Listados / Procesos LIQUIDACIONES ALMAZARA====
    '================================
    ' 130 .- Prevision de Pagos de Liquidacion
    ' 140 .- Facturación de Liquidacion
    ' 150 .- Deshacer proceso de Facturación Anticipos


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


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim Indice As Integer

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub Check1_Click(Index As Integer)
    If Index = 7 Then
'        CertificadoRetencionesVisible
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 7 Then
'        CertificadoRetencionesVisible
    End If
End Sub

Private Sub CmdAcepDesF_Click()
Dim Tipo As Byte
    If DatosOK Then
        Pb2.visible = True
        Select Case OpcionListado
            Case 5 ' anticipo de bodega
                Tipo = 6
            Case 15 ' liquidacion de bodega
                Tipo = 7
            Case 50 ' anticipo de almazara
                Tipo = 4
            Case 150 ' liquidacion de almazara
                Tipo = 5
        End Select
        If DeshacerFacturacion(Tipo, txtCodigo(9).Text, txtCodigo(10).Text, txtCodigo(11).Text, Pb2) Then
            MsgBox "Proceso realizado correctamente", vbExclamation
            cmdCancelDesF_Click
        End If
    End If
End Sub


Private Sub CmdAcepLiqAlmzCas_Click()
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
Dim B As Boolean
Dim Sql2 As String

Dim Seccion As Integer
Dim vTipo As Byte

    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOK Then
    
        ' indicamos en la seccion si se trata de bodega o de almazara
        Select Case OpcionListado
            Case 20, 30, 50, 130, 140, 150
                If vParamAplic.SeccionAlmaz = "" Then
                    MsgBox "No tiene asignada la seccion de almazara en parámetros. Revise", vbExclamation
                    Exit Sub
                Else
                    Seccion = CInt(vParamAplic.SeccionAlmaz)
                End If
        End Select
            
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtCodigo(12).Text)
        cHasta = Trim(txtCodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtCodigo(20).Text)
        cHasta = Trim(txtCodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtCodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtCodigo(20).Text, "N")
        If txtCodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtCodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtCodigo(6).Text)
        cHasta = Trim(txtCodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & tabla & ".fecalbar}"
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
        
        nTabla = "((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
        Select Case OpcionListado
            Case 20, 30, 50, 130, 140, 150
                nTabla = nTabla & " and grupopro.codgrupo = 5 " ' grupo SOLO puede ser 5=almazara
        End Select
        
        Select Case OpcionListado
            
            Case 13, 130 ' Prevision de pago de liquidacion
'[Monica] 09/09/2009: parametrizamos la prevision de pago de liquidacion
'                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
'                    cadNombreRPT = "rPrevPagosLiq.rpt"
'                Else ' agrupado por variedad
'                    cadNombreRPT = "rPrevPagosLiq1.rpt"
'                End If
                
                'Nombre fichero .rpt a Imprimir
                indRPT = 43 ' informe de prevision de pago de liquidacion de bodega
                
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"MoiPrevPagosLiqBod.rpt"
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    ' no hacemos nada dejamos el nombre de fichero como estaba
                    
                Else ' agrupado por variedad
                    cadNombreRPT = Replace(cadNombreRPT, "PrevPagosLiq.rpt", "PrevPagosLiq1.rpt")
                End If
                
                If OpcionListado = 13 Then
                    cadTitulo = "Previsión de Pago de Liquidación Bodega"
                Else
                    cadTitulo = "Previsión de Pago de Liquidación Almazara"
                End If
                
            Case 140 ' Facturación de Liquidacion
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación Almazara"
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
        
            Select Case OpcionListado
                Case 130, 140
                    TipoPrec = 8 ' LIQUIDACIONES ALMAZARA
                    vTipo = 1
            End Select
            
            'comprobamos que los tipos de iva existen en la contabilidad de horto
            If Not ComprobarTiposIVA(nTabla, cadSelect) Then Exit Sub
            
            If Not HayAlbaranesSinPrecio(vTipo, nTabla, cadSelect, vParamAplic.Cooperativa) Then
                
                Select Case OpcionListado
                    Case 130 '13- listado de prevision de pagos de liquidaciones de almazara
                        B = CargarTemporalLiquidacionAlmazaraCastelduc(nTabla, cadSelect)
                        
                        If B Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                            ConSubInforme = False
                            
                            LlamarImprimir
                        End If
                    
                    Case 30, 140 '30 .- factura de anticipos de almazara
                                 '140.- factura de liquidaciones de almazara
                        Nregs = TotalFacturas(nTabla, cadSelect)
                        If Nregs <> 0 Then
                            If Not ComprobarTiposMovimiento(TipoPrec, nTabla, cadSelect) Then
                                Exit Sub
                            End If
                            
                            Me.Pb1.visible = True
                            Me.Pb1.Max = Nregs
                            Me.Pb1.Value = 0
                            Me.Refresh
                            B = FacturacionLiquidacionesAlmazaraCastelduc(nTabla, cadSelect, txtCodigo(15).Text, Me.Pb1)
                                
                            If B Then
                                MsgBox "Proceso realizado correctamente.", vbExclamation
                                               
                                'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                                If Me.Check1(2).Value Then
                                    cadFormula = ""
                                    CadParam = CadParam & "pFecFac= """ & txtCodigo(15).Text & """|"
                                    numParam = numParam + 1
                                    CadParam = CadParam & "pTitulo= ""Resumen Facturación Liquidaciones Almazara""|"
                                    numParam = numParam + 1
                                    
                                    FecFac = CDate(txtCodigo(15).Text)
                                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                    ConSubInforme = False
                                    
                                    LlamarImprimir
                                End If
                                'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS/LIQUIDACION ALMAZARA
                                If Me.Check1(3).Value Then
                                    cadFormula = ""
                                    cadSelect = ""
                                    If TipoPrec = 7 Then 'Tipo de Factura: Anticipo
                                        cadAux = "({stipom.tipodocu} = 7)"
                                    Else  'Tipo de Factura: Liquidación
                                        cadAux = "({stipom.tipodocu} = 8)"
                                    End If
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                    'Nº Factura
                                    If TipoPrec = 7 Then
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(4) & "])"
                                    Else
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(5) & "])"
                                    End If
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                                    'Fecha de Factura
                                    FecFac = CDate(txtCodigo(15).Text)
                                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                                    indRPT = 42 'Impresion de facturas de socios
                                    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                                    'Nombre fichero .rpt a Imprimir
                                    cadNombreRPT = nomDocu
                                    'Nombre fichero .rpt a Imprimir
                                    cadTitulo = "Reimpresión de Facturas Liquidaciones Almazara"
                                    ConSubInforme = True

                                    LlamarImprimir

                                    If frmVisReport.EstaImpreso Then
                                        ActualizarRegistrosFac "rfactsoc", cadSelect
                                    End If
                                End If
                                'SALIR DE LA FACTURACION DE ANTICIPOS / LIQUIDACIONES
                                cmdCancelAnt_Click
                            End If
                        Else
                            MsgBox "No hay entradas a facturar.", vbExclamation
                        End If
                
                End Select
            End If
        End If
    End If

End Sub

Private Sub cmdAceptarAnt_Click()

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
Dim B As Boolean
Dim Sql2 As String

Dim Seccion As Integer
Dim vTipo As Byte

    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOK Then
    
        ' indicamos en la seccion si se trata de bodega o de almazara
        Select Case OpcionListado
            Case 2, 3, 5, 13, 14, 15
                If vParamAplic.SeccionBodega = "" Then
                    MsgBox "No tiene asignada la seccion de bodega en parámetros. Revise", vbExclamation
                    Exit Sub
                Else
                    Seccion = CInt(vParamAplic.SeccionBodega)
                End If
            Case 20, 30, 50, 130, 140, 150
                If vParamAplic.SeccionAlmaz = "" Then
                    MsgBox "No tiene asignada la seccion de almazara en parámetros. Revise", vbExclamation
                    Exit Sub
                Else
                    Seccion = CInt(vParamAplic.SeccionAlmaz)
                End If
        End Select
            
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtCodigo(12).Text)
        cHasta = Trim(txtCodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtCodigo(20).Text)
        cHasta = Trim(txtCodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtCodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtCodigo(20).Text, "N")
        If txtCodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtCodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtCodigo(6).Text)
        cHasta = Trim(txtCodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & tabla & ".fecalbar}"
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
        
        
        nTabla = "((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
        
        Select Case OpcionListado
            Case 2, 3, 5, 13, 14, 15
                nTabla = nTabla & " and grupopro.codgrupo = 6 " ' grupo SOLO puede ser 6=bodega
            Case 20, 30, 50, 130, 140, 150
                nTabla = nTabla & " and grupopro.codgrupo = 5 " ' grupo SOLO puede ser 5=almazara
        End Select
        
        
        Select Case OpcionListado
            Case 2 ' Prevision de pago de anticipos 2=bodega
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    cadNombreRPT = "rPrevPagosAntBod.rpt"
                Else ' agrupado por variedad
                    cadNombreRPT = "rPrevPagosAntBod1.rpt"
                End If
                cadTitulo = "Previsión de Pago de Anticipos Bodega"
                    
            Case 20  ' Prevision de pago de anticipos 20=almazara
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                        cadNombreRPT = "rPrevPagosAnt.rpt"
                    Else ' agrupado por variedad
                        cadNombreRPT = "rPrevPagosAnt1.rpt"
                    End If
                cadTitulo = "Previsión de Pago de Anticipos Almazara"
            
            Case 3 ' Facturación de Anticipos bodega
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Anticipos Bodega"
                
            Case 30 'Facturacion de Anticipos de almazara
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Anticipos Almazara"
            
            Case 13, 130 ' Prevision de pago de liquidacion
'[Monica] 09/09/2009: parametrizamos la prevision de pago de liquidacion
'                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
'                    cadNombreRPT = "rPrevPagosLiq.rpt"
'                Else ' agrupado por variedad
'                    cadNombreRPT = "rPrevPagosLiq1.rpt"
'                End If
                
                'Nombre fichero .rpt a Imprimir
                indRPT = 43 ' informe de prevision de pago de liquidacion de bodega
                
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"MoiPrevPagosLiqBod.rpt"
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    ' no hacemos nada dejamos el nombre de fichero como estaba
                    
                Else ' agrupado por variedad
                    cadNombreRPT = Replace(cadNombreRPT, "PrevPagosLiq.rpt", "PrevPagosLiq1.rpt")
                End If
                
                If OpcionListado = 13 Then
                    cadTitulo = "Previsión de Pago de Liquidación Bodega"
                Else
                    cadTitulo = "Previsión de Pago de Liquidación Almazara"
                End If
                
            Case 14 ' Facturación de Liquidacion
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación Bodega"
            
            Case 140 ' Facturación de Liquidacion
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación Almazara"
            
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
        
            Select Case OpcionListado
                Case 2, 3
                    TipoPrec = 9 ' ANTICIPOS BODEGA
                    vTipo = 0
                Case 13, 14
                    TipoPrec = 10 ' LIQUIDACIONES BODEGA
                    vTipo = 1
                Case 20, 30
                    TipoPrec = 7 ' ANTICIPOS ALMAZARA
                    vTipo = 0
                Case 130, 140
                    TipoPrec = 8 ' LIQUIDACIONES ALMAZARA
                    vTipo = 1
            End Select
            
            'comprobamos que los tipos de iva existen en la contabilidad de horto
            If Not ComprobarTiposIVA(nTabla, cadSelect) Then Exit Sub
            
            If HayPreciosVariedadesBodegaAlmazara(vTipo, nTabla, cadSelect, vParamAplic.Cooperativa) Then
                'D/H fecha
                cDesde = Trim(txtCodigo(6).Text)
                cHasta = Trim(txtCodigo(7).Text)
                cadDesde = CDate(cDesde)
                cadhasta = CDate(cHasta)
                cadAux = "{rprecios.fechaini}= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechaini}=" & DBSet(txtCodigo(6).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                cadAux = "{rprecios.fechafin}= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechafin}=" & DBSet(txtCodigo(7).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                ' si se trata de anticipos--> seleccionamos los precios de anticipos
                ' sino los de liquidaciones
                If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = " & vTipo) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = " & vTipo) Then Exit Sub
                
                Select Case OpcionListado
                    Case 2  '2 - listado de prevision de pagos de anticipos bodega
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rhisfruta.codvarie = rprecios.codvarie "
                        
                        If CargarTemporalAnticiposBodega(nTabla, cadSelect) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                            ConSubInforme = False
                            
                            LlamarImprimir
                        End If
                        
                    Case 20 '20- listado de prevision de pagos de anticipos almazara
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rhisfruta.codvarie = rprecios.codvarie "
                        
                        '[Monica]15/04/2013: Castelduc anticipa en almazara a partir de la tabla de precios pero sobre el total de kilos (no litros)
                        If vParamAplic.Cooperativa = 5 Then
                            If CargarTemporalAnticiposAlmazaraCastelduc(nTabla, cadSelect) Then
                                cadFormula = ""
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                ConSubInforme = False
                                
                                LlamarImprimir
                        
                            End If
                        Else
                            If CargarTemporalAnticiposAlmazara(nTabla, cadSelect) Then
                                cadFormula = ""
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                ConSubInforme = False
                                
                                LlamarImprimir
                            End If
                        End If
                        
                    Case 13 '13- listado de prevision de pagos de liquidaciones
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rhisfruta.codvarie = rprecios.codvarie "
                        If CargarTemporalLiquidacionBodega(nTabla, cadSelect) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                            ConSubInforme = False
                            
                            LlamarImprimir
                        End If
                    
                    Case 130 '13- listado de prevision de pagos de liquidaciones de almazara
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rhisfruta.codvarie = rprecios.codvarie "
                        
                        If vParamAplic.Cooperativa = 1 Then
                            nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rprecios.codvarie = rprecios_calidad.codvarie "
                            nTabla = nTabla & " and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
                            
                            B = CargarTemporalLiquidacionAlmazaraValsur(nTabla, cadSelect, txtCodigo(6).Text, txtCodigo(7).Text)
                        Else
                            B = CargarTemporalLiquidacionAlmazara(nTabla, cadSelect)
                        End If
                        
                        If B Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                            ConSubInforme = False
                            
                            LlamarImprimir
                        End If
                    
                    
                    Case 3, 14 '3 .- factura de anticipos
                               '14.- factura de liquidaciones
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rhisfruta.codvarie = rprecios.codvarie "
                        
                        Nregs = TotalFacturas(nTabla, cadSelect)
                        If Nregs <> 0 Then
                            If Not ComprobarTiposMovimiento(TipoPrec, nTabla, cadSelect) Then
                                Exit Sub
                            End If
                            
                            Me.Pb1.visible = True
                            Me.Pb1.Max = Nregs
                            Me.Pb1.Value = 0
                            Me.Refresh
                            B = False
                            If TipoPrec = 9 Then
                                B = FacturacionAnticiposBodega(nTabla, cadSelect, txtCodigo(15).Text, Me.Pb1)
                            Else
                                B = FacturacionLiquidacionesBodega(nTabla, cadSelect, txtCodigo(15).Text, Me.Pb1, Check1(0).Value = 1)
                            End If
                                
                            If B Then
                                MsgBox "Proceso realizado correctamente.", vbExclamation
                                               
                                'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                                If Me.Check1(2).Value Then
                                    cadFormula = ""
                                    CadParam = CadParam & "pFecFac= """ & txtCodigo(15).Text & """|"
                                    numParam = numParam + 1
                                    If TipoPrec = 9 Then
                                        CadParam = CadParam & "pTitulo= ""Resumen Facturación Anticipos Bodega""|"
                                    Else
                                        CadParam = CadParam & "pTitulo= ""Resumen Facturación Liquidaciones Bodega""|"
                                    End If
                                    numParam = numParam + 1
                                    
                                    FecFac = CDate(txtCodigo(15).Text)
                                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                    ConSubInforme = False
                                    
                                    LlamarImprimir
                                End If
                                'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS/LIQUIDACION BODEGA
                                If Me.Check1(3).Value Then
                                    cadFormula = ""
                                    cadSelect = ""
                                    If TipoPrec = 9 Then 'Tipo de Factura: Anticipo
                                        cadAux = "({stipom.tipodocu} = 9)"
                                    Else  'Tipo de Factura: Liquidación
                                        cadAux = "({stipom.tipodocu} = 10)"
                                    End If
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                    'Nº Factura
                                    If TipoPrec = 9 Then
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(6) & "])"
                                    Else
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(7) & "])"
                                    End If
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                                    'Fecha de Factura
                                    FecFac = CDate(txtCodigo(15).Text)
                                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                                    indRPT = 42 'Impresion de facturas de socios
                                    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                                    'Nombre fichero .rpt a Imprimir
                                    cadNombreRPT = nomDocu
                                    'Nombre fichero .rpt a Imprimir
                                    If TipoPrec = 9 Then
                                        cadTitulo = "Reimpresión de Facturas Anticipos Bodega"
                                    Else
                                        cadTitulo = "Reimpresión de Facturas Liquidaciones Bodega"
                                    End If
                                    ConSubInforme = True

                                    LlamarImprimir

                                    If frmVisReport.EstaImpreso Then
                                        ActualizarRegistrosFac "rfactsoc", cadSelect
                                    End If
                                End If
                                'SALIR DE LA FACTURACION DE ANTICIPOS / LIQUIDACIONES
                                cmdCancelAnt_Click
                            End If
                        Else
                            MsgBox "No hay entradas a facturar.", vbExclamation
                        End If
                
                    Case 30, 140 '30 .- factura de anticipos de almazara
                                 '140.- factura de liquidaciones de almazara
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rhisfruta.codvarie = rprecios.codvarie "
                        
                        Nregs = TotalFacturas(nTabla, cadSelect)
                        If Nregs <> 0 Then
                            If Not ComprobarTiposMovimiento(TipoPrec, nTabla, cadSelect) Then
                                Exit Sub
                            End If
                            
                            Me.Pb1.visible = True
                            Me.Pb1.Max = Nregs
                            Me.Pb1.Value = 0
                            Me.Refresh
                            B = False
                            If TipoPrec = 7 Then
                                '[Monica]15/04/2013: Castelduc hace factura de anticipo de almazara por kilos no por litros
                                If vParamAplic.Cooperativa = 5 Then
                                    B = FacturacionAnticiposAlmazaraCastelduc(nTabla, cadSelect, txtCodigo(15).Text, Me.Pb1)
                                Else
                                    B = FacturacionAnticiposAlmazara(nTabla, cadSelect, txtCodigo(15).Text, Me.Pb1)
                                End If
                            Else
                                If vParamAplic.Cooperativa = 1 Then
                                    nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rprecios.codvarie = rprecios_calidad.codvarie "
                                    nTabla = nTabla & " and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
                                
                                    B = FacturacionLiquidacionesAlmazaraValsur(nTabla, cadSelect, txtCodigo(15).Text, Me.Pb1, txtCodigo(6).Text, txtCodigo(7).Text)
                                Else
                                    B = FacturacionLiquidacionesAlmazara(nTabla, cadSelect, txtCodigo(15).Text, Me.Pb1)
                                End If
                            End If
                                
                            If B Then
                                MsgBox "Proceso realizado correctamente.", vbExclamation
                                               
                                'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                                If Me.Check1(2).Value Then
                                    cadFormula = ""
                                    CadParam = CadParam & "pFecFac= """ & txtCodigo(15).Text & """|"
                                    numParam = numParam + 1
                                    If TipoPrec = 7 Then
                                        CadParam = CadParam & "pTitulo= ""Resumen Facturación Anticipos Almazara""|"
                                    Else
                                        CadParam = CadParam & "pTitulo= ""Resumen Facturación Liquidaciones Almazara""|"
                                    End If
                                    numParam = numParam + 1
                                    
                                    FecFac = CDate(txtCodigo(15).Text)
                                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                    ConSubInforme = True
                                    
                                    LlamarImprimir
                                End If
                                'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS/LIQUIDACION ALMAZARA
                                If Me.Check1(3).Value Then
                                    cadFormula = ""
                                    cadSelect = ""
                                    If TipoPrec = 7 Then 'Tipo de Factura: Anticipo
                                        cadAux = "({stipom.tipodocu} = 7)"
                                    Else  'Tipo de Factura: Liquidación
                                        cadAux = "({stipom.tipodocu} = 8)"
                                    End If
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                    'Nº Factura
                                    If TipoPrec = 7 Then
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(4) & "])"
                                    Else
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(5) & "])"
                                    End If
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                                    'Fecha de Factura
                                    FecFac = CDate(txtCodigo(15).Text)
                                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                                    indRPT = 42 'Impresion de facturas de socios
                                    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                                    'Nombre fichero .rpt a Imprimir
                                    cadNombreRPT = nomDocu
                                    'Nombre fichero .rpt a Imprimir
                                    If TipoPrec = 9 Then
                                        cadTitulo = "Reimpresión de Facturas Anticipos Almazara"
                                    Else
                                        cadTitulo = "Reimpresión de Facturas Liquidaciones Almazara"
                                    End If
                                    ConSubInforme = True

                                    LlamarImprimir

                                    If frmVisReport.EstaImpreso Then
                                        ActualizarRegistrosFac "rfactsoc", cadSelect
                                    End If
                                End If
                                'SALIR DE LA FACTURACION DE ANTICIPOS / LIQUIDACIONES
                                cmdCancelAnt_Click
                            End If
                        Else
                            MsgBox "No hay entradas a facturar.", vbExclamation
                        End If
                
                End Select
            '++monica:27/07/2009
            Else
                MsgBox "No hay precios para las variedades en este rango. Revise.", vbExclamation
            End If
        End If
    End If
End Sub

Private Sub cmdCancelAnt_Click()
    Unload Me
End Sub


Private Sub cmdCancelDesF_Click()
    Unload Me
End Sub


Private Sub Combo1_LostFocus(Index As Integer)
    If Index = 1 Then
        Select Case Combo1(Index).ListIndex
            Case 0 ' anticipo venta campo
                ' si solo hay un tipo de movimiento de anticipo venta campo
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(3) = 1 Then
                    txtCodigo(9).Text = vParamAplic.PrimFactAntVC
                    txtCodigo(10).Text = vParamAplic.UltFactAntVC
                End If
            Case 1 ' liquidacion venta campo
                ' si solo hay un tipo de movimiento de liquidacion venta campo
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(4) = 1 Then
                    txtCodigo(9).Text = vParamAplic.PrimFactLiqVC
                    txtCodigo(10).Text = vParamAplic.UltFactLiqVC
                End If
        End Select
    End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 2, 3, 20, 30
                         ' 2 / 20-Listado de Previsión de pago
                         ' 3 / 30-Facturas de Anticipos
                PonerFoco txtCodigo(12)
                
            Case 5    ' deshacer proceso de facturacion de anticipos bodega
                PonerFoco txtCodigo(8)
                Me.Pb2.visible = False
                ' si solo hay un tipo de movimiento de anticipo
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(9) = 1 Then
                    txtCodigo(9).Text = vParamAplic.PrimFactAntBOD
                    txtCodigo(10).Text = vParamAplic.UltFactAntBOD
                End If
                
            Case 50    ' deshacer proceso de facturacion de anticipos almazara
                PonerFoco txtCodigo(8)
                Me.Pb2.visible = False
                ' si solo hay un tipo de movimiento de anticipo
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(7) = 1 Then
                    txtCodigo(9).Text = vParamAplic.PrimFactAntAlmz
                    txtCodigo(10).Text = vParamAplic.UltFactAntAlmz
                End If
                
            Case 13, 14, 130, 140
                            ' 13 / 130 -Listado de Previsión de pago
                            ' 14 / 140 -Facturas de Liquidacion
                PonerFoco txtCodigo(12)
            
            Case 15    ' deshacer proceso de facturacion de liquidacion
                PonerFoco txtCodigo(8)
                Me.Pb2.visible = False
                ' si solo hay un tipo de movimiento de liquidacion
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(10) = 1 Then
                    txtCodigo(9).Text = vParamAplic.PrimFactLiqBOD
                    txtCodigo(10).Text = vParamAplic.UltFactLiqBOD
                End If
                
            Case 150   ' deshacer proceso de facturacion de liquidacion de almazara
                PonerFoco txtCodigo(8)
                Me.Pb2.visible = False
                ' si solo hay un tipo de movimiento de liquidacion
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(8) = 1 Then
                    txtCodigo(9).Text = vParamAplic.PrimFactLiqAlmz
                    txtCodigo(10).Text = vParamAplic.UltFactLiqAlmz
                End If
                
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
    
    For H = 12 To 13
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 20 To 21
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    FrameAnticipos.visible = False
    FrameDesFacturacion.visible = False
    '###Descomentar
'    CommitConexion
    
    '[Monica]23/11/2012: No es Complementaria en ningun sitio, ni en bodega ni en almazara
    Me.Check1(0).Value = 0
    
    Select Case OpcionListado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 2, 13   '2 - Listado de prevision de pagos de anticipos
                 '13- Listado de prevision de pagos de liquidacion
        FrameAnticiposVisible True, H, W
        tabla = "rhisfruta"
        Me.FrameFechaAnt.visible = False
        Me.FrameFechaAnt.Enabled = False
        Me.FrameAgrupado.visible = True
        Me.FrameAgrupado.Enabled = True
        If OpcionListado = 2 Then
            Me.Label3.Caption = "Previsión de Pagos Anticipos Bodega"
        Else
            Me.Label3.Caption = "Previsión de Pagos Liquidación Bodega"
        End If
        
        Me.Pb1.visible = False
        Me.Label2(10).Caption = ""
        Me.Label2(12).Caption = ""
        Me.FrameOpciones.visible = False
        Me.FrameOpciones.Enabled = False
        
        CargaCombo
        Combo1(2).ListIndex = 2
        Combo1(3).ListIndex = 0
        
        FrameRecolectado.visible = False
        FrameRecolectado.Enabled = False
        FrameAgrupado.visible = False
        FrameAgrupado.Enabled = False
        
        Me.Check1(0).Value = 0
        If OpcionListado = 13 Then
            Me.Check1(0).Enabled = True
            Me.Check1(0).visible = True
        End If
        
        
    Case 20, 130   '20 - Listado de prevision de pagos de anticipos almazara
                   '130- Listado de prevision de pagos de liquidacion almazara
        FrameAnticiposVisible True, H, W
        tabla = "rhisfruta"
        Me.FrameFechaAnt.visible = False
        Me.FrameFechaAnt.Enabled = False
        Me.FrameAgrupado.visible = True
        Me.FrameAgrupado.Enabled = True
        If OpcionListado = 20 Then
            Me.Label3.Caption = "Previsión de Pagos Anticipos Almazara"
        Else
            Me.Label3.Caption = "Previsión de Pagos Liquidación Almazara"
        End If
        
        Me.Pb1.visible = False
        Me.Label2(10).Caption = ""
        Me.Label2(12).Caption = ""
        Me.FrameOpciones.visible = False
        Me.FrameOpciones.Enabled = False
        
        CargaCombo
        Combo1(2).ListIndex = 2
        Combo1(3).ListIndex = 0
        
        FrameRecolectado.visible = False
        FrameRecolectado.Enabled = False
        FrameAgrupado.visible = False
        FrameAgrupado.Enabled = False
        
        '[Monica]24/02/2011: Nuevo boton de Aceptar para el caso de Castelduc
        '                    Prevision de liquidacion Almazara
        If vParamAplic.Cooperativa = 5 And OpcionListado = 130 Then
            Me.CmdAcepLiqAlmzCas.visible = True
            Me.CmdAcepLiqAlmzCas.Enabled = True
            Me.CmdAcepLiqAlmzCas.Top = 5310
            Me.CmdAcepLiqAlmzCas.Caption = "&Aceptar"
            Me.cmdAceptarAnt.visible = False
            Me.cmdAceptarAnt.Enabled = False
        Else
            Me.CmdAcepLiqAlmzCas.visible = False
            Me.CmdAcepLiqAlmzCas.Enabled = False
        End If
        
    Case 3, 14   '3 - Factura de Anticipos
                 '14- Factura de Liquidacion
        FrameAnticiposVisible True, H, W
        tabla = "rhisfruta"
        Me.FrameFechaAnt.visible = True
        Me.FrameFechaAnt.Enabled = True
        Me.FrameAgrupado.visible = False
        Me.FrameAgrupado.Enabled = False
        Me.Caption = "Facturación"
        If OpcionListado = 3 Then
            Me.Label3.Caption = "Factura de Anticipos Bodega"
            Me.Label2(25).Caption = "Fecha Anticipo"
        Else
            Me.Label3.Caption = "Factura de Liquidación Bodega"
            Me.Label2(25).Caption = "Fecha Liquidación"
        End If
        Me.Pb1.visible = False
        Me.Label2(10).Caption = ""
        Me.Label2(12).Caption = ""
        Me.FrameOpciones.visible = True
        Me.FrameOpciones.Enabled = True
        
        Me.Check1(3).Enabled = (vParamAplic.Cooperativa <> 4)
        Me.Check1(3).visible = (vParamAplic.Cooperativa <> 4)
            
        Me.Check1(2).Value = 1
        ' en el caso de alzira no imprimos las facturas pq tiene que añadirle los gastos a pie de factura
'        Me.Check1(3).Value = 1
        If vParamAplic.Cooperativa = 4 Then
            Check1(3).Value = 0
        Else
            Check1(3).Value = 1
        End If
        
        txtCodigo(15).Text = Format(Now, "dd/mm/yyyy")
        
        CargaCombo
        Combo1(2).ListIndex = 2
    
        Me.Check1(0).Value = 0
        If OpcionListado = 14 Then
            Me.Check1(0).Enabled = True
            Me.Check1(0).visible = True
        End If
    
    Case 30, 140   '30 - Factura de Anticipos almazara
                   '140- Factura de Liquidacion almazara
        FrameAnticiposVisible True, H, W
        tabla = "rhisfruta"
        Me.FrameFechaAnt.visible = True
        Me.FrameFechaAnt.Enabled = True
        Me.FrameAgrupado.visible = False
        Me.FrameAgrupado.Enabled = False
        Me.Caption = "Facturación"
        If OpcionListado = 30 Then
            Me.Label3.Caption = "Factura de Anticipos Almazara"
            Me.Label2(25).Caption = "Fecha Anticipo"
        Else
            Me.Label3.Caption = "Factura de Liquidación Almazara"
            Me.Label2(25).Caption = "Fecha Liquidación"
        End If
        Me.Pb1.visible = False
        Me.Label2(10).Caption = ""
        Me.Label2(12).Caption = ""
        Me.FrameOpciones.visible = True
        Me.FrameOpciones.Enabled = True
        
        Me.Check1(3).Enabled = (vParamAplic.Cooperativa <> 4)
        Me.Check1(3).visible = (vParamAplic.Cooperativa <> 4)
            
        
        Me.Check1(2).Value = 1
        ' en el caso de alzira no imprimos las facturas pq tiene que añadirle los gastos a pie de factura
'        Me.Check1(3).Value = 1
        If vParamAplic.Cooperativa = 4 Then
            Check1(3).Value = 0
        Else
            Check1(3).Value = 1
        End If
        
        txtCodigo(15).Text = Format(Now, "dd/mm/yyyy")
        
        CargaCombo
        Combo1(2).ListIndex = 2
    
        '[Monica]24/02/2011: Nuevo boton de Aceptar para el caso de Castelduc
        '                    Factura liquidacion Almazara
        If vParamAplic.Cooperativa = 5 And OpcionListado = 140 Then
            Me.CmdAcepLiqAlmzCas.visible = True
            Me.CmdAcepLiqAlmzCas.Enabled = True
            Me.CmdAcepLiqAlmzCas.Top = 5310
            Me.CmdAcepLiqAlmzCas.Caption = "&Aceptar"
            Me.cmdAceptarAnt.visible = False
            Me.cmdAceptarAnt.Enabled = False
        Else
            Me.CmdAcepLiqAlmzCas.visible = False
            Me.CmdAcepLiqAlmzCas.Enabled = False
        End If
    
    Case 5, 50   ' Deshacer Proceso de facturación de Anticipos
        ActivarCLAVE
        FrameTipoFactura.visible = False
        FrameDesFacturacionVisible True, H, W
        tabla = "rfactsoc"
        If OpcionListado = 5 Then
            Me.Caption = "Deshacer Proceso Facturación de Anticipos Bodega"
        Else
            Me.Caption = "Deshacer Proceso Facturación de Anticipos Almazara"
        End If
        
    
    Case 15, 150   ' Deshacer Proceso de facturación de Liquidacion
        ActivarCLAVE
        FrameTipoFactura.visible = False
        FrameDesFacturacionVisible True, H, W
        tabla = "rfactsoc"
        If OpcionListado = 15 Then
            Me.Caption = "Deshacer Proceso Facturación de Liquidación Bodega"
        Else
            Me.Caption = "Deshacer Proceso Facturación de Liquidación Almazara"
        End If
        
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case OpcionListado
        Case 3
            DesBloqueoManual ("FACFNB")
        Case 14
            DesBloqueoManual ("FACFLB")
        Case 30
            DesBloqueoManual ("FACFNZ")
        Case 140
            DesBloqueoManual ("FACFLZ")
    End Select
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de calidades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        SQL = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
    Else
        SQL = " {variedades.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, SQL) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 18, 19, 20, 21, 28, 29 'Clases
            AbrirFrmClase (Index)
        
        Case 0, 1, 12, 13, 16, 17, 24, 25 'SOCIOS
            AbrirFrmSocios (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
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
        Case 0
            Indice = 6
        Case 1
            Indice = 7
        Case 2
            Indice = 15
        Case 3, 4
            Indice = Index - 1
        Case 5
            Indice = 14
        Case 6
            Indice = 11
        Case 7, 8
            Indice = Index + 15
        Case 9, 10
            Indice = Index + 17
        Case 11, 12
            Indice = Index + 21
        Case 13
            Indice = 32
    End Select

    imgFec(0).Tag = Indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Indice).Text <> "" Then frmC.NovaData = txtCodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFec(0).Tag)) '<===
    ' ********************************************

End Sub



Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
    If OpcionListado = 10 Then
'        If Index = 40 Then
'            BarraEst.SimpleText = " CL = Calle    AV = Avenida."
'        Else
'            BarraEst.SimpleText = ""
'        End If
'        BarraEst.visible = (BarraEst.SimpleText <> "")
    End If
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
            Case 0: KEYBusqueda KeyAscii, 0 'socio desde
            Case 1: KEYBusqueda KeyAscii, 1 'socio hasta
            Case 12: KEYBusqueda KeyAscii, 12 'socio desde
            Case 13: KEYBusqueda KeyAscii, 13 'socio hasta
            Case 16: KEYBusqueda KeyAscii, 16 'socio desde
            Case 17: KEYBusqueda KeyAscii, 17 'socio hasta
            Case 24: KEYBusqueda KeyAscii, 24 'socio desde
            Case 25: KEYBusqueda KeyAscii, 25 'socio hasta
            Case 34: KEYBusqueda KeyAscii, 34 'socio desde
            Case 35: KEYBusqueda KeyAscii, 35 'socio hasta
            Case 18: KEYBusqueda KeyAscii, 18 'clase desde
            Case 19: KEYBusqueda KeyAscii, 19 'clase hasta
            Case 20: KEYBusqueda KeyAscii, 20 'clase desde
            Case 21: KEYBusqueda KeyAscii, 21 'clase hasta
            Case 28: KEYBusqueda KeyAscii, 28 'clase desde
            Case 29: KEYBusqueda KeyAscii, 29 'clase hasta
            Case 26: KEYFecha KeyAscii, 9 'fecha desde
            Case 27: KEYFecha KeyAscii, 10 'fecha hasta
            Case 6: KEYFecha KeyAscii, 0 'fecha desde
            Case 7: KEYFecha KeyAscii, 1 'fecha hasta
            Case 22: KEYFecha KeyAscii, 7 'fecha desde
            Case 23: KEYFecha KeyAscii, 8 'fecha hasta
            Case 2: KEYFecha KeyAscii, 3 'fecha desde
            Case 3: KEYFecha KeyAscii, 4 'fecha hasta
            Case 32: KEYFecha KeyAscii, 13 'fecha desde
            
            
            Case 11: KEYFecha KeyAscii, 6 'fecha
            Case 14: KEYFecha KeyAscii, 5 'fecha
            Case 15: KEYFecha KeyAscii, 2 'fecha
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
    imgFec_Click (Indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim B As Boolean

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
    
        Case 0, 1, 12, 13, 16, 17, 24, 25, 34, 35 'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
        Case 4, 5 ' NROS DE FACTURA
            PonerFormatoEntero txtCodigo(Index)
            
        Case 2, 3, 6, 7, 11, 15, 26, 27, 32 'FECHAS
            B = True
           '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            If txtCodigo(Index).Text <> "" Then
                If Index = 6 Or Index = 7 Then
                    B = PonerFormatoFecha(txtCodigo(Index), True)
                Else
                    B = PonerFormatoFecha(txtCodigo(Index))
                End If
            End If
            If B And Index = 7 And (Me.OpcionListado = 1 Or Me.OpcionListado = 3 Or Me.OpcionListado = 12 Or Me.OpcionListado = 14) Then PonerFoco txtCodigo(15)
            If B And Index = 15 And (Me.OpcionListado = 1 Or Me.OpcionListado = 3 Or Me.OpcionListado = 12 Or Me.OpcionListado = 14) Then
                 cmdAceptarAnt.SetFocus
            End If
            
        Case 14, 22, 23  ' FECHA DE GENERACION DE FACTURA
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 8 ' password de deshacer facturacion
            If txtCodigo(Index).Text = "" Then Exit Sub
            If Trim(txtCodigo(Index).Text) <> Trim(txtCodigo(Index).Tag) Then
                MsgBox "    ACCESO DENEGADO    ", vbExclamation
                txtCodigo(Index).Text = ""
                PonerFoco txtCodigo(Index)
            Else
                DesactivarCLAVE
                Select Case OpcionListado
                    Case 5, 15 '5 = anticipos
                               '15= liquidaciones
                        PonerFoco txtCodigo(9)
                    Case 7 ' venta campo
                        PonerFocoCmb Combo1(1)
                End Select
            End If
        
        Case 9, 10 ' numero de facturas
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
        
        Case 30, 31, 37, 39 ' datos de modelo190 y modelo346
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
            
        Case 33 ' nro de justificante en el certificado de retenciones
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoEntero txtCodigo(Index)
            End If
            
        Case 18, 19, 20, 21, 28, 29 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
    End Select
End Sub

Private Sub FrameAnticiposVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
Dim B As Boolean

'Frame para el listado de socios por seccion
    Me.FrameAnticipos.visible = visible
    If visible = True Then
        Me.FrameAnticipos.Top = -90
        Me.FrameAnticipos.Left = 0
        Me.FrameAnticipos.Height = 6015
        Me.FrameAnticipos.Width = 6615
        W = Me.FrameAnticipos.Width
        H = Me.FrameAnticipos.Height
        
        B = (OpcionListado = 1 Or OpcionListado = 2 Or OpcionListado = 3 Or _
             OpcionListado = 12 Or OpcionListado = 13 Or OpcionListado = 14 Or _
             OpcionListado = 10 Or OpcionListado = 20 Or OpcionListado = 30 Or _
             OpcionListado = 120 Or OpcionListado = 130 Or OpcionListado = 140)
             
        
        FrameRecolectado.Enabled = B
        FrameRecolectado.visible = B
    
    End If
End Sub




Private Sub FrameDesFacturacionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameDesFacturacion.visible = visible
    If visible = True Then
        Me.FrameDesFacturacion.Top = -90
        Me.FrameDesFacturacion.Left = 0
        Me.FrameDesFacturacion.Height = 5055
        Me.FrameDesFacturacion.Width = 6615
        W = Me.FrameDesFacturacion.Width
        H = Me.FrameDesFacturacion.Height
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
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = OpcionListado
        .ConSubInforme = ConSubInforme
        .Show vbModal
    End With
End Sub


Private Sub AbrirFrmCalidad(Indice As Integer)
    indCodigo = Indice
    Set frmCal = New frmManCalidades
    frmCal.DatosADevolverBusqueda = "2|3|"
    frmCal.Show vbModal
    Set frmCal = Nothing
End Sub

Private Sub AbrirFrmSeccion(Indice As Integer)
    indCodigo = Indice
    Set frmSec = New frmManSeccion
    frmSec.DatosADevolverBusqueda = "0|1|"
    frmSec.Show vbModal
    Set frmSec = Nothing
End Sub

Private Sub AbrirFrmSocios(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmSituacion(Indice As Integer)
    indCodigo = Indice
    Set frmSit = New frmManSituCamp
    frmSit.DatosADevolverBusqueda = "0|1|"
    frmSit.Show vbModal
    Set frmSit = Nothing
End Sub

Private Sub AbrirFrmSocio(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmClase(Indice As Integer)
    indCodigo = Indice
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtCodigo(Indice).Text
    
    Set frmCla = Nothing
End Sub



Private Sub AbrirFrmVariedad(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim vClien As cSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String

    B = True
    Select Case OpcionListado
        Case 3, 30
            '3 - Factura de Anticipos bodega
            '30 - Factura de anticipos almazara
            If B Then
                If txtCodigo(6).Text = "" Or txtCodigo(7) = "" Then
                    MsgBox "Debe introducir obligatoriamente el rango de fechas.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(6)
                End If
            End If
            If B Then
                If txtCodigo(15).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de Anticipo.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(15)
                End If
            End If
            If B Then
                '[Monica]20/06/2017: control de fechas que antes no estaba
                ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(15)))
                If ResultadoFechaContaOK > 0 Then
                    If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                    B = False
                End If
            End If
            
       Case 2, 20  'Prevision de pagos bodega y almazara
            If B Then
                If txtCodigo(6).Text = "" Or txtCodigo(7) = "" Then
                    MsgBox "Para realizar la Previsión de Pago de Anticipos debe introducir obligatoriamente el rango de fechas.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(6)
                End If
            End If
       
       Case 5, 50 'Deshacer proceso de facturacion de anticipos (bodega, almazara)
            If txtCodigo(9).Text = "" Or txtCodigo(10).Text = "" Then
                MsgBox "Debe introducir la primera y última factura de la Facturación de Anticipos", vbExclamation
                B = False
                PonerFoco txtCodigo(9)
            End If
            
            If B Then
                If txtCodigo(11).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Anticipo.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(11)
                End If
            End If
    
       Case 13, 130 'Prevision de pagos de liquidacion
            If B Then
                If txtCodigo(6).Text = "" Or txtCodigo(7) = "" Then
                    MsgBox "Para realizar la Previsión de Pago debe introducir obligatoriamente el rango de fechas.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(6)
                End If
            End If
            
            
       Case 14, 140 '14 =  factura de liquidaciones bodega
                    '140=  factura de liquidaciones almazara
            If B Then
                If txtCodigo(15).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de Liquidación.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(15)
                End If
            End If
            If B Then
                '[Monica]20/06/2017: control de fechas que antes no estaba
                ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(15)))
                If ResultadoFechaContaOK > 0 Then
                    If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                    B = False
                End If
            End If
    End Select
    DatosOK = B

End Function


Private Function CargarTemporalAnticiposBodega(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
'29/10/2010 añado
Dim KiloGrado As Currency

Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim cad As String
Dim HayReg As Boolean

Dim Seccion As Integer
Dim PrecInd As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposBodega = False

    If vParamAplic.SeccionBodega = "" Then
        MsgBox "No tiene asignada en parámetros la seccion de bodega. Revise.", vbExclamation
        Exit Function
    Else
        Seccion = vParamAplic.SeccionBodega
    End If


    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, rhisfruta.kgradobonif as prestimado, rhisfruta.numalbar, "
    SQL = SQL & "rprecios.precioindustria,sum(rhisfruta.kilosnet) as kilos "
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7 "


    '[Monica]10/11/2010: calculamos el grado bonificado
'    CalcularGradoBonificado ctabla, cwhere
    If Not CalcularGradoBonificadoRealizado(cTabla, cWhere) Then
        MsgBox "No se ha realizado el cálculo del grado bonificado. Revise.", vbExclamation
        Exit Function
    End If


    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto, baseimpo, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    
    If vSeccion.LeerDatos(CStr(Seccion)) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not Rs.EOF Then
        SocioAnt = Rs!Codsocio
        VarieAnt = Rs!codvarie
        NVarieAnt = Rs!nomvarie

'29/10/2010 añado
        KilosNet = 0
        KiloGrado = 0
'end
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Rs!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), CStr(Seccion)) Then
                
                '[Monica]03/02/2016: Metemos las facturas internas en Quatretonda
                If vParamAplic.Cooperativa = 7 Then
                    vPorcIva = ""
                    If vSocio.EsFactADVInt Then                                                   '[Monica]16/06/2016: antes vParamAplic.CodIvaExeADV
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSeccion.TipIvaExento, "N")
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                Else
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                
                If vPorcIva = "" Then
                    MsgBox "El iva del socio " & DBLet(Rs!Codsocio, "N") & " no existe. Revise.", vbExclamation
                    Set vSeccion = Nothing
                    Set vSocio = Nothing
                    Set Rs = Nothing
                    Exit Function
                End If
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not Rs.EOF
        '++monica:28/07/2009 añadida la segunda condicion
        If VarieAnt <> Rs!codvarie Or SocioAnt <> Rs!Codsocio Then
            If OpcionListado = 2 Then
'[Monica]29/10/2010 cambio
'                BaseImpo = Round2(KilosNet * PrecInd, 2)
                 baseimpo = Round2(KiloGrado * PrecInd, 2)
'29/10/2010 añado
                 KilosNet = KiloGrado
                 
            End If
            
            ImpoIva = Round2(baseimpo * ImporteSinFormato(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
            
            VarieAnt = Rs!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
'29/10/2010 añado
            KiloGrado = 0
        End If
        
        If Rs!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(Rs!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), CStr(Seccion)) Then
                    '[Monica]03/02/2016: Metemos las facturas internas en Quatretonda
                    If vParamAplic.Cooperativa = 7 Then
                        vPorcIva = ""
                        If vSocio.EsFactADVInt Then                                                   '[Monica]16/06/2016: antes vParamAplic.CodIvaExeADV
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSeccion.TipIvaExento, "N")
                        Else
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                    Else
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                
                    If vPorcIva = "" Then
                        MsgBox "El iva del socio " & DBLet(Rs!Codsocio, "N") & " no existe. Revise.", vbExclamation
                        Set vSeccion = Nothing
                        Set vSocio = Nothing
                        Set Rs = Nothing
                        Exit Function
                    End If
                
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(Rs!Kilos, "N")
        
        If OpcionListado = 2 Then ' anticipo de bodega
'            baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * RS!precioindustria, 2)

'[Monica]29/10/2010: antes era los kilos por el precio en bodega, ahora lo quieren kilogrado por precio
'quito
'            PrecInd = RS!precioindustria
'añado
            KiloGrado = KiloGrado + Round2(DBLet(Rs!Kilos, "N") * Rs!PrEstimado, 2)
            PrecInd = Rs!precioindustria

        Else
            ' anticipo de almazara
            baseimpo = baseimpo + Round2(DBLet(Rs!Kilos, "N") * Rs!PrEstimado / 100 * Rs!precioindustria, 2)
        End If
            
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        If OpcionListado = 2 Then
'29/10/2010 cambio
'            BaseImpo = Round2(KilosNet * PrecInd, 2)
             baseimpo = Round2(KiloGrado * PrecInd, 2)
             KilosNet = KiloGrado
        End If
        
        ImpoIva = Round2(baseimpo * ImporteSinFormato(vPorcIva) / 100, 2)
    
        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                PorcReten = 0
        End Select
    
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalAnticiposBodega = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalAnticiposAlmazara(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
'29/10/2010 añado
Dim KiloGrado As Currency

Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim cad As String
Dim HayReg As Boolean

Dim Seccion As Integer
Dim PrecInd As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposAlmazara = False

    If vParamAplic.SeccionAlmaz = "" Then
        MsgBox "No tiene asignada en parámetros la seccion de almazara. Revise.", vbExclamation
        Exit Function
    Else
        Seccion = vParamAplic.SeccionAlmaz
    End If

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, rhisfruta.prestimado, rhisfruta.numalbar, "
    SQL = SQL & "rprecios.precioindustria,sum(rhisfruta.kilosnet) as kilos "
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7 "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto, baseimpo, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    
    If vSeccion.LeerDatos(CStr(Seccion)) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not Rs.EOF Then
        SocioAnt = Rs!Codsocio
        VarieAnt = Rs!codvarie
        NVarieAnt = Rs!nomvarie

'29/10/2010 añado
        KilosNet = 0
        KiloGrado = 0
'end
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Rs!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), CStr(Seccion)) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                If vPorcIva = "" Then
                    MsgBox "El iva del socio " & DBLet(Rs!Codsocio, "N") & " no existe. Revise.", vbExclamation
                    Set vSeccion = Nothing
                    Set vSocio = Nothing
                    Set Rs = Nothing
                    Exit Function
                End If
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not Rs.EOF
        '++monica:28/07/2009 añadida la segunda condicion
        If VarieAnt <> Rs!codvarie Or SocioAnt <> Rs!Codsocio Then
            If OpcionListado = 2 Then
'[Monica]29/10/2010 cambio
'                BaseImpo = Round2(KilosNet * PrecInd, 2)
                 baseimpo = Round2(KiloGrado * PrecInd, 2)
'29/10/2010 añado
                 KilosNet = KiloGrado
                 
            End If
            
            ImpoIva = Round2(baseimpo * ImporteSinFormato(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
            
            VarieAnt = Rs!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
'29/10/2010 añado
            KiloGrado = 0
        End If
        
        If Rs!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(Rs!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), CStr(Seccion)) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                
                    If vPorcIva = "" Then
                        MsgBox "El iva del socio " & DBLet(Rs!Codsocio, "N") & " no existe. Revise.", vbExclamation
                        Set vSeccion = Nothing
                        Set vSocio = Nothing
                        Set Rs = Nothing
                        Exit Function
                    End If
                
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(Rs!Kilos, "N")
        
        If OpcionListado = 2 Then ' anticipo de bodega
'            baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * RS!precioindustria, 2)

'[Monica]29/10/2010: antes era los kilos por el precio en bodega, ahora lo quieren kilogrado por precio
'quito
'            PrecInd = RS!precioindustria
'añado
            KiloGrado = KiloGrado + Round2(DBLet(Rs!Kilos, "N") * Rs!PrEstimado, 2)
            PrecInd = Rs!precioindustria

        Else
            ' anticipo de almazara
            baseimpo = baseimpo + Round2(DBLet(Rs!Kilos, "N") * Rs!PrEstimado / 100 * Rs!precioindustria, 2)
        End If
            
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        If OpcionListado = 2 Then
'29/10/2010 cambio
'            BaseImpo = Round2(KilosNet * PrecInd, 2)
             baseimpo = Round2(KiloGrado * PrecInd, 2)
             KilosNet = KiloGrado
        End If
        
        ImpoIva = Round2(baseimpo * ImporteSinFormato(vPorcIva) / 100, 2)
    
        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                PorcReten = 0
        End Select
    
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalAnticiposAlmazara = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function TotalFacturas(cTabla As String, cWhere As String) As Long
Dim SQL As String

    TotalFacturas = 0
    
    SQL = "SELECT  count(distinct rhisfruta.codsocio) "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If

    TotalFacturas = TotalRegistros(SQL)

End Function


Private Sub ActivarCLAVE()
Dim I As Integer
    
    For I = 9 To 11
        txtCodigo(I).Enabled = False
    Next I
    txtCodigo(8).Enabled = True
    imgFec(6).Enabled = False
    CmdAcepDesF.Enabled = False
    cmdCancelDesF.Enabled = True
    Combo1(1).Enabled = False
End Sub

Private Sub DesactivarCLAVE()
Dim I As Integer

    For I = 9 To 11
        txtCodigo(I).Enabled = True
    Next I
    txtCodigo(8).Enabled = False
    imgFec(6).Enabled = True
    CmdAcepDesF.Enabled = True
    Combo1(1).Enabled = True
End Sub

Private Sub CargaCombo()

    On Error GoTo ErrCarga
    
    
    ' Tipo de facturacion venta campo (anticipo o liquidacion)
    ' para deshacer proceso de facturacion de venta campo
    Combo1(1).Clear
    Combo1(1).AddItem "Anticipo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Liquidación"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    'recolectado por
    Combo1(2).Clear
    Combo1(2).AddItem "Cooperativa"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Socio"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Todos"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
    
    'agrupado por
    Combo1(3).Clear
    Combo1(3).AddItem "Socio"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 0
    Combo1(3).AddItem "Variedad"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 1
    
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub


Private Function NroTotalMovimientos(Tipo As Byte) As Long
' Tipo: 1 - anticipos
'       2 - liquidacion
'       3 - anticipos venta campo
'       4 - liquidacion venta campo
Dim SQL As String
    
    SQL = "select distinct "
    Select Case Tipo
        Case 1
            SQL = SQL & " CodTipomAnt "
        Case 2
            SQL = SQL & " codtipomliq "
        Case 3
            SQL = SQL & " codtipomantvc "
        Case 4
            SQL = SQL & " codtipomliqvc "
        Case 7
            SQL = SQL & " CodTipomAntAlmz "
        Case 8
            SQL = SQL & " CodTipomLiqAlmz "
        Case 9
            SQL = SQL & " CodTipomAntBod "
        Case 10
            SQL = SQL & " CodTipomLiqBod "
    End Select
    
    SQL = SQL & " from rcoope, usuarios.stipom stipom "
    SQL = SQL & " WHERE stipom.tipodocu=" & Tipo
    SQL = SQL & " and stipom.codtipom = rcoope."
    Select Case Tipo
        Case 1
            SQL = SQL & "CodTipomAnt "
        Case 2
            SQL = SQL & "codtipomliq "
        Case 3
            SQL = SQL & "codtipomantvc "
        Case 4
            SQL = SQL & "codtipomliqvc "
        Case 7
            SQL = SQL & "CodTipomAntalmz "
        Case 8
            SQL = SQL & "CodTipomliqalmz "
        Case 9
            SQL = SQL & "CodTipomAntBod "
        Case 10
            SQL = SQL & "CodTipomLiqbod "
    End Select
    
    NroTotalMovimientos = TotalRegistrosConsulta(SQL)

End Function


Private Function ActualizarRegistrosFac(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim SQL As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistrosFac = False
    SQL = "update " & cTabla & ", usuarios.stipom set impreso = 1 "
    SQL = SQL & " where usuarios.stipom.codtipom = rfactsoc.codtipom "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " and " & cWhere
    End If
    
    conn.Execute SQL
    
    ActualizarRegistrosFac = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function


Private Function ComprobarTiposIVA(tabla As String, cSelect As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim vSeccion As CSeccion
Dim B As Boolean


    On Error GoTo eComprobarTiposIVA

    ComprobarTiposIVA = False
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            
            SQL = "select distinct codiva from rsocios_seccion where codsecci = " & vParamAplic.Seccionhorto
            SQL = SQL & " and codsocio in (select rhisfruta.codsocio from " & Trim(tabla) & " where " & Trim(cSelect) & ")"
            SQL = SQL & " group by 1 "
            SQL = SQL & " order by 1 "
            
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            B = True
            
            While Not Rs.EOF And B
                If DBLet(Rs.Fields(0).Value, "N") = 0 Then
                    B = False
                    MsgBox "Hay socios sin iva en la sección hortofrutícola. Revise.", vbExclamation
                Else
                    SQL = ""
                    SQL = DevuelveDesdeBDNew(cConta, "tiposiva", "codigiva", "codigiva", DBLet(Rs.Fields(0).Value, "N"), "N")
                    If SQL = "" Then
                        B = False
                        MsgBox "No existe el codigo de iva " & DBLet(Rs.Fields(0).Value, "N") & ". Revise.", vbExclamation
                    End If
                End If
            
                Rs.MoveNext
            Wend
        
            Set Rs = Nothing
        
            ComprobarTiposIVA = B
        
            vSeccion.CerrarConta
            
            Set vSeccion = Nothing
        End If
    End If
    Exit Function
    
eComprobarTiposIVA:
    MuestraError Err.Number, "Comprobar Tipos Iva", Err.Description
End Function





Private Function CargarTemporalLiquidacionBodega(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CampoAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bruto As Currency
Dim ImpoIva As Currency
Dim ImpoGastos As Currency
Dim ImpoBonif As Currency '09/09/2009: las bonificaciones las quitamos de los gastos
Dim ImpoReten As Currency
Dim ImpoAport As Currency
Dim Anticipos As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim vPorcGasto As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim cad As String
Dim HayReg As Boolean
Dim vGastos As Currency

Dim EsComplementaria As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionBodega = False

    EsComplementaria = (Check1(0).Value = 1)

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
    SQL = SQL & "rprecios.precioindustria, rprecios.tipofact, rhisfruta.kgradobonif as prestimado, rhisfruta.numalbar,  sum(rhisfruta.kilosnet) as kilos "
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6,7,8 " '30/04/2010 añadido precio estimado y numalbar
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6,7,8 "

    '[Monica]23/11/2012: añadida lo de si es complementaria
    If Not EsComplementaria Then
    '[Monica]10/11/2010: calculamos el grado bonificado
    '    CalcularGradoBonificado ctabla, cwhere
        If Not CalcularGradoBonificadoRealizado(cTabla, cWhere) Then
            MsgBox "No se ha realizado el cálculo del grado bonificado. Revise.", vbExclamation
            Exit Function
        End If
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  bonificacion, gastos,  anticipos, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionBodega) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not Rs.EOF Then
        SocioAnt = Rs!Codsocio
        VarieAnt = Rs!codvarie
        NVarieAnt = Rs!nomvarie
        CampoAnt = Rs!codcampo
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Rs!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.SeccionBodega) Then
                
                '[Monica]03/02/2016: Metemos las facturas internas en Quatretonda
                If vParamAplic.Cooperativa = 7 Then
                    vPorcIva = ""
                    If vSocio.EsFactADVInt Then                                                 '[Monica]16/06/2016: antes estaba el porcentaje de iva del adv
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSeccion.TipIvaExento, "N")
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                Else
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                
                vPorcGasto = ""
                vPorcGasto = vParamAplic.PorcGtoMantBOD
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not Rs.EOF
        If CampoAnt <> Rs!codcampo Or VarieAnt <> Rs!codvarie Or SocioAnt <> Rs!Codsocio Then
            '[Monica]23/11/2012: añadida lo de si es complementaria
            If Not EsComplementaria Then
                Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos from rhisfruta_gastos, rhisfruta  "
                Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
                Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
                Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
                Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtCodigo(6).Text, "F")
                Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtCodigo(7).Text, "F")
                Sql4 = Sql4 & " and rhisfruta.numalbar = rhisfruta_gastos.numalbar "
                
                ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
            End If
            
            CampoAnt = Rs!codcampo
        End If
    
        ' 23/07/2009: añadido el or con la segunda condicion
        If VarieAnt <> Rs!codvarie Or SocioAnt <> Rs!Codsocio Then
            
            '[Monica]23/11/2012: añadida lo de si es complementaria
            Anticipos = 0
            If Not EsComplementaria Then
                ' anticipos
                Sql4 = "select sum(rfactsoc_variedad.imporvar) "
                Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntBod, "T") ' "FAA"
                Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
                Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
                Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
                
                Anticipos = DevuelveValor(Sql4)
            End If
            
            Bruto = baseimpo
            ImpoGastos = ImpoGastos
            
            '[Monica]23/11/2012: añadida lo de si es complementaria
            If Not EsComplementaria Then
                '[Monica] 09/09/2009: el gasto de la cooperativa lo añado a la columna de gastos que no usa Valsur
                ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
            End If
            
            baseimpo = baseimpo - ImpoGastos - Anticipos
            
            ImpoIva = Round2((baseimpo) * ComprobarCero(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    PorcReten = 0
            End Select
        
'            ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten '- ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto, "N") & ","
            Sql1 = Sql1 & DBSet(0, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
            
            VarieAnt = Rs!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
            
            ImpoGastos = 0
            Anticipos = 0
            
        End If
        
        If Rs!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(Rs!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.SeccionBodega) Then
                    '[Monica]03/02/2016: Metemos las facturas internas en Quatretonda
                    If vParamAplic.Cooperativa = 7 Then
                        vPorcIva = ""
                        If vSocio.EsFactADVInt Then                                                   '[Monica]16/06/2016: antes tipo de iva exento de parametros adv
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSeccion.TipIvaExento, "N")
                        Else
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                    Else
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                    
                    vPorcGasto = ""
                    vPorcGasto = vParamAplic.PorcGtoMantBOD  'DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(Rs!Kilos, "N")
        
        '[Monica]23/11/2012: añadida lo de si es complementaria ( se calculara precio por kilo en lugar de por kilogrado )
        If Not EsComplementaria Then
            baseimpo = baseimpo + Round2(DBLet(Rs!Kilos, "N") * Rs!precioindustria * Rs!PrEstimado, 2)  'Round2(DBLet(Rs!Kilos, "N") * Rs!precioindustria, 2)
        Else
            baseimpo = baseimpo + Round2(DBLet(Rs!Kilos, "N") * Rs!precioindustria, 2)  'Round2(DBLet(Rs!Kilos, "N") * Rs!precioindustria, 2)
        End If
            
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        Bruto = baseimpo
        
        
        '[Monica]23/11/2012: añadida lo de si es complementaria
        Anticipos = 0
        If Not EsComplementaria Then
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntBod, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
            
            Anticipos = DevuelveValor(Sql4)
        End If
        
        '[Monica]23/11/2012: añadida lo de si es complementaria
        ImpoGastos = 0
        If Not EsComplementaria Then
            ' gastos
            Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos from rhisfruta_gastos, rhisfruta "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtCodigo(6).Text, "F")
            Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtCodigo(7).Text, "F")
            Sql4 = Sql4 & " and rhisfruta.numalbar = rhisfruta_gastos.numalbar "
                
            ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
                
            '[Monica] 09/09/2009: el gasto de la cooperativa lo añado a la columna de gastos que no usa Valsur
            ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
        End If
        
        baseimpo = baseimpo - ImpoGastos - Anticipos
        
        ImpoIva = Round2((baseimpo) * ComprobarCero(vPorcIva) / 100, 2)
        
        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                PorcReten = 0
        End Select
    
'        ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
    
        TotalFac = baseimpo + ImpoIva - ImpoReten ' - ImpoAport
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
        Sql1 = Sql1 & DBSet(Bruto, "N") & ","
        Sql1 = Sql1 & DBSet(0, "N") & ","
        Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
        Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'        Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalLiquidacionBodega = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalLiquidacionAlmazara(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CampoAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bruto As Currency
Dim ImpoIva As Currency
Dim ImpoGastos As Currency
Dim ImpoBonif As Currency '09/09/2009: las bonificaciones las quitamos de los gastos
Dim ImpoReten As Currency
Dim ImpoAport As Currency
Dim Anticipos As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim vPorcGasto As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim cad As String
Dim HayReg As Boolean
Dim vGastos As Currency
Dim campo As String


    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionAlmazara = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, "
    SQL = SQL & "rprecios.precioindustria, rprecios.tipofact, rhisfruta.prestimado, sum(rhisfruta.kilosnet) as kilos "
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6 "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  bonificacion, gastos,  anticipos, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    ' en almazara no se insertan campos: metemos el minimo codcampo sin condiciones
    campo = "0" 'DevuelveValor("select min(codcampo) from rcampos")

    HayReg = False
    If Not Rs.EOF Then
        SocioAnt = Rs!Codsocio
        VarieAnt = Rs!codvarie
        NVarieAnt = Rs!nomvarie
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Rs!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.SeccionAlmaz) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not Rs.EOF
    
    
        ' 23/07/2009: añadido el or con la segunda condicion
        If VarieAnt <> Rs!codvarie Or SocioAnt <> Rs!Codsocio Then
            
            Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos from rhisfruta_gastos, rhisfruta  "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtCodigo(6).Text, "F")
            Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtCodigo(7).Text, "F")
            Sql4 = Sql4 & " and rhisfruta.numalbar = rhisfruta_gastos.numalbar "
            
            ImpoGastos = DevuelveValor(Sql4)
            
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
            
            Anticipos = DevuelveValor(Sql4)
            
            Bruto = baseimpo
            
            baseimpo = baseimpo - ImpoGastos - Anticipos
            
            ImpoIva = Round2((baseimpo) * ComprobarCero(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    PorcReten = 0
            End Select
        
'            ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten '- ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto, "N") & ","
            Sql1 = Sql1 & DBSet(0, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
            
            VarieAnt = Rs!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
            
            ImpoGastos = 0
            Anticipos = 0
            
        End If
        
        If Rs!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(Rs!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.SeccionAlmaz) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        
        KilosNet = KilosNet + DBLet(Rs!Kilos, "N")
        
        baseimpo = baseimpo + Int((DBLet(Rs!Kilos, "N") * Rs!precioindustria * DBLet(Rs!PrEstimado, "N") / 100) * 100) / 100 'Round2(DBLet(Rs!Kilos, "N") * Rs!precioindustria, 2)
            
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        Bruto = baseimpo
        
        
        ' anticipos
        Sql4 = "select sum(rfactsoc_variedad.imporvar) "
        Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
        Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
        Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' "FAA"
        Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
        
        Anticipos = DevuelveValor(Sql4)
            
        
        ' gastos
        Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos from rhisfruta_gastos, rhisfruta "
        Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtCodigo(6).Text, "F")
        Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtCodigo(7).Text, "F")
        Sql4 = Sql4 & " and rhisfruta.numalbar = rhisfruta_gastos.numalbar "
            
        ImpoGastos = DevuelveValor(Sql4)
                
        baseimpo = baseimpo - ImpoGastos - Anticipos
        
        ImpoIva = Round2((baseimpo) * ComprobarCero(vPorcIva) / 100, 2)
        
        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                PorcReten = 0
        End Select
    
'        ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
    
        TotalFac = baseimpo + ImpoIva - ImpoReten ' - ImpoAport
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
        Sql1 = Sql1 & DBSet(Bruto, "N") & ","
        Sql1 = Sql1 & DBSet(0, "N") & ","
        Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
        Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'        Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalLiquidacionAlmazara = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalLiquidacionAlmazaraValsur(cTabla As String, cWhere As String, FIni As String, FFin As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CampoAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bruto As Currency
Dim ImpoIva As Currency
Dim ImpoGastos As Currency
Dim ImpoBonif As Currency '09/09/2009: las bonificaciones las quitamos de los gastos
Dim ImpoReten As Currency
Dim ImpoAport As Currency
Dim Anticipos As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim vPorcGasto As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim cad As String
Dim HayReg As Boolean
Dim vGastos As Currency
Dim campo As String

Dim LitrosConsumidos As Currency
Dim LitrosProducidos As Currency

Dim PrecioConsumido As Currency
Dim PrecioProducido As Currency

Dim Rdto As Currency
Dim KilosConsu As Long
Dim KilosComer As Long

Dim PrecioRetirado As Currency
Dim ImporteRetirado As Currency
Dim ImporteMoltura As Currency
Dim ImporteMoltura1 As Currency
Dim ImporteEnvasado As Currency

Dim Importe As Currency
Dim PrecioMoltura As Currency
Dim PrecioEnvasado As Currency

Dim Sql5 As String
Dim Sql3 As String

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionAlmazaraValsur = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
                                        '%%%%%
    SQL = "SELECT rhisfruta.codsocio, variedades.codclase codvarie, variedades.nomvarie, rhisfruta.numalbar, "
    SQL = SQL & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios.tipofact, rhisfruta.prestimado, sum(rhisfruta.kilosnet) as kilos "
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8 "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  bonificacion, gastos,  anticipos, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    ' en almazara no se insertan campos: metemos el minimo codcampo sin condiciones
    campo = "0" 'DevuelveValor("select min(codcampo) from rcampos")

    HayReg = False
    If Not Rs.EOF Then
        SocioAnt = Rs!Codsocio
        VarieAnt = Rs!codvarie
        NVarieAnt = Rs!nomvarie
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Rs!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.SeccionAlmaz) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                
                '[Monica] 05/07/2010 : tiene gasto de cooperativa
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not Rs.EOF
    
    
        ' 23/07/2009: añadido el or con la segunda condicion
        If VarieAnt <> Rs!codvarie Or SocioAnt <> Rs!Codsocio Then
            Sql3 = "select rbodalbaran_variedad.*, variedades.* "
            Sql3 = Sql3 & " from rbodalbaran_variedad, rbodalbaran, variedades where rbodalbaran.codsocio = " & DBSet(SocioAnt, "N")
            Sql3 = Sql3 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
            Sql3 = Sql3 & " and variedades.codclase = " & DBSet(VarieAnt, "N")
            Sql3 = Sql3 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
            Sql3 = Sql3 & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
            Sql3 = Sql3 & " order by rbodalbaran_variedad.numalbar desc, rbodalbaran_variedad.numlinea desc"
        
            ' litros consumidos a otro precio
            Sql4 = "select rbodalbaran_variedad.codvarie, variedades.eurdesta, variedades.eursegsoc, sum(cantidad) cantidad, round(variedades.eurdesta * sum(cantidad), 2) importevta, round(variedades.eursegsoc * sum(cantidad), 2) importeenv  "
            Sql4 = Sql4 & " from rbodalbaran_variedad, rbodalbaran, variedades where rbodalbaran.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
            '[Monica]10/03/2016: ahora jugamos con la clase
            Sql4 = Sql4 & " and variedades.codclase =  " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
            Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
            Sql4 = Sql4 & " group by 1,2,3 "
            Sql4 = Sql4 & " order by 1,2,3 "
            
            Sql5 = "select sum(cantidad) from (" & Sql4 & ") aaaaa"
            
            LitrosConsumidos = DevuelveValor(Sql5)
            
'Ahora 01/04/2011
            If LitrosProducidos > LitrosConsumidos Then
'                Sql4 = "select min(precioar) from rbodalbaran_variedad, rbodalbaran, variedades where rbodalbaran.codsocio = " & DBSet(SocioAnt, "N")
'                Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
'                '[Monica]10/03/2016: ahora jugamos con la clase
'                Sql4 = Sql4 & " and variedades.codclase = " & DBSet(VarieAnt, "N")
'                Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
'                Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
'
'                PrecioRetirado = DevuelveValor(Sql4)
            
                Rdto = Round2(LitrosProducidos * 100 / KilosNet, 4)
                
                KilosComer = Round2((LitrosProducidos - LitrosConsumidos) * 100 / Rdto, 0)
                KilosConsu = KilosNet - KilosComer
                
'                ImporteRetirado = Round2(LitrosConsumidos * PrecioRetirado, 2)
'                Sql5 = "select sum(importevta) from (" & Sql4 & ") aaaaa"
'                ImporteRetirado = DevuelveValor(Sql5)
                ImporteRetirado = CalculoImporteRetirado(Sql3, CStr(LitrosConsumidos), False)
                
                PrecioMoltura = DevuelveValor("select eurmanob from variedades where codvarie = " & DBSet(VarieAnt, "N"))
                
                ImporteMoltura = Round2(KilosConsu * PrecioMoltura, 2)
                ImporteMoltura1 = Round2(KilosComer * PrecioMoltura, 2)
                
'                ImporteEnvasado = Round2(LitrosConsumidos * vParamAplic.GtoEnvasado, 2)
'                Sql5 = "select sum(importeenv) from (" & Sql4 & ") aaaaa"
'                ImporteEnvasado = DevuelveValor(Sql5) 'Round2(LitrosConsumidos * vParamAplic.GtoEnvasado, 2)
                ImporteEnvasado = CalculoImporteRetirado(Sql3, LitrosConsumidos, True)

                
                baseimpo = baseimpo + Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado) + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido), 2) - ImporteMoltura1
                Importe = Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado) + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido) - ImporteMoltura1, 2)
            
                Bruto = ImporteRetirado + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido)
            
            Else
'                Sql4 = "select min(precioar) from rbodalbaran_variedad, rbodalbaran where rbodalbaran.codsocio = " & DBSet(SocioAnt, "N")
'                Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
'                Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = " & DBSet(VarieAnt, "N")
'                Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
'
'                PrecioRetirado = DevuelveValor(Sql4)
                
'                Sql5 = "select eurdesta from (" & Sql4 & ") aaaaa"
'                PrecioRetirado = DevuelveValor(Sql5)
                
                Rdto = Round2(LitrosProducidos * 100 / KilosNet, 4)
                
                KilosConsu = Round2(LitrosProducidos * 100 / Rdto, 0)
                
'                ImporteRetirado = Round2(LitrosProducidos * PrecioRetirado, 2)
                ImporteRetirado = CalculoImporteRetirado(Sql3, LitrosProducidos, False)
                PrecioRetirado = Round2(ImporteRetirado / LitrosProducidos, 4)

                PrecioMoltura = DevuelveValor("select eurmanob from variedades where codvarie = " & DBSet(VarieAnt, "N"))
                
                ImporteMoltura = Round2(KilosConsu * PrecioMoltura, 2)
'                ImporteEnvasado = Round2(LitrosProducidos * vParamAplic.GtoEnvasado, 2)
                ImporteEnvasado = CalculoImporteRetirado(Sql3, LitrosProducidos, True)
                PrecioEnvasado = Round2(ImporteEnvasado / LitrosProducidos, 4)

                
                baseimpo = baseimpo + Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado), 2)
                Importe = Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado), 2)
            
                Bruto = ImporteRetirado
            End If
'fahora 01/04/2011
            
            
            Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos from rhisfruta_gastos, rhisfruta  "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtCodigo(6).Text, "F")
            Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtCodigo(7).Text, "F")
            Sql4 = Sql4 & " and rhisfruta.numalbar = rhisfruta_gastos.numalbar "
            
            ImpoGastos = DevuelveValor(Sql4)
            
            
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
            
            Anticipos = DevuelveValor(Sql4)
            
            
'[Monica]14/04/2011: ahora lo vuelven a querer
'[Monica]07/04/2011: ahora no lo quieren en linea va en el precio
'            '[Monica] 05/07/2010: el gasto de la cooperativa lo añado a la columna de gastos que no usa Valsur
            ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
            
            baseimpo = baseimpo - ImpoGastos - Anticipos
            
            ImpoIva = Round2((baseimpo) * ComprobarCero(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    PorcReten = 0
            End Select
        
'            ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten '- ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(baseimpo + ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(0, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
            
            VarieAnt = Rs!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
            
            ImpoGastos = 0
            Anticipos = 0
            LitrosProducidos = 0
            
        End If
        
        If Rs!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(Rs!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.SeccionAlmaz) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    
                    '[Monica] 05/07/2010 : tiene gasto de cooperativa
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        
        KilosNet = KilosNet + DBLet(Rs!Kilos, "N")
        
        LitrosProducidos = LitrosProducidos + Round2(DBLet(Rs!Kilos, "N") * DBLet(Rs!PrEstimado, "N") / 100, 0)
        
        PrecioProducido = DBLet(Rs!PreSocio, "N")
        PrecioConsumido = DBLet(Rs!PreCoop, "N")
        
'        baseimpo = baseimpo + Int((DBLet(RS!Kilos, "N") * RS!precioindustria * DBLet(RS!PrEstimado, "N") / 100) * 100) / 100 'Round2(DBLet(Rs!Kilos, "N") * Rs!precioindustria, 2)
            
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        Sql3 = "select rbodalbaran_variedad.*, variedades.* "
        Sql3 = Sql3 & " from rbodalbaran_variedad, rbodalbaran, variedades where rbodalbaran.codsocio = " & DBSet(SocioAnt, "N")
        Sql3 = Sql3 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
        Sql3 = Sql3 & " and variedades.codclase = " & DBSet(VarieAnt, "N")
        Sql3 = Sql3 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
        Sql3 = Sql3 & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
        Sql3 = Sql3 & " order by rbodalbaran_variedad.numalbar desc, rbodalbaran_variedad.numlinea desc"
        
        ' litros consumidos a otro precio
        Sql4 = "select rbodalbaran_variedad.codvarie, variedades.eurdesta, variedades.eursegsoc, sum(cantidad) cantidad, round(variedades.eurdesta * sum(cantidad), 2) importevta, round(variedades.eursegsoc * sum(cantidad), 2) importeenv "
        Sql4 = Sql4 & " from rbodalbaran_variedad, rbodalbaran, variedades where rbodalbaran.codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
        Sql4 = Sql4 & " and variedades.codclase = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
        Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
        Sql4 = Sql4 & " group by 1,2  order by 1,2 "
        
        Sql5 = "select sum(cantidad) from (" & Sql4 & ") aaaaa"
        
        LitrosConsumidos = DevuelveValor(Sql5)
            
'antes
'        If LitrosProducidos > LitrosConsumidos Then
'            BaseImpo = Round2((LitrosConsumidos * PrecioConsumido) + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido), 2)
'        Else
'            BaseImpo = Round2(LitrosProducidos * PrecioConsumido, 2)
'        End If
        
'Ahora 01/04/2011
        If LitrosProducidos > LitrosConsumidos Then
'            Sql4 = "select min(precioar) from rbodalbaran_variedad, rbodalbaran where rbodalbaran.codsocio = " & DBSet(SocioAnt, "N")
'            Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
'            Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = " & DBSet(VarieAnt, "N")
'            Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
'
'            PrecioRetirado = DevuelveValor(Sql4)
        
            Rdto = Round2(LitrosProducidos * 100 / KilosNet, 4)
            
            KilosComer = Round2((LitrosProducidos - LitrosConsumidos) * 100 / Rdto, 0)
            KilosConsu = KilosNet - KilosComer
            
'            ImporteRetirado = Round2(LitrosConsumidos * PrecioRetirado, 2)
'            Sql5 = "select sum(importevta) from (" & Sql5 & ") aaaaa"
'            ImporteRetirado = DevuelveValor(Sql5)
            ImporteRetirado = CalculoImporteRetirado(Sql3, CStr(LitrosConsumidos), False)
            
            PrecioMoltura = DevuelveValor("select eurmanob from variedades where codvarie = " & DBSet(VarieAnt, "N"))
            
            ImporteMoltura = Round2(KilosConsu * PrecioMoltura, 2)
            ImporteMoltura1 = Round2(KilosComer * PrecioMoltura, 2)
            
'            ImporteEnvasado = Round2(LitrosConsumidos * vParamAplic.GtoEnvasado, 2)
'            Sql5 = "select sum(importeenv) from (" & Sql4 & ") aaaaa"
'            ImporteEnvasado = DevuelveValor(Sql5) 'Round2(LitrosConsumidos * vParamAplic.GtoEnvasado, 2)
            ImporteEnvasado = CalculoImporteRetirado(Sql3, LitrosConsumidos, True)

            
            baseimpo = baseimpo + Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado) + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido), 2) - ImporteMoltura1
            Importe = Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado) + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido) - ImporteMoltura1, 2)
            
            Bruto = ImporteRetirado + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido)
        Else
'            Sql4 = "select min(precioar) from rbodalbaran_variedad, rbodalbaran where rbodalbaran.codsocio = " & DBSet(SocioAnt, "N")
'            Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
'            Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = " & DBSet(VarieAnt, "N")
'            Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
'
'            PrecioRetirado = DevuelveValor(Sql4)
'            Sql5 = "select eurdesta from (" & Sql4 & ") aaaaa"
'            PrecioRetirado = DevuelveValor(Sql5)

            
            Rdto = Round2(LitrosProducidos * 100 / KilosNet, 4)
            
            KilosConsu = Round2(LitrosProducidos * 100 / Rdto, 0)
            
'            ImporteRetirado = Round2(LitrosProducidos * PrecioRetirado, 2)
            ImporteRetirado = CalculoImporteRetirado(Sql3, LitrosProducidos, False)
            PrecioRetirado = Round2(ImporteRetirado / LitrosProducidos, 4)
            
            PrecioMoltura = DevuelveValor("select eurmanob from variedades where codvarie = " & DBSet(VarieAnt, "N"))
            
            ImporteMoltura = Round2(KilosConsu * PrecioMoltura, 2)
'            ImporteEnvasado = Round2(LitrosProducidos * vParamAplic.GtoEnvasado, 2)
            ImporteEnvasado = CalculoImporteRetirado(Sql3, LitrosProducidos, True)
            PrecioEnvasado = Round2(ImporteEnvasado / LitrosProducidos, 4)
            
            baseimpo = baseimpo + Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado), 2)
            Importe = Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado), 2)
        
            Bruto = ImporteRetirado
        End If
'fahora 01/04/2011
        
        
        
'        Bruto = BaseImpo
        
        ' anticipos
        Sql4 = "select sum(rfactsoc_variedad.imporvar) "
        Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
        Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
        Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' "FAA"
        Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
        
        Anticipos = DevuelveValor(Sql4)
            
        
        ' gastos
        Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos from rhisfruta_gastos, rhisfruta "
        Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtCodigo(6).Text, "F")
        Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtCodigo(7).Text, "F")
        Sql4 = Sql4 & " and rhisfruta.numalbar = rhisfruta_gastos.numalbar "
            
        ImpoGastos = DevuelveValor(Sql4)

'[Monica]14/04/2011: ahora lo vuelven a querer
'[Monica]07/04/2011: ahora no lo quieren en linea va en el precio
'        '[Monica] 05/07/2010: el gasto de la cooperativa lo añado a la columna de gastos que no usa Valsur
        ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
                
        baseimpo = baseimpo - ImpoGastos - Anticipos
        
        ImpoIva = Round2((baseimpo) * ComprobarCero(vPorcIva) / 100, 2)
        
        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                PorcReten = 0
        End Select
    
'        ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
    
        TotalFac = baseimpo + ImpoIva - ImpoReten ' - ImpoAport
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
        Sql1 = Sql1 & DBSet(baseimpo + ImpoGastos, "N") & ","
        Sql1 = Sql1 & DBSet(0, "N") & ","
        Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
        Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'        Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalLiquidacionAlmazaraValsur = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function






Private Function HayPreciosVariedadesBodegaAlmazara(Tipo As Byte, cTabla As String, cWhere As String, Cooperativa As Integer) As Boolean
'Comprobar si hay precios para cada una de las variedades seleccionadas
Dim SQL As String
Dim vPrecios As CPrecios
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim B As Boolean
Dim Sql2 As String
Dim Sql5 As String
Dim VarieAnt As Long
Dim NumReg As Long

    On Error GoTo eHayPreciosVariedadesBodegaAlmazara
    
    HayPreciosVariedadesBodegaAlmazara = False
    
    conn.Execute " DROP TABLE IF EXISTS tmpVarie;"
    
    SQL = "CREATE TEMPORARY TABLE tmpVarie ( " 'TEMPORARY
    SQL = SQL & "codvarie INT(6) UNSIGNED  DEFAULT '0' NOT NULL) "
    conn.Execute SQL
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    SQL = "Select distinct rhisfruta.codvarie, rhisfruta.fecalbar FROM " & QuitarCaracterACadena(cTabla, "_1")
    
'    Sql2 = "Select distinct rhisfruta_clasif.codvarie, rhisfruta_clasif.codcalid FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
'        Sql2 = Sql2 & " where " & cWhere
    End If
    SQL = SQL & " group by 1,2 "
    SQL = SQL & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    B = True
    
    If Not Rs.EOF Then VarieAnt = DBLet(Rs!codvarie, "N")
    NumReg = 0
    ' comprobamos que existen registros para todos las variedades seleccionadas
    While Not Rs.EOF And B
        If vParamAplic.Cooperativa = 1 And Tipo = 1 Then
            ' si estamos en liquidacion de valsur
            ' obligamos a que metan linea en rprecios_calidad
            ' precio cooperativa = precio consumido /  Precio socio = precio no consumido
            Sql2 = "select * from rprecios where (codvarie, tipofact, contador) = ("
            Sql2 = Sql2 & "SELECT rprecios.codvarie, rprecios.tipofact, max(rprecios.contador) FROM rprecios, rprecios_calidad WHERE rprecios.codvarie=" & DBSet(Rs!codvarie, "N") & " and "
            Sql2 = Sql2 & " rprecios.tipofact = " & DBSet(Tipo, "N") & " and fechaini <= " & DBSet(Rs!Fecalbar, "F")
            Sql2 = Sql2 & " and fechafin >= " & DBSet(Rs!Fecalbar, "F") & " and "
            Sql2 = Sql2 & " rprecios.codvarie = rprecios_calidad.codvarie and "
            Sql2 = Sql2 & " rprecios.tipofact = rprecios_calidad.tipofact and "
            Sql2 = Sql2 & " rprecios.contador = rprecios_calidad.contador "
            Sql2 = Sql2 & " group by 1, 2) "
        Else
            Sql2 = "select * from rprecios where (codvarie, tipofact, contador) = ("
            Sql2 = Sql2 & "SELECT codvarie, tipofact, max(contador) FROM rprecios WHERE codvarie=" & DBSet(Rs!codvarie, "N") & " and "
            Sql2 = Sql2 & " tipofact = " & DBSet(Tipo, "N") & " and fechaini <= " & DBSet(Rs!Fecalbar, "F")
            Sql2 = Sql2 & " and fechafin >= " & DBSet(Rs!Fecalbar, "F") & " and precioindustria <> 0 and precioindustria is not null "
            Sql2 = Sql2 & " group by 1, 2) "
        End If
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Rs2.EOF Then
            B = False
            MsgBox "No existe precio para la variedad " & DBLet(Rs!codvarie, "N") & " de fecha " & DBLet(Rs!Fecalbar, "F") & ". Revise.", vbExclamation
        Else
            Sql5 = "select count(*) from tmpvarie where codvarie = " & DBSet(Rs!codvarie, "N")
            If TotalRegistros(Sql5) = 0 Then
                Sql5 = "insert into tmpVarie (codvarie) values (" & DBSet(Rs!codvarie, "N") & ")"
                conn.Execute Sql5
            End If
        End If
            
        Set Rs2 = Nothing
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    HayPreciosVariedadesBodegaAlmazara = B
    Exit Function
    
eHayPreciosVariedadesBodegaAlmazara:
    MuestraError Err.nume, "Comprobando si hay precios de Bodega/Almazara en variedades", Err.Description
End Function



Private Function CargarTemporalBodega(Tipo As Byte, cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CampoAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bruto As Currency
Dim ImpoIva As Currency
Dim ImpoGastos As Currency
Dim ImpoReten As Currency
Dim ImpoAport As Currency
Dim Anticipos As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim vPorcGasto As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim cad As String
Dim HayReg As Boolean

Dim Contador As Long
Dim FechaIni As Date
Dim FechaFin As Date
Dim Gastos As Currency
Dim Sql3 As String
Dim Precio As Currency
Dim Importe As Currency
Dim Kilos As Currency
Dim Nregs As Long
Dim Sql5 As String

    On Error GoTo eCargarTemporal
    
    CargarTemporalBodega = False
    
    
    Sql2 = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    Sql2 = "delete from tmpliquidacion1 where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, "
    SQL = SQL & " rhisfruta.fecalbar, "
    SQL = SQL & " sum(rhisfruta.kilosnet) as kilos "
    SQL = SQL & " FROM  (" & cTabla & ") inner join tmpvarie on rhisfruta.codvarie = tmpvarie.codvarie "
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5"
    SQL = SQL & " having sum(rhisfruta.kilosnet) <> 0 "
    SQL = SQL & " order by 1, 2, 3, 4, 5"


    Nregs = TotalRegistrosConsulta(SQL)
    
    Label2(10).Caption = "Cargando Tabla Temporal"
    Me.Pb1.visible = True
    Me.Pb1.Max = Nregs
    Me.Pb1.Value = 0
    Me.Refresh

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    
    While Not Rs.EOF
    
        Label2(12).Caption = "Socio " & Rs!Codsocio & " Variedad " & Rs!codvarie & "- Campo " & Rs!codcampo
        IncrementarProgresNew Pb1, 1
        Me.Refresh
        DoEvents
    
        Sql3 = "select fechaini, fechafin, max(contador) as contador from rprecios where codvarie = " & DBSet(Rs!codvarie, "N")
        Sql3 = Sql3 & " and tipofact =  " & DBSet(Tipo, "N")
        Sql3 = Sql3 & " and fechaini <= " & DBSet(Rs!Fecalbar, "F")
        Sql3 = Sql3 & " and fechafin >= " & DBSet(Rs!Fecalbar, "F")
        Sql3 = Sql3 & " group by 1,2"
                
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS1.EOF Then
            Contador = DBLet(RS1!Contador, "N")
            FechaIni = DBLet(RS1!FechaIni, "F")
            FechaFin = DBLet(RS1!FechaFin, "F")
        End If
        Set RS1 = Nothing
        
        Sql3 = "select precioindustria from rprecios where codvarie = " & DBSet(Rs!codvarie, "N")
        Sql3 = Sql3 & " and tipofact = " & DBSet(Tipo, "N")
        Sql3 = Sql3 & " and contador = " & DBSet(Contador, "N")
        
        Precio = DevuelveValor(Sql3)
        
        
        Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos"
        Sql4 = Sql4 & "  from rhisfruta, rhisfruta_gastos "
        Sql4 = Sql4 & " where rhisfruta.codsocio = " & DBSet(Rs!Codsocio, "N") & "  and "
        Sql4 = Sql4 & " rhisfruta.codvarie = " & DBSet(Rs!codvarie, "N") & "  and "
        Sql4 = Sql4 & " rhisfruta.codcampo = " & DBSet(Rs!codcampo, "N") & " and "
        Sql4 = Sql4 & " rhisfruta.fecalbar >= " & DBSet(FechaIni, "F") & " and "
        Sql4 = Sql4 & " rhisfruta.fecalbar <= " & DBSet(FechaFin, "F") & " and "
        Sql4 = Sql4 & " rhisfruta.numalbar = rhisfruta_gastos.numalbar "
         
        Gastos = DevuelveValor(Sql4)
        
        
        Sql5 = "select count(*) from tmpliquidacion1 where codsocio = " & DBSet(Rs!Codsocio, "N") & "  and "
        Sql5 = Sql5 & " tmpliquidacion1.codvarie = " & DBSet(Rs!codvarie, "N") & "  and "
        Sql5 = Sql5 & " tmpliquidacion1.codcampo = " & DBSet(Rs!codcampo, "N") & " and "
        Sql5 = Sql5 & " tmpliquidacion1.fechaini = " & DBSet(FechaIni, "F") & " and "
        Sql5 = Sql5 & " tmpliquidacion1.fechafin = " & DBSet(FechaFin, "F") & " and "
        Sql5 = Sql5 & " tmpliquidacion1.codusu = " & vUsu.Codigo
        
        If TotalRegistros(Sql5) = 0 Then
            Sql5 = "insert into tmpliquidacion1 values (" & vUsu.Codigo & "," & DBSet(Rs!Codsocio, "N") & ","
            Sql5 = Sql5 & DBSet(Rs!codvarie, "N") & ","
            Sql5 = Sql5 & DBSet(Rs!codcampo, "N") & ","
            Sql5 = Sql5 & DBSet(FechaIni, "F") & ","
            Sql5 = Sql5 & DBSet(FechaFin, "F") & ","
            Sql5 = Sql5 & DBSet(Gastos, "N") & ")"
            
            conn.Execute Sql5
        End If

        ' si no tiene precio no insertamos en la tabla
        
'30/07/2009
'        If Precio <> 0 Then
'            Importe = Round2(Precio * DBLet(RS!kilos, "N"), 2)
            Sql2 = "select count(*) from tmpliquidacion where codusu = " & vUsu.Codigo
            Sql2 = Sql2 & " and codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(Rs!codcampo, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " and contador = " & DBSet(Contador, "N")
            Sql2 = Sql2 & " and fechaini = " & DBSet(FechaIni, "F")
            Sql2 = Sql2 & " and fechafin = " & DBSet(FechaFin, "F")
            
            If TotalRegistros(Sql2) = 0 Then
                Kilos = 0
                
                Sql3 = "insert into tmpliquidacion (codusu,codsocio,codcampo,codvarie,codcalid,contador,kilosnet,precio,importe, "
                Sql3 = Sql3 & " nomvarie, fechaini, fechafin, gastos)"
                Sql3 = Sql3 & " values (" & vUsu.Codigo & "," & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!codcampo, "N") & ","
                Sql3 = Sql3 & DBSet(Rs!codvarie, "N") & ",0," & DBSet(Contador, "N") & ","
                Sql3 = Sql3 & DBSet(Rs!Kilos, "N") & "," & DBSet(Precio, "N") & "," & DBSet(0, "N") & ","
                Sql3 = Sql3 & DBSet(Rs!nomvarie, "T") & "," & DBSet(FechaIni, "F") & ","
                Sql3 = Sql3 & DBSet(FechaFin, "F") & "," & DBSet(Gastos, "N") & ")"
                
                conn.Execute Sql3
                Kilos = Rs!Kilos
            Else
                Kilos = Kilos + Rs!Kilos
                Sql3 = "update tmpliquidacion set kilosnet = kilosnet + " & DBSet(Rs!Kilos, "N")
                Sql3 = Sql3 & " where codusu = " & vUsu.Codigo
                Sql3 = Sql3 & " and codsocio = " & DBSet(Rs!Codsocio, "N")
                Sql3 = Sql3 & " and codcampo = " & DBSet(Rs!codcampo, "N")
                Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql3 = Sql3 & " and contador = " & DBSet(Contador, "N")
                Sql3 = Sql3 & " and fechaini = " & DBSet(FechaIni, "F")
                Sql3 = Sql3 & " and fechafin = " & DBSet(FechaFin, "F")
                
                conn.Execute Sql3
            End If
'30/07/2009
'        End If
        
        Rs.MoveNext
    Wend
                                    
    Sql3 = "update tmpliquidacion set importe = round(kilosnet * precio,2) where codusu = " & vUsu.Codigo
    conn.Execute Sql3
                                    
                                    
    Me.Pb1.visible = False
    Me.Label2(10).Caption = ""
    Me.Label2(12).Caption = ""
    Me.Refresh
                                    
    CargarTemporalBodega = True
    Exit Function
    
eCargarTemporal:
    Me.Pb1.visible = False
    Me.Label2(10).Caption = ""
    Me.Label2(12).Caption = ""
    Me.Refresh
    
    MuestraError "Cargando temporal Bodega", Err.Description
End Function


Private Function HayAlbaranesSinPrecio(Tipo As Byte, cTabla As String, cWhere As String, Cooperativa As Integer) As Boolean
'Comprobar si hay precios para cada una de las variedades seleccionadas
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim B As Boolean
Dim CadAlbaranes As String
Dim cad As String

    On Error GoTo eHayAlbaranesSinPrecio
    
    HayAlbaranesSinPrecio = True
    
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    SQL = "Select numalbar FROM " & QuitarCaracterACadena(cTabla, "_1")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    If cWhere <> "" Then
        SQL = SQL & " and (prliquidalmz is null or prliquidalmz = 0)"
    Else
        SQL = SQL & " where (prliquidalmz is null or prliquidalmz = 0) "
    End If
    SQL = SQL & " order by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    B = True
    
    CadAlbaranes = ""
    ' comprobamos que existen registros para todos las variedades seleccionadas
    While Not Rs.EOF
        CadAlbaranes = CadAlbaranes & DBLet(Rs!numalbar, "N") & ", "
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    If CadAlbaranes <> "" Then
        cad = "Los siguientes albaranes no tienen precio. Revise. " & vbCrLf & vbCrLf
        cad = cad & CadAlbaranes
        
        MsgBox cad, vbExclamation
        
        B = True
    Else
        B = False
    End If
    
    HayAlbaranesSinPrecio = B
    Exit Function
    
eHayAlbaranesSinPrecio:
    MuestraError Err.nume, "Comprobando si hay precios en albaranes de Almazara", Err.Description
End Function






Private Function CargarTemporalLiquidacionAlmazaraCastelduc(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CampoAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bruto As Currency
Dim ImpoIva As Currency
Dim ImpoGastos As Currency
Dim ImpoBonif As Currency '09/09/2009: las bonificaciones las quitamos de los gastos
Dim ImpoReten As Currency
Dim ImpoAport As Currency
Dim Anticipos As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim vPorcGasto As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim cad As String
Dim HayReg As Boolean
Dim vGastos As Currency
Dim campo As String

Dim LitrosConsumidos As Long
Dim LitrosProducidos As Long

Dim PrecioConsumido As Currency
Dim PrecioProducido As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionAlmazaraCastelduc = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.numalbar, rhisfruta.prliquidalmz, "
    SQL = SQL & "rhisfruta.kilosnet as kilos "
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6"

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  bonificacion, gastos,  anticipos, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    ' en almazara no se insertan campos: metemos el minimo codcampo sin condiciones
    campo = "0" 'DevuelveValor("select min(codcampo) from rcampos")

    HayReg = False
    If Not Rs.EOF Then
        SocioAnt = Rs!Codsocio
        VarieAnt = Rs!codvarie
        NVarieAnt = Rs!nomvarie
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Rs!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.SeccionAlmaz) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                
                '[Monica] 05/07/2010 : tiene gasto de cooperativa
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    
    baseimpo = 0
    
    While Not Rs.EOF
        ' 23/07/2009: añadido el or con la segunda condicion
        If VarieAnt <> Rs!codvarie Or SocioAnt <> Rs!Codsocio Then
            
            Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos from rhisfruta_gastos, rhisfruta  "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtCodigo(6).Text, "F")
            Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtCodigo(7).Text, "F")
            Sql4 = Sql4 & " and rhisfruta.numalbar = rhisfruta_gastos.numalbar "
            
            ImpoGastos = DevuelveValor(Sql4)
            
            
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
            
            Anticipos = DevuelveValor(Sql4)
            
            Bruto = baseimpo
            
            '[Monica] 05/07/2010: el gasto de la cooperativa lo añado a la columna de gastos que no usa Valsur
            ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
            
            
            baseimpo = baseimpo - ImpoGastos - Anticipos
            
            ImpoIva = Round2((baseimpo) * ComprobarCero(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    PorcReten = 0
            End Select
        
'            ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten '- ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto, "N") & ","
            Sql1 = Sql1 & DBSet(0, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
            
            VarieAnt = Rs!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
            
            ImpoGastos = 0
            Anticipos = 0
            
        End If
        
        If Rs!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(Rs!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.SeccionAlmaz) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    
                    '[Monica] 05/07/2010 : tiene gasto de cooperativa
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(Rs!Kilos, "N")
        baseimpo = baseimpo + Round2(DBLet(Rs!Kilos, "N") * DBLet(Rs!Prliquidalmz, "N"), 2)
                
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        Bruto = baseimpo
        
        ' anticipos
        Sql4 = "select sum(rfactsoc_variedad.imporvar) "
        Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
        Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
        Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' "FAA"
        Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
        
        Anticipos = DevuelveValor(Sql4)
            
        
        ' gastos
        Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos from rhisfruta_gastos, rhisfruta "
        Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtCodigo(6).Text, "F")
        Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtCodigo(7).Text, "F")
        Sql4 = Sql4 & " and rhisfruta.numalbar = rhisfruta_gastos.numalbar "
            
        ImpoGastos = DevuelveValor(Sql4)
                
        '[Monica] 05/07/2010: el gasto de la cooperativa lo añado a la columna de gastos que no usa Valsur
        ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
                
                
                
        baseimpo = baseimpo - ImpoGastos - Anticipos
        
        ImpoIva = Round2((baseimpo) * ComprobarCero(vPorcIva) / 100, 2)
        
        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                PorcReten = 0
        End Select
    
'        ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
    
        TotalFac = baseimpo + ImpoIva - ImpoReten ' - ImpoAport
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
        Sql1 = Sql1 & DBSet(Bruto, "N") & ","
        Sql1 = Sql1 & DBSet(0, "N") & ","
        Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
        Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'        Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalLiquidacionAlmazaraCastelduc = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function



Private Function CargarTemporalAnticiposAlmazaraCastelduc(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
'29/10/2010 añado
Dim KiloGrado As Currency

Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim cad As String
Dim HayReg As Boolean

Dim Seccion As Integer
Dim PrecInd As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposAlmazaraCastelduc = False

    If vParamAplic.SeccionAlmaz = "" Then
        MsgBox "No tiene asignada en parámetros la seccion de almazara. Revise.", vbExclamation
        Exit Function
    Else
        Seccion = vParamAplic.SeccionAlmaz
    End If

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, "
    SQL = SQL & "rprecios.precioindustria,sum(rhisfruta.kilosnet) as kilos "
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5"
    SQL = SQL & " order by 1, 2, 3, 4, 5 "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto, baseimpo, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    
    If vSeccion.LeerDatos(CStr(Seccion)) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not Rs.EOF Then
        SocioAnt = Rs!Codsocio
        VarieAnt = Rs!codvarie
        NVarieAnt = Rs!nomvarie

        KilosNet = 0
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Rs!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), CStr(Seccion)) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                If vPorcIva = "" Then
                    MsgBox "El iva del socio " & DBLet(Rs!Codsocio, "N") & " no existe. Revise.", vbExclamation
                    Set vSeccion = Nothing
                    Set vSocio = Nothing
                    Set Rs = Nothing
                    Exit Function
                End If
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not Rs.EOF
        '++monica:28/07/2009 añadida la segunda condicion
        If VarieAnt <> Rs!codvarie Or SocioAnt <> Rs!Codsocio Then
            If OpcionListado = 2 Then
                  baseimpo = Round2(KilosNet * PrecInd, 2)
            End If
            
            ImpoIva = Round2(baseimpo * ImporteSinFormato(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
            
            VarieAnt = Rs!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
        End If
        
        If Rs!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(Rs!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), CStr(Seccion)) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                
                    If vPorcIva = "" Then
                        MsgBox "El iva del socio " & DBLet(Rs!Codsocio, "N") & " no existe. Revise.", vbExclamation
                        Set vSeccion = Nothing
                        Set vSocio = Nothing
                        Set Rs = Nothing
                        Exit Function
                    End If
                
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(Rs!Kilos, "N")
        
        If OpcionListado = 2 Then ' anticipo de bodega
            baseimpo = baseimpo + Round2(DBLet(Rs!Kilos, "N") * Rs!precioindustria, 2)

            PrecInd = Rs!precioindustria

        Else
            ' anticipo de almazara
            baseimpo = baseimpo + Round2(DBLet(Rs!Kilos, "N") * Rs!precioindustria, 2)
        End If
            
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        If OpcionListado = 2 Then
             baseimpo = Round2(KilosNet * PrecInd, 2)
        End If
        
        ImpoIva = Round2(baseimpo * ImporteSinFormato(vPorcIva) / 100, 2)
    
        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                PorcReten = 0
        End Select
    
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalAnticiposAlmazaraCastelduc = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function

