VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFVARListados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   9075
   Icon            =   "frmFVARListados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFacturaMaquila 
      Height          =   7380
      Left            =   0
      TabIndex        =   140
      Top             =   0
      Width           =   6855
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
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   150
         Tag             =   "F.Factura|F|S|||facturas|fecfactu|dd/mm/yyyy|S|"
         Top             =   6165
         Width           =   1230
      End
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
         Index           =   3
         ItemData        =   "frmFVARListados.frx":000C
         Left            =   1710
         List            =   "frmFVARListados.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   143
         Tag             =   "Recolecci�n|N|N|0|3|rhisfruta|recolect|||"
         Top             =   1620
         Width           =   4740
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
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   149
         Top             =   5685
         Width           =   870
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
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   163
         Text            =   "Text5"
         Top             =   5685
         Width           =   3810
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
         Index           =   27
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   148
         Tag             =   "F.Factura|F|S|||facturas|fecfactu|dd/mm/yyyy|S|"
         Top             =   4905
         Width           =   1230
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
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   146
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3435
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
         Index           =   25
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   145
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3030
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
         Index           =   2
         Left            =   5340
         TabIndex        =   152
         Top             =   6675
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepFraMaquila 
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
         Left            =   4215
         TabIndex        =   151
         Top             =   6675
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
         Index           =   22
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   154
         Text            =   "Text5"
         Top             =   4095
         Width           =   3810
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
         Index           =   22
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   147
         Top             =   4095
         Width           =   870
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
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   144
         Top             =   2115
         Width           =   870
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
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   153
         Text            =   "Text5"
         Top             =   2115
         Width           =   3810
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
         Index           =   19
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   142
         Top             =   1140
         Width           =   870
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
         Index           =   19
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   141
         Text            =   "Text5"
         Top             =   1140
         Width           =   3810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Precio Kilo"
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
         Index           =   26
         Left            =   495
         TabIndex        =   166
         Top             =   6210
         Width           =   1005
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   22
         Left            =   495
         TabIndex        =   165
         Top             =   1575
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
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
         Left            =   495
         TabIndex        =   164
         Top             =   5355
         Width           =   1155
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1440
         MouseIcon       =   "frmFVARListados.frx":0010
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar forma pago"
         Top             =   5685
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   9
         Left            =   1440
         Picture         =   "frmFVARListados.frx":0162
         ToolTipText     =   "Buscar fecha"
         Top             =   4905
         Width           =   240
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
         Index           =   24
         Left            =   495
         TabIndex        =   162
         Top             =   2115
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Generaci�n Factura de Maquila"
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
         Index           =   4
         Left            =   495
         TabIndex        =   161
         Top             =   315
         Width           =   5820
      End
      Begin VB.Label Label4 
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
         Index           =   39
         Left            =   495
         TabIndex        =   160
         Top             =   4635
         Width           =   1440
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Albaran"
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
         Index           =   38
         Left            =   495
         TabIndex        =   159
         Top             =   2685
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
         Index           =   37
         Left            =   795
         TabIndex        =   158
         Top             =   3030
         Width           =   645
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
         Index           =   36
         Left            =   795
         TabIndex        =   157
         Top             =   3435
         Width           =   600
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1440
         Picture         =   "frmFVARListados.frx":01ED
         ToolTipText     =   "Buscar fecha"
         Top             =   3465
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1440
         Picture         =   "frmFVARListados.frx":0278
         ToolTipText     =   "Buscar fecha"
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1440
         MouseIcon       =   "frmFVARListados.frx":0303
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   4095
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   29
         Left            =   495
         TabIndex        =   156
         Top             =   4095
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1440
         MouseIcon       =   "frmFVARListados.frx":0455
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   2115
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Secci�n"
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
         Left            =   495
         TabIndex        =   155
         Top             =   1080
         Width           =   780
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1440
         MouseIcon       =   "frmFVARListados.frx":05A7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar secci�n"
         Top             =   1140
         Width           =   240
      End
   End
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
   Begin VB.Frame FrameReimpresion 
      Height          =   7110
      Left            =   0
      TabIndex        =   13
      Top             =   -60
      Width           =   6855
      Begin VB.Frame FrameTipoFactura 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   3105
         TabIndex        =   128
         Top             =   1860
         Width           =   3405
         Begin MSComctlLib.ListView ListView1 
            Height          =   1110
            Index           =   0
            Left            =   180
            TabIndex        =   129
            Top             =   360
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   1958
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   2820
            Picture         =   "frmFVARListados.frx":06F9
            ToolTipText     =   "Marcar todos"
            Top             =   90
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   3060
            Picture         =   "frmFVARListados.frx":6F4B
            ToolTipText     =   "Desmarcar todos"
            Top             =   90
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
            Left            =   210
            TabIndex        =   130
            Top             =   75
            Width           =   1815
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
         Index           =   64
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text5"
         Top             =   1140
         Width           =   3810
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
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1140
         Width           =   870
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
         Index           =   65
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Text5"
         Top             =   1500
         Width           =   3810
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
         Index           =   65
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1500
         Width           =   870
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
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   10
         Top             =   5985
         Width           =   870
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
         Index           =   63
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text5"
         Top             =   5985
         Width           =   3810
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
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   9
         Top             =   5580
         Width           =   870
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
         Index           =   62
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Text5"
         Top             =   5580
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
         Index           =   1
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   4815
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
         Index           =   0
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   4395
         Width           =   3810
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
         Index           =   1
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   8
         Top             =   4815
         Width           =   870
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
         Index           =   0
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   7
         Top             =   4395
         Width           =   870
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
         Left            =   4215
         TabIndex        =   11
         Top             =   6405
         Width           =   1065
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
         Left            =   5340
         TabIndex        =   12
         Top             =   6405
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
         Index           =   2
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3345
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
         Index           =   3
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3750
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
         Index           =   4
         Left            =   1725
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "N� Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2235
         Width           =   1230
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
         Left            =   1725
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "N� Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2625
         Width           =   1230
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   64
         Left            =   1470
         MouseIcon       =   "frmFVARListados.frx":794D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar secci�n"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Secci�n"
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
         Index           =   55
         Left            =   480
         TabIndex        =   60
         Top             =   810
         Width           =   780
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
         Index           =   54
         Left            =   750
         TabIndex        =   59
         Top             =   1515
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
         Index           =   53
         Left            =   750
         TabIndex        =   58
         Top             =   1140
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   65
         Left            =   1470
         MouseIcon       =   "frmFVARListados.frx":7A9F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar secci�n"
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   63
         Left            =   1440
         MouseIcon       =   "frmFVARListados.frx":7BF1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   5985
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
         Index           =   52
         Left            =   750
         TabIndex        =   54
         Top             =   5580
         Width           =   645
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
         Index           =   51
         Left            =   750
         TabIndex        =   53
         Top             =   6000
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   480
         TabIndex        =   52
         Top             =   5250
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   62
         Left            =   1440
         MouseIcon       =   "frmFVARListados.frx":7D43
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   5580
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1440
         MouseIcon       =   "frmFVARListados.frx":7E95
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   4815
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         MouseIcon       =   "frmFVARListados.frx":7FE7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   4395
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
         Index           =   11
         Left            =   480
         TabIndex        =   25
         Top             =   4110
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
         Index           =   12
         Left            =   750
         TabIndex        =   24
         Top             =   4815
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
         Index           =   13
         Left            =   750
         TabIndex        =   23
         Top             =   4395
         Width           =   645
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1485
         Picture         =   "frmFVARListados.frx":8139
         ToolTipText     =   "Buscar fecha"
         Top             =   3750
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1485
         Picture         =   "frmFVARListados.frx":81C4
         ToolTipText     =   "Buscar fecha"
         Top             =   3345
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
         Index           =   14
         Left            =   750
         TabIndex        =   22
         Top             =   3750
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
         Index           =   15
         Left            =   750
         TabIndex        =   21
         Top             =   3345
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
         Index           =   16
         Left            =   480
         TabIndex        =   20
         Top             =   3000
         Width           =   1815
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
         Left            =   480
         TabIndex        =   19
         Top             =   1905
         Width           =   1170
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
         Index           =   0
         Left            =   750
         TabIndex        =   18
         Top             =   2265
         Width           =   645
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
         Index           =   1
         Left            =   750
         TabIndex        =   17
         Top             =   2625
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Reimpresi�n de Facturas Varias"
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
         Left            =   495
         TabIndex        =   16
         Top             =   315
         Width           =   5820
      End
   End
   Begin VB.Frame FrameIntConta 
      Height          =   6780
      Left            =   30
      TabIndex        =   92
      Top             =   45
      Width           =   7680
      Begin VB.CommandButton cmdAceptar 
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
         Left            =   5175
         TabIndex        =   106
         Top             =   6150
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
         Left            =   6360
         TabIndex        =   107
         Top             =   6150
         Width           =   1065
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos para la contabilizaci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2400
         Left            =   90
         TabIndex        =   116
         Top             =   2670
         Width           =   7425
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
            Index           =   18
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   105
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1920
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
            Index           =   18
            Left            =   3825
            Locked          =   -1  'True
            TabIndex        =   119
            Top             =   1920
            Width           =   3450
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   102
            Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
            Top             =   705
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
            Index           =   16
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   1110
            Width           =   3630
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   103
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|000||"
            Top             =   1110
            Width           =   1215
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   101
            Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
            Top             =   330
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
            Index           =   14
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   104
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|000||"
            Top             =   1515
            Width           =   1215
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
            Index           =   14
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   117
            Top             =   1515
            Width           =   3630
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2160
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1965
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
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
            Index           =   24
            Left            =   180
            TabIndex        =   124
            Top             =   1965
            Width           =   1935
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   2160
            Picture         =   "frmFVARListados.frx":824F
            ToolTipText     =   "Buscar fecha"
            Top             =   750
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
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
            Index           =   19
            Left            =   180
            TabIndex        =   123
            Top             =   750
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Positivas"
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
            Height          =   285
            Index           =   5
            Left            =   180
            TabIndex        =   122
            Top             =   1155
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   2160
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1155
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   2160
            Picture         =   "frmFVARListados.frx":82DA
            ToolTipText     =   "Buscar fecha"
            Top             =   375
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Recepci�n"
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
            Index           =   18
            Left            =   180
            TabIndex        =   121
            Top             =   330
            Width           =   1755
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   2160
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Negativas"
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
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   120
            Top             =   1560
            Width           =   1800
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selecci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2370
         Left            =   90
         TabIndex        =   93
         Top             =   225
         Width           =   7410
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
            Left            =   1590
            MaxLength       =   10
            TabIndex        =   98
            Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
            Top             =   1500
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
            Index           =   10
            Left            =   4125
            MaxLength       =   10
            TabIndex        =   99
            Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
            Top             =   1500
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
            Index           =   7
            Left            =   1590
            MaxLength       =   7
            TabIndex        =   96
            Tag             =   "N� de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
            Top             =   900
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
            Index           =   8
            Left            =   4125
            MaxLength       =   7
            TabIndex        =   97
            Tag             =   "N� de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
            Top             =   900
            Width           =   1350
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
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   1905
            Width           =   3330
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
            Left            =   1590
            MaxLength       =   10
            TabIndex        =   95
            Tag             =   "Seccion|N|S|||sparam|codsecci|000||"
            Top             =   360
            Width           =   825
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
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   94
            Top             =   360
            Width           =   4725
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
            Index           =   17
            Left            =   165
            TabIndex        =   115
            Top             =   1245
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
            Index           =   10
            Left            =   555
            TabIndex        =   114
            Top             =   1485
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
            Index           =   9
            Left            =   3165
            TabIndex        =   113
            Top             =   1515
            Width           =   645
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   1305
            Picture         =   "frmFVARListados.frx":8365
            ToolTipText     =   "Buscar fecha"
            Top             =   1485
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   6
            Left            =   3840
            Picture         =   "frmFVARListados.frx":83F0
            ToolTipText     =   "Buscar fecha"
            Top             =   1515
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
            Index           =   8
            Left            =   555
            TabIndex        =   112
            Top             =   900
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
            Index           =   7
            Left            =   3180
            TabIndex        =   111
            Top             =   945
            Width           =   645
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Factura"
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
            Left            =   165
            TabIndex        =   110
            Top             =   660
            Width           =   765
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
            Index           =   3
            Left            =   150
            TabIndex        =   109
            Top             =   1935
            Width           =   1830
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1290
            ToolTipText     =   "Buscar secci�n"
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Secci�n"
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
            Index           =   1
            Left            =   150
            TabIndex        =   108
            Top             =   360
            Width           =   915
         End
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   270
         Left            =   90
         TabIndex        =   125
         Top             =   5250
         Visible         =   0   'False
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   127
         Top             =   5880
         Width           =   7410
      End
      Begin VB.Label lblProgres 
         Height          =   240
         Index           =   2
         Left            =   60
         TabIndex        =   126
         Top             =   5520
         Width           =   7425
      End
   End
   Begin VB.Frame FrameCargaMasivaFras 
      Height          =   9060
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   8890
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
         Index           =   2
         ItemData        =   "frmFVARListados.frx":847B
         Left            =   3735
         List            =   "frmFVARListados.frx":847D
         Style           =   2  'Dropdown List
         TabIndex        =   138
         Tag             =   "Recolecci�n|N|N|0|3|rhisfruta|recolect|||"
         Top             =   6300
         Width           =   2175
      End
      Begin VB.Frame FrameConta 
         Caption         =   "Datos para la contabilizaci�n"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   225
         TabIndex        =   131
         Top             =   6690
         Visible         =   0   'False
         Width           =   8235
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   4545
            TabIndex        =   135
            Top             =   180
            Width           =   3600
            Begin VB.OptionButton Option1 
               Caption         =   "Cobros"
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
               Left            =   1230
               TabIndex        =   137
               Top             =   105
               Width           =   1095
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Pagos"
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
               Left            =   2430
               TabIndex        =   136
               Top             =   105
               Width           =   1095
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
            Index           =   13
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   47
            Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
            Top             =   360
            Width           =   1400
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
            Left            =   3825
            Locked          =   -1  'True
            TabIndex        =   132
            Top             =   750
            Width           =   4260
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   48
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   750
            Width           =   1400
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
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
            Index           =   20
            Left            =   180
            TabIndex        =   134
            Top             =   360
            Width           =   1965
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   2160
            Picture         =   "frmFVARListados.frx":847F
            ToolTipText     =   "Buscar fecha"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
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
            Index           =   3
            Left            =   180
            TabIndex        =   133
            Top             =   795
            Width           =   1935
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   2160
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   750
            Width           =   240
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Insertar en Tesoreria"
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
         Height          =   240
         Index           =   1
         Left            =   6015
         TabIndex        =   46
         Top             =   6420
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Personalizar Importes"
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
         Left            =   6015
         TabIndex        =   45
         Top             =   6150
         Width           =   2490
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   420
         TabIndex        =   84
         Top             =   720
         Width           =   8085
         Begin VB.OptionButton Option1 
            Caption         =   "Clientes"
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
            Left            =   4470
            TabIndex        =   86
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Socios"
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
            Left            =   1275
            TabIndex        =   85
            Top             =   270
            Width           =   2535
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
         Index           =   46
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   2340
         Width           =   5565
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
         Index           =   46
         Left            =   2025
         MaxLength       =   6
         TabIndex        =   36
         Tag             =   "Cta.Contable|T|S|||sparam|codforpa|000||"
         Top             =   2340
         Width           =   915
      End
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
         Index           =   0
         ItemData        =   "frmFVARListados.frx":850A
         Left            =   4800
         List            =   "frmFVARListados.frx":850C
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Tag             =   "Recolecci�n|N|N|0|3|rhisfruta|recolect|||"
         Top             =   1860
         Width           =   3750
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
         Left            =   2025
         MaxLength       =   10
         TabIndex        =   34
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   1920
         Width           =   1275
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
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   40
         Tag             =   "C�digo Propio|N|N|1|99|sbanco|codbanpr|000|S|"
         Top             =   5310
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
         Index           =   71
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text5"
         Top             =   5310
         Width           =   5715
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
         Left            =   390
         MaxLength       =   10
         TabIndex        =   42
         Tag             =   "C�digo Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   6300
         Width           =   870
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
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   43
         Tag             =   "C�digo Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   6300
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
         Index           =   68
         Left            =   2385
         MaxLength       =   10
         TabIndex        =   44
         Tag             =   "C�digo Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   6300
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
         Height          =   795
         Index           =   67
         Left            =   360
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Tag             =   "Observaciones|T|S|||cabfact|observac|||"
         Top             =   4305
         Width           =   8130
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
         Index           =   66
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   41
         Tag             =   "C�digo Propio|N|N|1|99|sbanco|codbanpr|000|S|"
         Top             =   5700
         Width           =   6795
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
         Left            =   2025
         MaxLength       =   6
         TabIndex        =   33
         Top             =   1470
         Width           =   915
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
         Index           =   52
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   1470
         Width           =   5595
      End
      Begin VB.CommandButton CmdAcepRecalImp 
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
         Left            =   6225
         TabIndex        =   49
         Top             =   8490
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
         Left            =   7395
         TabIndex        =   50
         Top             =   8475
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   225
         Left            =   225
         TabIndex        =   89
         Top             =   7980
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame FrameSocio 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1215
         Left            =   300
         TabIndex        =   0
         Top             =   2760
         Width           =   8400
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
            Left            =   1710
            MaxLength       =   6
            TabIndex        =   37
            Top             =   390
            Width           =   870
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
            Left            =   1710
            MaxLength       =   6
            TabIndex        =   38
            Top             =   780
            Width           =   870
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
            Index           =   73
            Left            =   2610
            Locked          =   -1  'True
            TabIndex        =   73
            Text            =   "Text5"
            Top             =   390
            Width           =   5595
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
            Left            =   2610
            Locked          =   -1  'True
            TabIndex        =   72
            Text            =   "Text5"
            Top             =   780
            Width           =   5595
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   73
            Left            =   1380
            MouseIcon       =   "frmFVARListados.frx":850E
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   390
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
            Index           =   67
            Left            =   585
            TabIndex        =   76
            Top             =   390
            Width           =   645
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
            Index           =   66
            Left            =   600
            TabIndex        =   75
            Top             =   765
            Width           =   600
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
            Index           =   65
            Left            =   90
            TabIndex        =   74
            Top             =   120
            Width           =   540
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   74
            Left            =   1380
            MouseIcon       =   "frmFVARListados.frx":8660
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   780
            Width           =   240
         End
      End
      Begin VB.Frame FrameClientes 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1335
         Left            =   315
         TabIndex        =   78
         Top             =   2700
         Width           =   8490
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
            Index           =   47
            Left            =   2610
            Locked          =   -1  'True
            TabIndex        =   80
            Text            =   "Text5"
            Top             =   480
            Width           =   5640
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
            Index           =   47
            Left            =   1710
            MaxLength       =   6
            TabIndex        =   87
            Top             =   480
            Width           =   870
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
            Index           =   48
            Left            =   2610
            Locked          =   -1  'True
            TabIndex        =   79
            Text            =   "Text5"
            Top             =   870
            Width           =   5640
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
            Index           =   48
            Left            =   1710
            MaxLength       =   6
            TabIndex        =   88
            Top             =   870
            Width           =   870
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   47
            Left            =   1380
            MouseIcon       =   "frmFVARListados.frx":87B2
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            Left            =   60
            TabIndex        =   83
            Top             =   210
            Width           =   675
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
            Index           =   63
            Left            =   555
            TabIndex        =   82
            Top             =   855
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
            Index           =   58
            Left            =   540
            TabIndex        =   81
            Top             =   480
            Width           =   690
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   48
            Left            =   1380
            MouseIcon       =   "frmFVARListados.frx":8904
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   870
            Width           =   240
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Descuenta en "
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
         Index           =   21
         Left            =   3735
         TabIndex        =   139
         Top             =   6075
         Width           =   1455
      End
      Begin VB.Label lblProgres 
         Caption         =   "12111111"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   91
         Top             =   8220
         Width           =   8205
      End
      Begin VB.Label lblProgres 
         Height          =   225
         Index           =   1
         Left            =   420
         TabIndex        =   90
         Top             =   7140
         Width           =   3585
      End
      Begin VB.Label Label13 
         Caption         =   "Carga Masiva de Facturas Varias"
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
         TabIndex        =   77
         Top             =   270
         Width           =   6120
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1680
         MouseIcon       =   "frmFVARListados.frx":8A56
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Forma Pago"
         Top             =   2370
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Forma  Pago"
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
         Left            =   390
         TabIndex        =   71
         Top             =   2370
         Width           =   1230
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   18
         Left            =   1680
         Picture         =   "frmFVARListados.frx":8BA8
         ToolTipText     =   "Buscar fecha"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fec.Factura"
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
         Index           =   59
         Left            =   390
         TabIndex        =   69
         Top             =   1920
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
         Left            =   3465
         TabIndex        =   68
         Top             =   1890
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
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
         Left            =   390
         TabIndex        =   67
         Top             =   6060
         Width           =   870
      End
      Begin VB.Label Label4 
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
         Index           =   61
         Left            =   1335
         TabIndex        =   66
         Top             =   6060
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Concepto "
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
         Left            =   390
         TabIndex        =   65
         Top             =   5310
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1395
         MouseIcon       =   "frmFVARListados.frx":8C33
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   5310
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   57
         Left            =   2415
         TabIndex        =   64
         Top             =   6060
         Width           =   765
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
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
         Left            =   405
         TabIndex        =   63
         Top             =   4050
         Width           =   1665
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ampliaci�n"
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
         Left            =   390
         TabIndex        =   62
         Top             =   5700
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Secci�n"
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
         Index           =   31
         Left            =   390
         TabIndex        =   31
         Top             =   1470
         Width           =   780
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   52
         Left            =   1680
         MouseIcon       =   "frmFVARListados.frx":8D85
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar secci�n"
         Top             =   1500
         Width           =   240
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Factura"
      ForeColor       =   &H00972E0B&
      Height          =   255
      Index           =   30
      Left            =   0
      TabIndex        =   29
      Top             =   -30
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Desde"
      Height          =   195
      Index           =   31
      Left            =   375
      TabIndex        =   28
      Top             =   300
      Width           =   465
   End
   Begin VB.Label Label4 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   32
      Left            =   375
      TabIndex        =   27
      Top             =   675
      Width           =   420
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   11
      Left            =   1020
      Picture         =   "frmFVARListados.frx":8ED7
      ToolTipText     =   "Buscar fecha"
      Top             =   255
      Width           =   240
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   12
      Left            =   1020
      Picture         =   "frmFVARListados.frx":8F62
      ToolTipText     =   "Buscar fecha"
      Top             =   645
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "Ejercicio"
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   27
      Left            =   0
      TabIndex        =   26
      Top             =   1065
      Width           =   705
   End
End
Attribute VB_Name = "frmFVARListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public Opcionlistado As Byte
    '==== Listados / Procesos FACTURAS VARIAS ====
    '=============================
    ' 1 .- Reimpresion de Facturas
    ' 2 .- Grabacion de Facturas Varias (dentro del mantenimiento de facturas varias)
    ' 3 .- Diario de Facturaci�n
    
    ' 4 .- Integracion contable facturas varias en registro de iva de cliente
    
    ' 5 .- Reimpresion de Facturas Proveedor
    ' 6 .- Diario de Facturacion de Proveedor
    
    ' 7 .- Integracion contable facturas varias en el registro de iva de proveedor
    
    ' 8 .- factura de maquila
    
    
Public AnticipoGastos As Boolean ' si true entonces es que se trata de anticipos de gastos de recoleccion
Public LiquidacionIndustria As Boolean ' si true entonces es que se trata de liquidacion de industria
Public AnticipoGenerico As Boolean ' si true entonces es que se trata de anticipos genericos,
    ' todos los kilos independientemente de que esten o no clasificados se anticipan a un mismo precio
    

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmSec As frmManSeccion 'Secciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTranspor 'Transportistas
Attribute frmTra.VB_VarHelpID = -1
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
Private WithEvents frmCli As frmBasico2 'Ayuda de clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMens1 As frmMensajes 'Mensajes
Attribute frmMens1.VB_VarHelpID = -1
Private WithEvents frmMens2 As frmMensajes 'Mensajes
Attribute frmMens2.VB_VarHelpID = -1
Private WithEvents frmMens3 As frmMensajes 'Mensajes
Attribute frmMens3.VB_VarHelpID = -1
Private WithEvents frmMens4 As frmMensajes 'Mensajes
Attribute frmMens4.VB_VarHelpID = -1
Private WithEvents frmCon As frmFVARConceptos  ' conceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmFPa As frmComFpa  ' formas de pago de comercial
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas de contabilidad
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmMaq As frmFVARMaquilaAux
Attribute frmMaq.VB_VarHelpID = -1

Private WithEvents frmFactV As frmFVARFactPerso 'personalizacion de las facturas generadas masivamente
Attribute frmFactV.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean

Private cadFormula2 As String
Private cadSelect2 As String
Private cadSelect1 As String

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As Byte

Dim Indice As Integer

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte

Dim Bodega As Boolean
Dim Industria As Boolean

Dim Variedades As String
Dim albaranes As String

Dim vReturn As Integer
Dim vSeccion As CSeccion

Dim cContaFra As cContabilizarFacturas

Dim cadTabla As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub Check1_Click(Index As Integer)
    If Index = 1 Then
        FrameConta.visible = (Check1(1).Value = 1)
        FrameConta.Enabled = (Check1(1).Value = 1)
        
        If FrameConta.Enabled Then Option1(3).Value = True
        
    Else
        Check1(1).Enabled = (Check1(0).Value = 1)
        Check1(1).visible = (Check1(0).Value = 1)
        If Not Check1(1).Enabled Then Check1(1).Value = 0
        
        FrameConta.visible = (Check1(1).Value = 1)
        FrameConta.Enabled = (Check1(1).Value = 1)
    
        If FrameConta.Enabled Then Option1(3).Value = True
    
    End If
End Sub

Private Sub CmdAcepFraMaquila_Click()
Dim SQL As String
Dim i As Byte
Dim cadWhere As String
Dim cDesde As String
Dim cHasta As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
            
    'D/H Fecha albaran
    cDesde = Trim(txtcodigo(25).Text)
    cHasta = Trim(txtcodigo(26).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
            
    If CargarTablaIntermedia(cadselect) Then
        If TotalRegistros("select * from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")) <> 0 Then
        
            Set frmMaq = New frmFVARMaquilaAux
            frmMaq.Show vbModal
            Set frmMaq = Nothing
        
        
            If GenerarFacturaMaquila Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
            
                If Not vSeccion Is Nothing Then
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                End If
                cmdCancel_Click (0)
                        
            End If
        End If
    Else
        MsgBox "No se ha realizado el proceso." & MensError, vbExclamation
    
        If Not vSeccion Is Nothing Then
            vSeccion.CerrarConta
            Set vSeccion = Nothing
        End If
        cmdCancel_Click (0)
    End If
        
End Sub

Private Function GenerarFacturaMaquila()
Dim SQL As String
Dim CodTipoMov As String
Dim vTipoMov As CTiposMov
Dim NumFact As String
Dim Existe As Boolean
Dim TipoIVA As String
Dim PorIva As String
Dim PorRec As String
Dim ImpoIva As Currency
Dim ImpoRec As Currency
Dim TotalFact As Currency
Dim CabSql As String
Dim LinSql As String
Dim LinSqlInsert As String
Dim NumLin As Long
Dim ImporteTot As Currency
Dim Importe As Currency
Dim Rs As ADODB.Recordset


    GenerarFacturaMaquila = False


    ' primero creamos cabecera
    SQL = "GENFAC" 'generar facturas de venta

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Generar Facturas. Hay otro usuario realizando el proceso.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    CodTipoMov = ""
    If Len(Combo1(3).Text) >= 3 Then CodTipoMov = Mid(Combo1(3).Text, 1, 3)

    conn.BeginTrans
        
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        NumFact = vTipoMov.ConseguirContador(CodTipoMov)
    
        Existe = False
        Do
            SQL = "select count(*) from fvarcabfact where "
            SQL = SQL & " codtipom = " & DBSet(CodTipoMov, "T")
            SQL = SQL & " and numfactu = " & DBSet(NumFact, "N")
            SQL = SQL & " and fecfactu = " & DBSet(txtcodigo(27).Text, "F")
            If TotalRegistros(SQL) > 0 Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (CodTipoMov)
                NumFact = vTipoMov.ConseguirContador(CodTipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
        
        TipoIVA = ""
        PorIva = ""
        ImpoIva = 0
        TotalFact = 0
        
        TipoIVA = DevuelveDesdeBDNew(cAgro, "fvarconce", "tipoiva", "codconce", txtcodigo(20).Text, "N")
        If CodTipoMov = "FVG" Then
            TipoIVA = vSeccion.TipIvaExento
        End If
        PorIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", TipoIVA, "N")
        PorRec = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", TipoIVA, "N")
        ' lo calculo despues
        ImpoIva = 0 'Round2(DBLet(Rs!Importe, "N") * ComprobarCero(PorIva) / 100, 2)
        ImpoRec = 0 'Round2(DBLet(Rs!Importe, "N") * ComprobarCero(PorRec) / 100, 2)
        
        TotalFact = 0 'DBLet(Rs!Importe, "N") + ImpoIva + ImpoREC
        
        ' Insertamos en la cabecera de factura
        CabSql = "insert into fvarcabfact ("
        CabSql = CabSql & "codsecci,codtipom,numfactu,fecfactu,codsocio,codclien,observac,intconta,baseiva1,baseiva2,baseiva3,"
        CabSql = CabSql & "impoiva1,impoiva2,impoiva3,imporec1,imporec2,imporec3,totalfac,tipoiva1,tipoiva2,tipoiva3,"
        CabSql = CabSql & "porciva1 , porciva2, porciva3, codforpa, porcrec1, porcrec2, porcrec3, retfaccl, trefaccl, cuereten, enliquidacion)  values  "
        
        CabSql = CabSql & "(" & DBSet(txtcodigo(19).Text, "N")
        CabSql = CabSql & "," & DBSet(CodTipoMov, "T")
        CabSql = CabSql & "," & DBSet(NumFact, "N")
        CabSql = CabSql & "," & DBSet(txtcodigo(27).Text, "F")
        CabSql = CabSql & "," & ValorNulo & "," & DBSet(txtcodigo(22), "N") ' cliente
            
        CabSql = CabSql & "," & ValorNulo 'DBSet(txtcodigo(67).Text, "T", "S")
        CabSql = CabSql & ",0"
        CabSql = CabSql & "," & DBSet(0, "N")
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & DBSet(ImpoIva, "N")
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & DBSet(ImpoRec, "N", "S")
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & ValorNulo
        
        CabSql = CabSql & "," & DBSet(TotalFact, "N")
        CabSql = CabSql & "," & DBSet(TipoIVA, "N")
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & DBSet(PorIva, "N")
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & DBSet(txtcodigo(21).Text, "N") ' forma de pago
        CabSql = CabSql & "," & DBSet(PorRec, "N", "S")
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & ValorNulo
        CabSql = CabSql & "," & ValorNulo
        
        '[Monica]29/05/2017: a�adimos donde descuenta
        CabSql = CabSql & ",0" '& DBSet(Rs!enliquidacion, "N")
        
        CabSql = CabSql & ")"
        
        conn.Execute CabSql
        
        
        ' insertamos en la linea de factura
        LinSqlInsert = "insert into fvarlinfact (codtipom, numfactu, fecfactu, NumLinea, codConce, ampliaci, precio, cantidad, Importe, TipoIva) values "
        LinSql = ""
        
        SQL = "select codigo1, importe1, variedades.nomvarie from tmpinformes inner join variedades  on tmpinformes.codigo1 = variedades.codvarie where codusu = " & vUsu.Codigo & " order by codigo1"
        
        NumLin = 0
        ImporteTot = 0
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
        
            NumLin = NumLin + 1
            Importe = Round2(DBLet(Rs!importe1, "N") * ImporteSinFormato(txtcodigo(23)), 2)
            ImporteTot = ImporteTot + Importe
        
            LinSql = LinSql & ",("
            LinSql = LinSql & DBSet(CodTipoMov, "T")
            LinSql = LinSql & "," & DBSet(NumFact, "N")
            LinSql = LinSql & "," & DBSet(txtcodigo(27).Text, "F")
            LinSql = LinSql & "," & DBSet(NumLin, "N")
            LinSql = LinSql & "," & DBSet(txtcodigo(20), "N")
            LinSql = LinSql & "," & DBSet(Rs!nomvarie, "T")
            LinSql = LinSql & "," & DBSet(txtcodigo(23), "N")
            LinSql = LinSql & "," & DBSet(Rs!importe1, "N")
            LinSql = LinSql & "," & DBSet(Importe, "N")
            LinSql = LinSql & "," & DBSet(TipoIVA, "N")
            LinSql = LinSql & ")"
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
        If LinSql <> "" Then
            conn.Execute LinSqlInsert & Mid(LinSql, 2)
            
            ImpoIva = Round2(ImporteTot * ComprobarCero(PorIva) / 100, 2)
            ImpoRec = 0 'Round2(DBLet(Rs!Importe, "N") * ComprobarCero(PorRec) / 100, 2)
            
            TotalFact = ImporteTot + ImpoIva + ImpoRec
            
            ' actualizamos los totales de la factura con las lines
            SQL = "update fvarcabfact set baseiva1 = " & DBSet(ImporteTot, "N")
            SQL = SQL & ", impoiva1 = " & DBSet(ImpoIva, "N")
            SQL = SQL & ", totalfac = " & DBSet(TotalFact, "N")
            SQL = SQL & " where codsecci = " & DBSet(txtcodigo(19).Text, "N")
            SQL = SQL & " and codtipom = " & DBSet(CodTipoMov, "T")
            SQL = SQL & " and numfactu = " & DBSet(NumFact, "N")
            SQL = SQL & " and fecfactu = " & DBSet(txtcodigo(27).Text, "F")
            
            conn.Execute SQL
        End If
        
        vTipoMov.IncrementarContador (CodTipoMov)
        Set vTipoMov = Nothing
    End If
    
EGenerarFacturaMaquila:
    If Err.Number <> 0 Then
        MensError = "Generar Factura de maquila " '& Err.Description
        conn.RollbackTrans
    Else
        GenerarFacturaMaquila = True
        conn.CommitTrans
    End If
End Function


Private Function CargarTablaIntermedia(vWhere As String) As Boolean
Dim SQL As String
Dim SQLinsert As String

    On Error GoTo eCargarTablaIntermedia

    CargarTablaIntermedia = False

    SQL = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
            
    SQLinsert = "insert into tmpinformes (codusu, codigo1, importe1) "

    SQL = "select " & vUsu.Codigo & ", codvarie, sum(albaran_variedad.pesoneto) from albaran inner join albaran_variedad on albaran.numalbar = albaran_variedad.numalbar "
    SQL = SQL & " where (1=1) "
    If vWhere <> "" Then SQL = SQL & " and " & Replace(Replace(vWhere, "{", ""), "}", "")
'    SQL = SQL & " and (albaran_variedad.numalbar, albaran_variedad.numlinea) in (select numalbar, numlinealbar from facturas_variedad) "
    
    SQL = SQL & " group by 1, 2"
    SQL = SQL & " order by 1, 2"
    
    conn.Execute SQLinsert & SQL
    CargarTablaIntermedia = True
    Exit Function

eCargarTablaIntermedia:
    MuestraError Err.Number, "Cargando Tabla Intermedia", Err.Description
End Function


Private Sub CmdAcepRecalImp_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cOrden As String
Dim cadMen As String
Dim i As Byte
Dim SQL As String
Dim Tipo As Byte
Dim Nregs As Long
Dim NumError As Long
Dim cWhere As String
    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
            
            
    If Check1(0).Value = 0 Then
        '[Monica]21/09/2018: se borra la tabla intermedia
        SQL = "delete from tmpfactvarias where codusu = " & DBSet(vUsu.Codigo, "N")
        conn.Execute SQL
            
        If Option1(0).Value Then
            cadTabla = "rsocios"
        
            cWhere = " rsocios.codsocio >= " & Trim(txtcodigo(73).Text) & " and rsocios.codsocio <= " & Trim(txtcodigo(74).Text)
            cWhere = cWhere & " and rsocios.codsocio in (select codsocio from rsocios_seccion where codsecci = " & DBSet(txtcodigo(52).Text, "N") & " and fecbaja is null) "
        
            Set frmMens = New frmMensajes
        
            frmMens.OpcionMensaje = 9
            frmMens.Label5 = "Socios"
            frmMens.cadWhere = cWhere
            frmMens.Show vbModal
        
            Set frmMens = Nothing
        Else
            cadTabla = "clientes"
            
            cWhere = " clientes.codclien >= " & Trim(txtcodigo(47).Text) & " and clientes.codclien <= " & Trim(txtcodigo(48).Text)
        
            Set frmMens = New frmMensajes
        
            frmMens.OpcionMensaje = 8
            frmMens.Label5 = "Clientes"
            frmMens.cadWhere = cWhere
            frmMens.Show vbModal
        
            Set frmMens = Nothing
        End If
    
    
        If TotalRegistros("select count(*) from " & cadTabla & " where " & cadselect) <> 0 Then
            If GenerarFacturasPrevio(cadTabla, cadselect, NumError, MensError) Then
                GenerarFacturas cadTabla, cadselect, NumError, MensError
            End If
        Else
            MsgBox "No se ha realizado el proceso." & MensError, vbExclamation
        
            Exit Sub
        End If
    Else
        If Check1(1).Value Then
            If txtcodigo(12).Text = "" Then
                MsgBox "Debe introducir la Cta de Banco Prevista. Revise", vbExclamation
                PonerFoco txtcodigo(12)
                Exit Sub
            End If
            
            If txtcodigo(13).Text = "" Then
                MsgBox "Debe introducir la fecha de Vencimiento. Revise", vbExclamation
                PonerFoco txtcodigo(13)
                Exit Sub
            End If
        End If
    
        Set frmFactV = New frmFVARFactPerso
    
        If Option1(0).Value Then
            cadTabla = "rsocios"
        Else
            cadTabla = "clientes"
        End If
    
        frmFactV.ParamSeccion = txtcodigo(52).Text
        frmFactV.ParamTabla = cadTabla
        frmFactV.ParamAmpliacion = txtcodigo(66).Text
        frmFactV.ParamConcepto = txtcodigo(71).Text
        frmFactV.ParamNomConcep = txtNombre(71).Text
        frmFactV.ParamCantidad = txtcodigo(70).Text
        frmFactV.ParamPrecio = txtcodigo(69).Text
        frmFactV.ParamImporte = txtcodigo(68).Text
        frmFactV.ParamDescuenta = Combo1(2).ListIndex
        
        
        frmFactV.Show vbModal
        
        Set frmFactV = Nothing
        
        
        If TotalRegistros("select count(*) from tmpfactvarias where codusu = " & vUsu.Codigo) <> 0 Then
        
            If MsgBox("� Desea continuar con el proceso ?", vbExclamation + vbYesNo) = vbYes Then
                If Check1(1).Value = 0 Then
                    GenerarFacturas cadTabla, cadselect, NumError, MensError
                Else
                    Dim B As Boolean
                    
                    B = True
                    If txtcodigo(13).Text = "" Then
                        MsgBox "Debe introducir la fecha de vencimiento. Revise.", vbExclamation
                        PonerFoco txtcodigo(13)
                        B = False
                    End If
                    If Not B Then Exit Sub
                    
                    If txtcodigo(12).Text = "" Then
                        MsgBox "Debe introducir la cuenta de banco. Revise.", vbExclamation
                        PonerFoco txtcodigo(12)
                        B = False
                    End If
                    If Not B Then Exit Sub
                    
                    ContabilizarCobros NumError
                
                End If
            Else
                MsgBox "No se ha realizado el proceso." & MensError, vbExclamation
            
                Exit Sub
            End If
        Else
            MsgBox "No se ha realizado el proceso." & MensError, vbExclamation
        
            Exit Sub
        End If
    End If
    
    
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("GENFAC") 'VENtas CONtabilizar
        
eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de generaci�n." & vbCrLf & MensError
    Else
        MsgBox "Proceso realizado correctamente.", vbExclamation
    End If
    
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    
    cmdCancel_Click (0)

End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
Dim i As Byte
Dim cadWhere As String
Dim cDesde As String
Dim cHasta As String

    If Not DatosOk Then Exit Sub

    cadselect = tabla & ".intconta=0 "
    cadselect = cadselect & " and " & tabla & ".codtipom = " & DBSet(Mid(Combo1(1).Text, 1, 3), "T")

    'D/H Fecha factura
    cDesde = Trim(txtcodigo(9).Text)
    cHasta = Trim(txtcodigo(10).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If

    'D/H numero de factura
    cDesde = Trim(txtcodigo(7).Text)
    cHasta = Trim(txtcodigo(8).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If

    If Not HayRegParaInforme(tabla, cadselect) Then Exit Sub

    ContabilizarFacturas tabla, cadselect
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("CONVAR") 'CONtabilizar facturas VARias

eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilizaci�n de facturas varias. Llame a soporte."
    End If

    Pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""
    cmdCancel_Click (1)

End Sub

Private Sub cmdAceptarReimp_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String

Dim CadSocios As String
Dim CadClien As String
Dim CadRes As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'Tipo de movimiento:
    If Opcionlistado = 1 Or Opcionlistado = 3 Then
        Tipos = ""
        For i = 1 To ListView1(0).ListItems.Count
            If ListView1(0).ListItems(i).Checked Then
                Tipos = Tipos & DBSet(ListView1(0).ListItems(i).Key, "T") & ","
            End If
        Next i
        
        If Tipos = "" Then
            MsgBox "Debe seleccionar al menos un tipo de factura.", vbExclamation
            Exit Sub
        Else
            ' quitamos la ultima coma
            Tipos = "{fvarcabfact.codtipom} in (" & Mid(Tipos, 1, Len(Tipos) - 1) & ")"
            If Not AnyadirAFormula(cadselect, Tipos) Then Exit Sub
            Tipos = Replace(Replace(Tipos, "(", "["), ")", "]")
            If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
        End If
    End If
    
    'D/H Seccion
    cDesde = Trim(txtcodigo(64).Text)
    cHasta = Trim(txtcodigo(65).Text)
    nDesde = txtNombre(64).Text
    nHasta = txtNombre(65).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsecci}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSeccion= """) Then Exit Sub
    End If
    
    'D/H Socio
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
 '       If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
        cadParam = cadParam & AnyadirParametroDH("pDHSocio= """, cDesde, cHasta, nDesde, nHasta)
        numParam = numParam + 1
    End If
    
    If Opcionlistado = 1 Or Opcionlistado = 3 Then
        'D/H Clientes
        cDesde = Trim(txtcodigo(62).Text)
        cHasta = Trim(txtcodigo(63).Text)
        nDesde = txtNombre(62).Text
        nHasta = txtNombre(63).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & tabla & ".codclien}"
            TipCod = "N"
    '        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
            cadParam = cadParam & AnyadirParametroDH("pDHCliente= """, cDesde, cHasta, nDesde, nHasta)
            numParam = numParam + 1
        End If
    End If
    
    CadSocios = ""
    If txtcodigo(0).Text <> "" Then CadSocios = CadSocios & "{" & tabla & ".codsocio}>= " & txtcodigo(0).Text
    If txtcodigo(1).Text <> "" Then CadSocios = CadSocios & " and {" & tabla & ".codsocio}<= " & txtcodigo(1).Text
    If CadSocios <> "" Then
        CadSocios = "(" & CadSocios & ")"
    Else
        CadSocios = "({" & tabla & ".codsocio}>=0 and {" & tabla & ".codsocio}<=9999999)"
    End If
    
    If Opcionlistado = 1 Or Opcionlistado = 3 Then
        CadClien = ""
        If txtcodigo(62).Text <> "" Then CadClien = CadClien & "{" & tabla & ".codclien}>= " & txtcodigo(62).Text
        If txtcodigo(63).Text <> "" Then CadClien = CadClien & " and {" & tabla & ".codclien}<= " & txtcodigo(63).Text
        If CadClien <> "" Then
            CadClien = "(" & CadClien & ")"
        Else
            CadClien = "({" & tabla & ".codclien}>=0 and {" & tabla & ".codclien}<=999999)"
        End If
    End If
    
    CadRes = ""
    If CadSocios <> "" Then CadRes = CadRes & CadSocios
    
    If Opcionlistado = 1 Or Opcionlistado = 3 Then
        If CadClien <> "" Then
            If CadRes <> "" Then CadRes = CadRes & " or "
            CadRes = CadRes & CadClien
        End If
        If CadRes <> "" Then
            CadRes = "(" & CadRes & ")"
            If Not AnyadirAFormula(cadselect, CadRes) Then Exit Sub
            If Not AnyadirAFormula(cadFormula, CadRes) Then Exit Sub
        End If
    End If
    
    'D/H Nro de Factura
    cDesde = Trim(txtcodigo(4).Text)
    cHasta = Trim(txtcodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    If HayRegistros(tabla, cadselect) Then
        If CargarTemporal(cadselect) Then
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        
            Select Case Opcionlistado
                Case 1
                    indRPT = 89 'Impresion de Facturas Varias
                    ConSubInforme = True
                    cadTitulo = "Reimpresi�n de Facturas Varias"
                    
                    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                    cadNombreRPT = nomDocu
                Case 3
                    indRPT = 90 'Diario de facturacion de facturas varias
                    ConSubInforme = True
                    cadTitulo = "Diario de Facturaci�n Facturas Varias"
                    
                    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                    cadNombreRPT = nomDocu
                Case 5
                    indRPT = 91 'Impresion de Facturas Varias Proveedor
                    ConSubInforme = True
                    cadTitulo = "Reimpresi�n Facturas Varias Proveedor"
                    
                    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                    cadNombreRPT = nomDocu
                Case 6
                    indRPT = 92 'Diario de facturacion de facturas varias proveedor
                    ConSubInforme = True
                    cadTitulo = "Diario Facturaci�n Varias Proveedor"
                    
                    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                    cadNombreRPT = nomDocu
            End Select
            
            LlamarImprimir
        End If
    End If
End Sub


Private Function CargarTemporal(cadselect As String)
Dim SQL As String

    On Error GoTo eCargarTemporal

    CargarTemporal = False

    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    SQL = "insert into tmpinformes (codusu, campo1, importe2, codigo1, nombre1,importe1, fecha1, importe3) "
    SQL = SQL & " select " & vUsu.Codigo & ",0,codsocio, codsecci, codtipom, numfactu, fecfactu, codforpa from " & tabla
    SQL = SQL & " where not codsocio is null and codsocio <> 0 "
    If cadselect <> "" Then SQL = SQL & " and " & cadselect
    
    If Opcionlistado = 1 Or Opcionlistado = 3 Then
        SQL = SQL & " union "
        SQL = SQL & " select " & vUsu.Codigo & ",1,codclien, codsecci, codtipom, numfactu, fecfactu, codforpa from " & tabla
        SQL = SQL & " where not codclien is null and codclien <> 0 "
        If cadselect <> "" Then SQL = SQL & " and " & cadselect
    End If
    
    conn.Execute SQL
    
    '[Monica]17/01/2019: para el caso de frutas inma cargamos el iban
    Dim Iban As String
    Dim Rs As ADODB.Recordset
    
    SQL = "select distinct codigo1, importe3 from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(DBLet(Rs!Codigo1, "N")) Then
            If vSeccion.AbrirConta Then
                Iban = DevuelveDesdeBDNew(cConta, "formapago", "iban", "codforpa", DBLet(Rs!importe3, "N"), "N")
                SQL = "update tmpinformes set nombre2 = " & DBSet(Iban, "T") & " where codusu = " & DBSet(vUsu.Codigo, "N")
                SQL = SQL & " and codigo1 = " & DBSet(Rs!Codigo1, "N")
                SQL = SQL & " and importe3 = " & DBSet(Rs!importe3, "N")
                
                conn.Execute SQL
            End If
        End If
        Set vSeccion = Nothing
            
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargando temporal", Err.Description
End Function

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCancelReimp_Click()
    Unload Me
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Index = 0 Then
        If Mid(Combo1(0).Text, 1, 3) = "FVG" Then
            If Not Option1(0).Value Then
                Me.Option1(0).Value = True
                Option1_Click (0)
            End If
        End If
    End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcionlistado
            Case 1, 3    '1 = reimpresion de facturas varias
                         '3 = diario de facturacion
                PonerFoco txtcodigo(4)
                
            Case 2 ' Grabacion de Facturas Varias
                Combo1(0).ListIndex = 0
                Combo1(2).ListIndex = 0
                
                txtcodigo(52).Text = vParamAplic.Seccionhorto
                PonerFormatoEntero txtcodigo(52)
                txtNombre(52).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", txtcodigo(52).Text, "N")
            
                txtcodigo(11).Text = Format(Now, "dd/mm/yyyy")
                
            Case 4, 7 '4 = integracion contable varias clientes
                      '7 = integracion contable varias de proveedores
                Combo1(1).ListIndex = 0
                txtcodigo(17).Text = Format(Now, "dd/mm/yyyy")
                If Opcionlistado = 7 Then
                    txtcodigo(15).Text = Format(Now, "dd/mm/yyyy")
                End If
                PonerFoco txtcodigo(6)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim i As Integer


    'Icono del formulario
    Me.Icon = frmPpal.Icon


    PrimeraVez = True
    limpiar Me

    For H = 0 To 12
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 52 To 52
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 47 To 48
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 62 To 65
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 73 To 74
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    'Ocultar todos los Frames de Formulario
    FrameReimpresion.visible = False
    FrameCargaMasivaFras.visible = False
    FrameIntConta.visible = False
    FrameFacturaMaquila.visible = False
    
    '###Descomentar
'    CommitConexion
    
    Select Case Opcionlistado
    Case 1, 3   '1= Reimpresion de facturas VARIAS
                '3= Diario de facturacion
        FrameReimpresionVisible True, H, W
        tabla = "fvarcabfact"
        CargarListView (0)
        
        If Opcionlistado = 1 Then
            Label1(0).Caption = "Reimpresi�n de Facturas Varias"
        Else
            Label1(0).Caption = "Diario de Facturaci�n"
        End If
        
    Case 2  ' Carga masiva de facturas varias
        FrameCargaMasivaFrasVisible True, H, W
        
        Option1(0).Value = True
        Me.FrameSocio.visible = True
        Me.FrameClientes.visible = False
        txtcodigo(73).TabIndex = 39
        txtcodigo(74).TabIndex = 40
    
    
        CargaCombo
        Me.Pb1.visible = False
        Me.lblProgres(0).visible = False
        Me.lblProgres(1).visible = False
    
    Case 4, 7 '4 = integracion contable registro iva cliente
              '7 = integracion contable registro iva proveedor
        If Opcionlistado = 4 Then
            frmFVARListados.Caption = "Contabilizaci�n de Facturas Varias"
            tabla = "fvarcabfact"
        Else
            frmFVARListados.Caption = "Contabilizaci�n de Facturas Varias Proveedores"
            tabla = "fvarcabfactpro"
        End If
        
        FrameIntContaVisible True, H, W
    
        txtcodigo(6).Text = Format(vParamAplic.Seccionhorto, "000")
        txtNombre(6).Text = PonerNombreDeCod(txtcodigo(6), "rseccion", "nomsecci", "codsecci", "N")
        
        CargaCombo
        
        ' la fecha de recepcion es solo para facturas de proveedor
        txtcodigo(15).visible = (Opcionlistado = 7)
        txtcodigo(15).Enabled = (Opcionlistado = 7)
        Me.Label4(18).visible = (Opcionlistado = 7)
        imgFec(0).visible = (Opcionlistado = 7)
        imgFec(0).Enabled = (Opcionlistado = 7)
        
        ConexionConta
        
        ' formas de pago
        txtcodigo(16).Text = Format(vParamAplic.ForpaPosi, "000")
        If vParamAplic.ContabilidadNueva Then
            txtNombre(16).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtcodigo(16).Text, "N")
            txtcodigo(14).Text = Format(vParamAplic.ForpaNega, "000")
            txtNombre(14).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtcodigo(14).Text, "N")
        Else
            txtNombre(16).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(16).Text, "N")
            txtcodigo(14).Text = Format(vParamAplic.ForpaNega, "000")
            txtNombre(14).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(14).Text, "N")
        End If
        ' cuentas contables
        txtcodigo(18).Text = vParamAplic.CtaBancoSoc   ' cuenta contable de banco prevista
        txtNombre(18).Text = PonerNombreCuenta(txtcodigo(18), 0)
'        txtcodigo(13).Text = vParamAplic.CtaRetenSoc ' cuenta contable de retencion
'        txtNombre(13).Text = PonerNombreCuenta(txtcodigo(13), 0)
'        txtcodigo(12).Text = vParamAplic.CtaAportaSoc ' cuenta contable de aportacion
'        txtNombre(12).Text = PonerNombreCuenta(txtcodigo(12), 0)
        
    
    Case 5, 6  '5= Reimpresion de facturas VARIAS Proveedor
               '6=  Diario de Facturacion
        FrameReimpresionVisible True, H, W
        tabla = "fvarcabfactpro"
        
        FrameTipoFactura.visible = False
        FrameTipoFactura.Enabled = False
        
        'escondemos d/h cliente
        For i = 50 To 52
            Label4(i).visible = False
        Next i
        For i = 62 To 63
            imgBuscar(i).visible = False
            imgBuscar(i).Enabled = False
            txtcodigo(i).visible = False
            txtcodigo(i).Enabled = False
            txtNombre(i).visible = False
            txtNombre(i).Enabled = False
        Next i
        
        'subimos botones de aceptar y cancelar
        cmdAceptarReimp.Top = 5300
        cmdCancelReimp.Top = 5300
        
        If Opcionlistado = 5 Then
            Label1(0).Caption = "Reimpresi�n Facturas Varias Proveedor"
        Else
            Label1(0).Caption = "Diario Facturaci�n Varias Proveedor"
        End If
        
    Case 8 ' factura de maquila
        FrameFacturaMaquilaVisible True, H, W
        CargaCombo
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Select Case Opcionlistado
        Case 4 ' Integracion contable
            If Not vSeccion Is Nothing Then
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If
    End Select
End Sub






Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(CByte(imgFec(3).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de calidades
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub




Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") ' codigo de cliente
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de conceptos
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
' form de consulta de formas de pago
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        If Option1(0).Value Then
            SQL = "rsocios.codsocio in (" & CadenaSeleccion & ")"
        Else
            SQL = "clientes.codclien in (" & CadenaSeleccion & ")"
        End If
    Else
        If Option1(0).Value Then
            SQL = "rsocios.codsocio is null "
        Else
            SQL = "clientes.codclien is null "
        End If
    End If

    cadselect = SQL

End Sub

Private Sub InsertarTemporal(Variedades As String)
Dim SQL As String
Dim Sql2 As String

    On Error GoTo eInsertarTemporal

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    If Variedades <> "" Then
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, fecha2, importe2)     "
        SQL = SQL & " select " & vUsu.Codigo & ", rprecios.codvarie, rprecios.fechaini, rprecios.fechafin, max(contador) from rprecios inner join variedades on rprecios.codvarie = variedades.codvarie "
        SQL = SQL & " where " & Replace(Replace(Variedades, "{", ""), "}", "")
        SQL = SQL & " and rprecios.fechaini = " & DBSet(txtcodigo(6).Text, "F")
        SQL = SQL & " and rprecios.fechafin = " & DBSet(txtcodigo(7).Text, "F")
        SQL = SQL & " group by 1,2,3,4 "
        
        conn.Execute SQL
        
    End If
    Exit Sub
    
eInsertarTemporal:
    MuestraError Err.Number, "Insertar Temporal", Err.Description
End Sub


Private Sub frmMens1_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(50).Text = RecuperaValor(CadenaSeleccion, 1)
End Sub


Private Sub frmMens2_DatoSeleccionado(CadenaSeleccion As String)
    albaranes = CadenaSeleccion
End Sub

Private Sub frmMens3_datoseleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        SQL = " {rclasifica.codsocio} in (" & CadenaSeleccion & ")"
        Sql2 = " {rclasifica.codsocio} in [" & CadenaSeleccion & "]"
    Else
        SQL = " {rclasifica.codsocio} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, SQL) Then Exit Sub


End Sub

Private Sub frmMens4_DatoSeleccionado(CadenaSeleccion As String)

    vReturn = 2
    If CadenaSeleccion <> "" Then vReturn = CInt(CadenaSeleccion)

End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub Image1_Click(Index As Integer)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        ' reimpresion de facturas socios
        Case 0
            For i = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(i).Checked = True
            Next i
        Case 1
            For i = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(i).Checked = False
            Next i
        ' informe de resultados y listado de retenciones
        Case 2
            For i = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(i).Checked = True
            Next i
        Case 3
            For i = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(i).Checked = False
            Next i
    End Select
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1, 24, 25  'SOCIOS
            AbrirFrmSocios (Index)
        
        Case 47, 48, 62, 63 'CLIENTES
            indCodigo = Index
            Set frmCli = New frmBasico2
            
            AyudaClienteCom frmCli
            
            Set frmCli = Nothing
            
        Case 52, 64, 65 'SECCIONES
            indCodigo = Index
            Set frmSec = New frmManSeccion
            frmSec.DatosADevolverBusqueda = "0|1|"
            frmSec.Show vbModal
            Set frmSec = Nothing
    
        Case 2 'SECCION
            indCodigo = Index + 4
            Set frmSec = New frmManSeccion
            frmSec.DatosADevolverBusqueda = "0|1|"
            frmSec.Show vbModal
            Set frmSec = Nothing
        
        Case 73, 74 ' socios
            AbrirFrmSocios (Index)
            
        Case 7 ' forma de pago
            AbrirFrmForpa (Index + 39)
        
        Case 6 'concepto
            Indice = 71
            AbrirFrmConceptos Indice
    
        Case 3 ' forma de pago positivas
            AbrirFrmForpa (16)
        Case 9  ' forma de pago negativas
            AbrirFrmForpa (14)
        Case 4 'cuenta contable banco
            AbrirFrmCuentas (18)
        Case 5 'cuenta contable banco
            AbrirFrmCuentas (12)
    
        '[Monica]27/05/2019: generacion de factura de maquila
        Case 8 'seccion
            indCodigo = 19
            Set frmSec = New frmManSeccion
            frmSec.DatosADevolverBusqueda = "0|1|"
            frmSec.Show vbModal
            Set frmSec = Nothing
        Case 10 'concepto
            Indice = 20
            AbrirFrmConceptos Indice
        Case 11 ' forma de pago
            AbrirFrmForpa (21)
        Case 12 'CLIENTES
            indCodigo = 22
            Set frmCli = New frmBasico2
            AyudaClienteCom frmCli
            Set frmCli = Nothing
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

    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    Select Case Index
        Case 0
            Indice = 15
        Case 1
            Indice = 17
        Case 2
            Indice = 13
        Case 3, 4
            Indice = Index - 1
        Case 5
            Indice = 9
        Case 6
            Indice = 10
        Case 11, 12
            Indice = Index + 21
        Case 18
            Indice = 11
    End Select

    imgFec(3).Tag = Indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(Indice).Text <> "" Then frmC.NovaData = txtcodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(CByte(imgFec(3).Tag)) '<===
    ' ********************************************

End Sub






Private Sub Option1_Click(Index As Integer)
    
    Me.FrameSocio.visible = (Option1(0).Value = True)
    Me.FrameClientes.visible = (Option1(0).Value = False)
    If Me.FrameSocio.visible Then
        txtcodigo(73).TabIndex = 39
        txtcodigo(74).TabIndex = 40
    
        txtcodigo(47).TabIndex = 102
        txtcodigo(48).TabIndex = 103
    
    Else
        txtcodigo(47).TabIndex = 39
        txtcodigo(48).TabIndex = 40
        
        txtcodigo(73).TabIndex = 100
        txtcodigo(74).TabIndex = 101
    End If
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
    If Opcionlistado = 10 Then
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
            ' reimpresion de facturas
            Case 0: KEYBusqueda KeyAscii, 0 'socio desde
            Case 1: KEYBusqueda KeyAscii, 1 'socio hasta
            Case 62: KEYBusqueda KeyAscii, 62 'cliente desde
            Case 63: KEYBusqueda KeyAscii, 63 'cliente hasta
            Case 64: KEYBusqueda KeyAscii, 64 'seccion desde
            Case 65: KEYBusqueda KeyAscii, 65 'seccion hasta
            Case 2: KEYFecha KeyAscii, 3 'fecha desde
            Case 3: KEYFecha KeyAscii, 4 'fecha hasta
            
            ' contabilizacion de facturas
            Case 9: KEYFecha KeyAscii, 5 'fecha factura desde
            Case 10: KEYFecha KeyAscii, 6 'fecha factura hasta
            Case 15: KEYFecha KeyAscii, 0 'fecha recepcion
            Case 17: KEYFecha KeyAscii, 1 'fecha vto
            
            ' insercion de cobros en tesoreria
            Case 13: KEYFecha KeyAscii, 2 'fecha vto
            Case 12: KEYBusqueda KeyAscii, 65 'seccion hasta
            
            ' generacion de factura de maquila
            Case 19: KEYBusqueda KeyAscii, 8 'seccion
            Case 20: KEYBusqueda KeyAscii, 10 'concepto
            Case 25: KEYFecha KeyAscii, 7 'fecha desde
            Case 26: KEYFecha KeyAscii, 8 'fecha hasta
            Case 22: KEYBusqueda KeyAscii, 12 'cliente
            Case 27: KEYFecha KeyAscii, 9 'fecha factura
            Case 21: KEYBusqueda KeyAscii, 11 'forma de pago
            
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
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1, 73, 74 'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            
        Case 62, 63, 47, 48, 22 'CLIENTES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            
                '[Monica]27/05/2019: seccion 19 factura de maquila
        Case 6, 52, 64, 65, 19 'SECCIONES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rseccion", "nomsecci", "codsecci", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
        
            '[Monica]26/11/2018: no abria la seccion
            If (Index = 52 Or Index = 6 Or Index = 19) Then ' And txtCodigo(Index) <> "" Then
                If txtcodigo(Index).Text = "" Then
                    PonerFoco txtcodigo(Index)
                Else
                    Set vSeccion = New CSeccion
                    If vSeccion.LeerDatos(txtcodigo(Index).Text) Then
                        txtNombre(Index).Text = vSeccion.Nombre
                        If vSeccion.AbrirConta Then
                
                        End If
                    End If
                End If
            End If
        
        Case 4, 5 ' NROS DE FACTURA
            PonerFormatoEntero txtcodigo(Index)
            
        Case 9, 10, 11, 13, 15, 17, 27 ' fecha de factura
            PonerFormatoFecha txtcodigo(Index)
            
        Case 14, 16, 46, 21 ' forma de pago
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtNombre(Index).Text = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", txtcodigo(Index).Text, "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(Index)
                End If
            End If
            
        Case 20, 71 ' concepto
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "fvarconce", "nomconce", "codconce", "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "No existe el Concepto. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(Index)
                End If
            End If
        
        Case 2, 3 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
        Case 25, 26 ' fecha de albaran ( factura de maquila )
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
        Case 12, 18
            If vSeccion Is Nothing Then Exit Sub
        
            If txtcodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtcodigo(Index), 2)
            If txtNombre(Index).Text = "" Then
                MsgBox "N�mero de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
        
        Case 70 ' cantidad
            PonerFormatoDecimal txtcodigo(Index), 3
            txtcodigo(68).Text = Round2(CCur(ComprobarCero(txtcodigo(69).Text)) * CCur(ComprobarCero(txtcodigo(70).Text)), 2)
            PonerFormatoDecimal txtcodigo(68), 3
        
        Case 69 ' precio
            PonerFormatoDecimal txtcodigo(Index), 11
            txtcodigo(68).Text = Round2(CCur(ComprobarCero(txtcodigo(69).Text)) * CCur(ComprobarCero(txtcodigo(70).Text)), 2)
            PonerFormatoDecimal txtcodigo(68), 3
            
        Case 68 ' importe
            PonerFormatoDecimal txtcodigo(Index), 3
    
    
        Case 23 ' precion por kilo para la factura de maquila
            PonerFormatoDecimal txtcodigo(Index), 7
    
    End Select
End Sub


Private Sub FrameFacturaMaquilaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameFacturaMaquila.visible = visible
    If visible = True Then
        Me.FrameFacturaMaquila.Top = -90
        Me.FrameFacturaMaquila.Left = 0
        Me.FrameFacturaMaquila.Width = 6855
        W = Me.FrameFacturaMaquila.Width
        H = Me.FrameFacturaMaquila.Height
    End If
End Sub



Private Sub FrameReimpresionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameReimpresion.visible = visible
    If visible = True Then
        Me.FrameReimpresion.Top = -90
        Me.FrameReimpresion.Left = 0
        If Opcionlistado = 5 Or Opcionlistado = 6 Then
            Me.FrameReimpresion.Height = 6110
        Else
            Me.FrameReimpresion.Height = 7110
        End If
        Me.FrameReimpresion.Width = 6855
        W = Me.FrameReimpresion.Width
        H = Me.FrameReimpresion.Height
    End If
End Sub


Private Sub FrameCargaMasivaFrasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCargaMasivaFras.visible = visible
    If visible = True Then
        Me.FrameCargaMasivaFras.Top = -90
        Me.FrameCargaMasivaFras.Left = 0
        Me.FrameCargaMasivaFras.Height = 9060
        Me.FrameCargaMasivaFras.Width = 8890
        W = Me.FrameCargaMasivaFras.Width
        H = Me.FrameCargaMasivaFras.Height
    End If
End Sub


Private Sub FrameIntContaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameIntConta.visible = visible
    If visible = True Then
        Me.FrameIntConta.Top = -90
        Me.FrameIntConta.Left = 0
        Me.FrameIntConta.Height = 6780
        Me.FrameIntConta.Width = 7680
        W = Me.FrameIntConta.Width
        H = Me.FrameIntConta.Height
    End If
End Sub




Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadSelect1 = ""
    cadSelect2 = ""
    cadParam = ""
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
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadselect, devuelve2) Then Exit Function
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
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = Opcionlistado
        .ConSubInforme = ConSubInforme
        .Show vbModal
    End With
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

Private Sub AbrirFrmForpa(Indice As Integer)
    indCodigo = Indice
    Set frmFPa = New frmComFpa
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtcodigo(indCodigo)
    frmFPa.DeConsulta = True
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub


Private Sub AbrirFrmConceptos(Indice As Integer)
    indCodigo = Indice
    Set frmCon = New frmFVARConceptos
    frmCon.DatosADevolverBusqueda = "0|1|"
    frmCon.Show vbModal
    Set frmCon = Nothing
End Sub

Private Sub AbrirFrmCuentas(Indice As Integer)
    indCodigo = Indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtcodigo(indCodigo)
    frmCtas.Show vbModal
    Set frmCtas = Nothing
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
        .Opcion = Opcionlistado
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


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim vClien As cSocio
' a�adido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String
Dim cad As String

    B = True
    Select Case Opcionlistado
        Case 2 ' carga masiva de facturas varias
            DatosOk = False
        
            If txtcodigo(52).Text = "" Then
                MsgBox "Debe introducir obligatoriamente una Secci�n.", vbExclamation
                PonerFoco txtcodigo(52)
                Exit Function
            Else
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(txtcodigo(52).Text) Then
                    txtNombre(52).Text = vSeccion.Nombre
                    
                    If vSeccion.AbrirConta Then
                    
                    End If
                End If
            End If
        
            If txtcodigo(11).Text = "" Then
                MsgBox "Debe introducir obligatoriamente una Fecha de Factura.", vbExclamation
                PonerFoco txtcodigo(11)
                Exit Function
            End If
            
            
            '[Monica]20/06/2017: en el caso de generar las facturas miramos la fecha de factura
            If Check1(1).Value = 0 Then
                ResultadoFechaContaOK = EsFechaOKConta(CDate(txtcodigo(11)))
                If ResultadoFechaContaOK > 0 Then
                    If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                    PonerFoco txtcodigo(11)
                    Exit Function
                End If
            End If
            
            
            If Check1(0).Value = 0 Then
                'codigo de socio o de cliente
                If Option1(0).Value Then
                    If txtcodigo(73).Text = "" Or txtcodigo(74).Text = "" Then
                        MsgBox "Obligatoriamente ha de poner el rango de socios. Revise.", vbExclamation
                        PonerFoco txtcodigo(73)
                        Exit Function
                    End If
                Else
                    If txtcodigo(47).Text = "" Or txtcodigo(48).Text = "" Then
                        MsgBox "Obligatoriamente ha de poner el rango de clientes. Revise.", vbExclamation
                        PonerFoco txtcodigo(47)
                        Exit Function
                    End If
                End If
            End If
            
            'Forma de pago
            If txtcodigo(46).Text = "" Then
                MsgBox "Debe introducir obligatoriamente una forma de pago.", vbExclamation
                PonerFoco txtcodigo(46)
                Exit Function
            Else
                cad = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", txtcodigo(46).Text, "N")
                If cad = "" Then
                    MsgBox "Forma de Pago no existe. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(46)
                    Exit Function
                End If
            End If
                
            'concepto
            If txtcodigo(71).Text = "" Then
                MsgBox "Debe introducir obligatoriamente un concepto.", vbExclamation
                PonerFoco txtcodigo(0)
                Exit Function
            Else
                cad = ""
                cad = DevuelveDesdeBDNew(cAgro, "fvarconce", "tipoiva", "codconce", txtcodigo(71).Text, "N")
                If cad = "" Then
                    MsgBox "El concepto no tiene asociado un tipo de iva. Revise.", vbExclamation
                    PonerFoco txtcodigo(0)
                    Exit Function
                Else
                    ' comprobamos que el concepto sea de la misma seccion que la seccion que hemos pedido
                    cad = DevuelveDesdeBDNew(cAgro, "fvarconce", "codsecci", "codconce", txtcodigo(71).Text, "N")
                    If Int(ComprobarCero(cad)) <> Int(txtcodigo(52).Text) Then
                        MsgBox "El concepto debe de ser de la misma seccion que se ha pedido. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(71)
                        B = False
                    End If
                    
                End If
            End If
            
            DatosOk = True
         
        Case 4 ' integracion contable de facturas varias
            DatosOk = False
        
            If txtcodigo(6).Text = "" Then
                MsgBox "Debe introducir obligatoriamente una Secci�n.", vbExclamation
                PonerFoco txtcodigo(6)
                Exit Function
            Else
                If vSeccion Is Nothing Then
                    Set vSeccion = New CSeccion
                    If vSeccion.LeerDatos(txtcodigo(6).Text) Then
                        txtNombre(6).Text = vSeccion.Nombre
                        
                        If vSeccion.AbrirConta Then
                
                
                        End If
                    End If
                End If
            End If
        
            If Combo1(1).ListIndex = -1 Then
                MsgBox "Debe introducir obligatoriamente un Tipo de Factura.", vbExclamation
                PonerFocoCmb Combo1(1)
                Exit Function
            End If
        
            If txtcodigo(17).Text = "" Then
                MsgBox "Debe introducir obligatoriamente una Fecha de Vto.", vbExclamation
                PonerFoco txtcodigo(17)
                Exit Function
            End If
         
         
        Case 8 ' generacion de facturas de maquila
            DatosOk = False
        
            If txtcodigo(19).Text = "" Then
                MsgBox "Debe introducir obligatoriamente una Secci�n.", vbExclamation
                PonerFoco txtcodigo(19)
                Exit Function
            Else
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(txtcodigo(19).Text) Then
                    txtNombre(52).Text = vSeccion.Nombre
                    
                    If vSeccion.AbrirConta Then
                    
                    End If
                End If
            End If
        
            If txtcodigo(27).Text = "" Then
                MsgBox "Debe introducir obligatoriamente una Fecha de Factura.", vbExclamation
                PonerFoco txtcodigo(27)
                Exit Function
            End If
            
            
            '[Monica]20/06/2017: en el caso de generar las facturas miramos la fecha de factura
            If Check1(1).Value = 0 Then
                ResultadoFechaContaOK = EsFechaOKConta(CDate(txtcodigo(27)))
                If ResultadoFechaContaOK > 0 Then
                    If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                    PonerFoco txtcodigo(27)
                    Exit Function
                End If
            End If
            
            If txtcodigo(22).Text = "" Then
                MsgBox "Obligatoriamente ha de poner el codigo de cliente. Revise.", vbExclamation
                PonerFoco txtcodigo(22)
                Exit Function
            End If
            
            'Forma de pago
            If txtcodigo(21).Text = "" Then
                MsgBox "Debe introducir obligatoriamente una forma de pago.", vbExclamation
                PonerFoco txtcodigo(21)
                Exit Function
            Else
                cad = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", txtcodigo(21).Text, "N")
                If cad = "" Then
                    MsgBox "Forma de Pago no existe. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(21)
                    Exit Function
                End If
            End If
                
            'concepto
            If txtcodigo(20).Text = "" Then
                MsgBox "Debe introducir obligatoriamente un concepto.", vbExclamation
                PonerFoco txtcodigo(20)
                Exit Function
            Else
                cad = ""
                cad = DevuelveDesdeBDNew(cAgro, "fvarconce", "tipoiva", "codconce", txtcodigo(20).Text, "N")
                If cad = "" Then
                    MsgBox "El concepto no tiene asociado un tipo de iva. Revise.", vbExclamation
                    PonerFoco txtcodigo(20)
                    Exit Function
                Else
                    ' comprobamos que el concepto sea de la misma seccion que la seccion que hemos pedido
                    cad = DevuelveDesdeBDNew(cAgro, "fvarconce", "codsecci", "codconce", txtcodigo(20).Text, "N")
                    If Int(ComprobarCero(cad)) <> Int(txtcodigo(19).Text) Then
                        MsgBox "El concepto debe de ser de la misma seccion que se ha pedido. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(20)
                        B = False
                    End If
                    
                End If
            End If
            
            DatosOk = True
         
         
    End Select
    DatosOk = B

End Function



Private Function ConcatenarCampos(cTabla As String, cWhere As String) As String
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim SQL1 As String

    ConcatenarCampos = ""

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    
    SQL = "select distinct rcampos.codcampo  from " & cTabla & " where " & cWhere
    Set Rs = New ADODB.Recordset
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL1 = ""
    While Not Rs.EOF
        SQL1 = SQL1 & DBLet(Rs.Fields(0).Value, "N") & ","
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    'quitamos el ultimo or
    ConcatenarCampos = Mid(SQL1, 1, Len(SQL1) - 1)
    
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

Private Function TotalFacturasNew(cTabla As String, cWhere As String, cCampos As String) As Long
Dim SQL As String

    TotalFacturasNew = 0
    
    SQL = "SELECT  count(distinct " & cCampos & ") "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If

    TotalFacturasNew = TotalRegistros(SQL)

End Function


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim i As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de factura
    SQL = "select codtipom, nomtipom from usuarios.stipom where tipodocu = 12"
    '[Monica]18/12/2013: a�adido or opcionlistado = 2
        '[Monica]27/05/2019: factura de maquila, opcion 8
    If Opcionlistado = 4 Or Opcionlistado = 2 Or Opcionlistado = 8 Then
        SQL = SQL & " and codtipom <> 'FVP'"
    Else
        SQL = SQL & " and codtipom = 'FVP'"
    End If

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 1
    While Not Rs.EOF
        SQL = Rs.Fields(1).Value
        SQL = Rs.Fields(0).Value & " - " & SQL
        
        Combo1(0).AddItem SQL 'campo del codigo
        Combo1(0).ItemData(Combo1(0).NewIndex) = i
        
        Combo1(1).AddItem SQL 'campo del codigo
        Combo1(1).ItemData(Combo1(1).NewIndex) = i
        
        Combo1(3).AddItem SQL 'campo del codigo
        Combo1(3).ItemData(Combo1(1).NewIndex) = i
        
        i = i + 1
        Rs.MoveNext
    Wend


    'donde se descuenta
    Combo1(2).AddItem "No descuenta"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Liquidaci�n"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Anticipo"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
    Combo1(2).AddItem "En 1�factura"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 3


End Sub


Private Sub CargarListView(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

 '   ListView1.ColumnHeaders.Add , , "Tipo", 650
    ListView1(Index).ColumnHeaders.Add , , "Tipo Movimiento", 2750
    
    SQL = "SELECT codtipom, nomtipom "
    SQL = SQL & " FROM usuarios.stipom "
    SQL = SQL & " WHERE stipom.tipodocu = 12 and codtipom <> 'FVP'"
    SQL = SQL & " ORDER BY codtipom "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1(Index).ListItems.Add
        ItmX.Text = Rs.Fields(1).Value ' Format(Rs.Fields(0).Value)
        ItmX.Key = Rs.Fields(0).Value
'        ItmX.SubItems(1) = Rs.Fields(1).Value
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Tipo de Facturas.", Err.Description
    End If
End Sub




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
                    MsgBox "Hay socios sin iva en la secci�n hortofrut�cola. Revise.", vbExclamation
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



Private Sub GenerarFacturas(cadTabla As String, cadWhere As String, NumError As Long, MensError As String)

Dim SQL As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cad As String
Dim NumF As Long
Dim CabSql As String
Dim LinSql As String

Dim vTipoMov As CTiposMov
Dim NumFact As Long

Dim TipoIVA As String
Dim PorIva As String
Dim PorRec As String
Dim ImpoIva As Currency
Dim ImpoRec As Currency
Dim TotalFact As Currency

Dim Rs As ADODB.Recordset
Dim NomCuenta As String
Dim Existe As Boolean
Dim CodTipoMov As String

    On Error GoTo EContab

    SQL = "GENFAC" 'generar facturas de venta

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Generar Facturas. Hay otro usuario realizando el proceso.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    CodTipoMov = ""
    If Len(Combo1(0).Text) >= 3 Then CodTipoMov = Mid(Combo1(0).Text, 1, 3)

    conn.BeginTrans
'--
'    BorrarTMP
'    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
'    B = CrearTMP(cadTabla, cadWHERE, True)
'    If Not B Then Exit Sub
            
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
    
    NumF = DevuelveValor("select count(*) from tmpfactvarias where codusu = " & vUsu.Codigo)
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgresNew Me.Pb1, CInt(NumF)
        
    SQL = "select * from tmpfactvarias where codusu = " & DBSet(vUsu.Codigo, "N") & " order by codsoccli"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    While Not Rs.EOF
    
        IncrementarProgresNew Me.Pb1, 1
        Me.lblProgres(1).Caption = "Procesando C�digo ..."
        Me.Refresh
        DoEvents
        
        
        Set vTipoMov = New CTiposMov
        If vTipoMov.Leer(CodTipoMov) Then
            NumFact = vTipoMov.ConseguirContador(CodTipoMov)
        
            Existe = False
            Do
                SQL = "select count(*) from fvarcabfact where "
                SQL = SQL & " codtipom = " & DBSet(CodTipoMov, "T")
                SQL = SQL & " and numfactu = " & DBSet(NumFact, "N")
                SQL = SQL & " and fecfactu = " & DBSet(txtcodigo(11).Text, "F")
                If TotalRegistros(SQL) > 0 Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (CodTipoMov)
                    NumFact = vTipoMov.ConseguirContador(CodTipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            TipoIVA = ""
            PorIva = ""
            ImpoIva = 0
            TotalFact = 0
            
            TipoIVA = DevuelveDesdeBDNew(cAgro, "fvarconce", "tipoiva", "codconce", DBLet(Rs!codConce, "N"), "N")
            If CodTipoMov = "FVG" Then
                TipoIVA = vSeccion.TipIvaExento
            End If
            PorIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", TipoIVA, "N")
            PorRec = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", TipoIVA, "N")
            ImpoIva = Round2(DBLet(Rs!Importe, "N") * ComprobarCero(PorIva) / 100, 2)
            ImpoRec = Round2(DBLet(Rs!Importe, "N") * ComprobarCero(PorRec) / 100, 2)
            
            TotalFact = DBLet(Rs!Importe, "N") + ImpoIva + ImpoRec
            
            
            ' Insertamos en la cabecera de factura
            CabSql = "insert into fvarcabfact ("
            CabSql = CabSql & "codsecci,codtipom,numfactu,fecfactu,codsocio,codclien,observac,intconta,baseiva1,baseiva2,baseiva3,"
            CabSql = CabSql & "impoiva1,impoiva2,impoiva3,imporec1,imporec2,imporec3,totalfac,tipoiva1,tipoiva2,tipoiva3,"
            CabSql = CabSql & "porciva1 , porciva2, porciva3, codforpa, porcrec1, porcrec2, porcrec3, retfaccl, trefaccl, cuereten, enliquidacion)  values  "
            
            CabSql = CabSql & "(" & DBSet(txtcodigo(52).Text, "N")
            CabSql = CabSql & "," & DBSet(CodTipoMov, "T")
            CabSql = CabSql & "," & DBSet(NumFact, "N")
            CabSql = CabSql & "," & DBSet(txtcodigo(11).Text, "F")
            If Option1(0).Value Then
                CabSql = CabSql & "," & DBSet(Rs!CODSOCCLI, "N") & "," & ValorNulo
            Else
                CabSql = CabSql & "," & ValorNulo & "," & DBSet(Rs!CODSOCCLI, "N")
            End If
                
            CabSql = CabSql & "," & DBSet(txtcodigo(67).Text, "T", "S")
            CabSql = CabSql & ",0"
            CabSql = CabSql & "," & DBSet(Rs!Importe, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & DBSet(ImpoIva, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & DBSet(ImpoRec, "N", "S")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            
            CabSql = CabSql & "," & DBSet(TotalFact, "N")
            CabSql = CabSql & "," & DBSet(TipoIVA, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & DBSet(PorIva, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & DBSet(txtcodigo(46).Text, "N") ' forma de pago
            CabSql = CabSql & "," & DBSet(PorRec, "N", "S")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            
            '[Monica]29/05/2017: a�adimos donde descuenta
            CabSql = CabSql & "," & DBSet(Rs!enliquidacion, "N")
            
            CabSql = CabSql & ")"
            
            conn.Execute CabSql
            
            
            ' insertamos en la linea de factura
            LinSql = "insert into fvarlinfact (codtipom, numfactu, fecfactu, NumLinea, codConce, ampliaci, precio, cantidad, Importe, TipoIva) values "
            LinSql = LinSql & "("
            LinSql = LinSql & DBSet(CodTipoMov, "T")
            LinSql = LinSql & "," & DBSet(NumFact, "N")
            LinSql = LinSql & "," & DBSet(txtcodigo(11).Text, "F")
            LinSql = LinSql & ",1"
            LinSql = LinSql & "," & DBSet(Rs!codConce, "N")
            LinSql = LinSql & "," & DBSet(Rs!ampliaci, "T")
            LinSql = LinSql & "," & DBSet(Rs!Precio, "N")
            LinSql = LinSql & "," & DBSet(Rs!cantidad, "N")
            LinSql = LinSql & "," & DBSet(Rs!Importe, "N")
            LinSql = LinSql & "," & DBSet(TipoIVA, "N")
            LinSql = LinSql & ")"
            
            conn.Execute LinSql
            
            
            vTipoMov.IncrementarContador (CodTipoMov)
            Set vTipoMov = Nothing
                    
        End If
        
        Rs.MoveNext
    Wend
    
EContab:
    If Err.Number <> 0 Then
        NumError = Err.Number
        MensError = "Generar Facturas " '& Err.Description
        conn.RollbackTrans
    Else
        conn.CommitTrans
        
    End If
End Sub


Private Sub BorrarTMP()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpfactuvar;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function CrearTMP(cadTabla As String, cadWhere As String, Optional Facturas As Boolean, Optional Telefono As Boolean) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
' facturas indica si viene de facturas varias o de telefonia
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMP = False
    
    SQL = "CREATE TEMPORARY TABLE tmpfactuvar ( "
    SQL = SQL & "codigo int(7) NOT NULL) "
    conn.Execute SQL
     
    If cadTabla = "rsocios" Then
        SQL = "SELECT codsocio "
    Else
        SQL = "SELECT codclien "
    End If

    SQL = SQL & " FROM " & cadTabla
    SQL = SQL & " WHERE " & cadWhere
    SQL = " INSERT INTO tmpfactuvar " & SQL
    conn.Execute SQL

    CrearTMP = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMP = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpfactuvar;"
        conn.Execute SQL
    End If
End Function



Private Sub ConexionConta()
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(txtcodigo(6).Text) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(txtcodigo(6).Text) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub


Private Sub ContabilizarFacturas(cadTabla As String, cadWhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String



    SQL = "CONVAR" 'contabilizar facturas VARias
    
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Facturas Varias. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtcodigo(9).Text = "" Then
        txtcodigo(9).Text = vEmpresa.FechaIni  'fechaini del ejercicio de la conta
     End If

     If txtcodigo(10).Text = "" Then
        txtcodigo(10).Text = vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
     End If


     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los par�metros
     If Not ComprobarFechasConta(10) Then Exit Sub

    'comprobar si existen  facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtcodigo(9).Text <> "" Then 'anteriores a fechadesde
        SQL = "SELECT COUNT(*) FROM " & cadTabla
        SQL = SQL & " WHERE fecfactu <"
        SQL = SQL & DBSet(txtcodigo(9), "F") & " AND intconta=0 and codtipom = " & DBSet(Mid(Combo1(1).Text, 1, 3), "T")
        If RegistrosAListar(SQL) > 0 Then
            MsgBox "Hay Facturas anteriores sin contabilizar.", vbExclamation
            Exit Sub
        End If
    End If


    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    BorrarTMPFacturas
    B = CrearTMPFacturas(cadTabla, cadWhere)
    If Not B Then Exit Sub
    

    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    TerminaBloquear
    SQL = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla
    SQL = SQL & ".codtipom=tmpFactu.codtipom AND "
    
    SQL = SQL & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    If Not BloqueaRegistro(SQL, cadWhere) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    'Visualizar la barra de Progreso
    Me.Pb2.visible = True
    Me.lblProgres(2).Caption = "Comprobaciones: "
    CargarProgres Me.Pb2, 100


    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariagrorec
    '--------------------------------------------------------------------------
    If cadTabla = "fvarcabfact" Then ' solo si son facturas de registro de iva de cliente
        Me.lblProgres(3).Caption = "Comprobando letras de serie ..."
        B = ComprobarLetraSerie(cadTabla)
    End If
    IncrementarProgres Me.Pb2, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub


    'comprobar que todas las CUENTAS de los distintos socios que vamos a
    'contabilizar existen en la Conta: rsocios_seccion.codmaccli IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(3).Caption = "Comprobando Cuentas Contables socios en contabilidad ..."
    
    B = ComprobarCtaContable_new(cadTabla, 1, 0, CInt(txtcodigo(6).Text))
    IncrementarProgres Me.Pb2, 30
    Me.Refresh
    DoEvents
    If Not B Then Exit Sub



    'comprobar que todas las CUENTAS de conceptos que vamos a
    'contabilizar existen en la Conta: fvarconcep.codmacta  IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(3).Caption = "Comprobando Cuentas Contables Conceptos en contabilidad ..."
    B = ComprobarCtaContable_new(cadTabla, 2, , CInt(txtcodigo(6).Text))
    IncrementarProgres Me.Pb2, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub
    
'    'comprobar que todas las CUENTAS de gastos a pie de factura
'    b = ComprobarCtaContable_new(cadTabla, 12, tipo, CInt(txtcodigo(2).Text))
    IncrementarProgres Me.Pb2, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub

    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: rfactsoc.tipoiva IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(3).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    B = ComprobarTiposIVA2(cadTabla)
    IncrementarProgres Me.Pb2, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub


    Me.lblProgres(1).Caption = "Comprobando Forma de Pago ..."
    B = ComprobarFormadePago(cadTabla)
    IncrementarProgres Me.Pb2, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub


    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    If vEmpresa.TieneAnalitica Then  'hay contab. analitica
       Me.lblProgres(3).Caption = "Comprobando Contabilidad Anal�tica ..."
           
       B = ComprobarCtaContable_new(cadTabla, 7, , CInt(txtcodigo(6).Text))

       '(si tiene anal�tica requiere un centro de coste para insertar en conta.linfact)
       If B Then
            CCoste = ""
            B = ComprobarCCoste_new(CCoste, cadTabla)
            If Not B Then Exit Sub
       End If
       CCoste = ""
       '[Monica]19/12/2013
       If Not B Then Exit Sub
    End If
    IncrementarProgres Me.Pb2, 20
    Me.Refresh
    DoEvents


    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    If Opcionlistado = 4 Then
        Me.lblProgres(2).Caption = "Contabilizar Facturas Varias: "
    Else
        Me.lblProgres(2).Caption = "Contabilizar Facturas Varias Proveedor: "
    End If
    CargarProgres Me.Pb2, 10
    Me.lblProgres(3).Caption = "Insertando Facturas en Contabilidad..."


    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    If Opcionlistado = 4 Then
        LOG.Insertar 3, vUsu, "Contabilizar Facturas Varias: " & vbCrLf & cadTabla & vbCrLf & cadWhere
    Else
        LOG.Insertar 3, vUsu, "Contabilizar Facturas Varias Proveedor: " & vbCrLf & cadTabla & vbCrLf & cadWhere
    End If
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    '---- Crear tabla TEMP para los posible errores de facturas
    BorrarTMPErrFact
    tmpErrores = CrearTMPErrFact(cadTabla)

    '---- Pasar las Facturas a la Contabilidad
    B = PasarFacturasAContab(cadTabla, CCoste)

    '---- Mostrar ListView de posibles errores (si hay)
    If Not B Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensajes.OpcionMensaje = 10
            frmMensajes.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If

'    'Este bien o mal, si son proveedores abriremos el listado
'    'Imprimimiremos un listado de contabilizacion de facturas
'    '------------------------------------------------------
    If cadTabla = "fvarcabfactpro" Then
        If DevuelveValor("Select count(*) from tmpinformes where codusu = " & vUsu.Codigo) > 0 Then
            InicializarVbles
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1

            cadParam = "|pDHFecha=""" & vUsu.Nombre & "   Hora: " & Format(Now, "hh:mm") & """|"
            numParam = numParam + 1
            cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
            ConSubInforme = False
            cadTitulo = "Listado contabilizacion FRASOC"
            cadNombreRPT = "rContabSOC.rpt"

            LlamarImprimir
        End If
    End If


    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact

End Sub

Private Function PasarFacturasAContab(cadTabla As String, CCoste As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim B As Boolean
Dim i As Integer
Dim numfactu As Integer
Dim Codigo1 As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False

    '---- Obtener el total de Facturas a Insertar en la contabilidad
    SQL = "SELECT count(*) "
    SQL = SQL & " FROM " & cadTabla & " INNER JOIN tmpFactu "
    Codigo1 = "codtipom"
    SQL = SQL & " ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1
    SQL = SQL & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        numfactu = Rs.Fields(0)
    Else
        numfactu = 0
    End If
    Rs.Close
    Set Rs = Nothing


    'Modificacion como David
    '-----------------------------------------------------------
    ' Mosrtaremos para cada factura de PROVEEDOR
    ' que numregis le ha asignado
    SQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute SQL

    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        SQL = "Si continua, las facturas se insertaran en el registro, pero no ser�n contabilizadas" & vbCrLf
        SQL = SQL & "en este momento. Deber�n ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        SQL = SQL & Space(50) & "�Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If


    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If numfactu > 0 Then
        CargarProgres Me.Pb2, numfactu

        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        SQL = "SELECT * "
        SQL = SQL & " FROM tmpFactu "

        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        B = True
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not Rs.EOF
            SQL = cadTabla & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " and numfactu=" & DBSet(Rs!numfactu, "T")
            SQL = SQL & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
            
            'facturas varias de cliente
            If cadTabla = "fvarcabfact" Then
                ' tipo = 0 factura de cliente a socio
                '        1 factura de cliente a cliente
                Tipo = DevuelveValor("select if(codclien is null, 0,1) from " & cadTabla & " where " & SQL)
                
                If PasarFacturaFVAR(SQL, CCoste, Orden2, txtcodigo(6).Text, Tipo, CDate(txtcodigo(17).Text), txtcodigo(16).Text, txtcodigo(14).Text, txtcodigo(18).Text, Mid(Combo1(1).Text, 1, 3), cContaFra) = False And B Then B = False

            Else 'facturas varias de proveedor
                If PasarFacturaFVAR(SQL, CCoste, Orden2, txtcodigo(6).Text, Tipo, CDate(txtcodigo(17).Text), txtcodigo(16).Text, txtcodigo(14).Text, txtcodigo(18).Text, Mid(Combo1(1).Text, 1, 3), cContaFra, CDate(txtcodigo(15).Text)) = False And B Then B = False
            End If

            '---- Laura 26/10/2006
            'Al pasar cada factura al hacer el commit desbloqueamos los registros
            'que teniamos bloqueados y los volvemos a bloquear
            'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
            SQL = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
            If Not BloqueaRegistro(SQL, cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu") Then
'                MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'                Screen.MousePointer = vbDefault
'                Exit Sub
            End If
            '----

            IncrementarProgres Me.Pb2, 1
            Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & numfactu & ")"
            Me.Refresh
            DoEvents
            
            i = i + 1
            Rs.MoveNext
        Wend

        Rs.Close
        
        Set Rs = Nothing
    End If
    
    Set cContaFra = Nothing

EPasarFac:
    If Err.Number <> 0 Then B = False

    If B Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function



Private Function ComprobarTiposIVA2(vtabla As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA2 = False
    
    SQL = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For i = 1 To 3
            SQL = "SELECT DISTINCT " & vtabla & ".tipoiva" & i
            SQL = SQL & " FROM " & vtabla
            SQL = SQL & " INNER JOIN tmpfactu ON " & vtabla & ".codtipom=tmpfactu.codtipom AND " & vtabla & ".numfactu=tmpfactu.numfactu AND " & vtabla & ".fecfactu=tmpfactu.fecfactu "
            SQL = SQL & " WHERE not isnull(tipoiva" & i & ")"

            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            B = True
            While Not Rs.EOF 'And b
                If Rs.Fields(0) <> 0 Then ' a�adido pq en arigasol sino tiene tipo de iva pone ceros
                    SQL = "codigiva= " & DBSet(Rs.Fields(0), "N")
                    RSconta.MoveFirst
                    RSconta.Find (SQL), , adSearchForward
                    If RSconta.EOF Then
                        B = False 'no encontrado
                        SQL = "No existe el tipo de IVA: " & Rs.Fields(0) & ". Revise."
                        MsgBox SQL, vbExclamation
                    End If
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not B Then
                ComprobarTiposIVA2 = False
                Exit For
            Else
                ComprobarTiposIVA2 = True
            End If
        Next i
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function


Private Function ComprobarTiposIVA3() As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA3 = False
    
    SQL = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        SQL = "SELECT distinct tipoiva FROM fvarconce inner join tmpfactvarias on fvarconce.codconce = tmpfactvarias.codconce where tmpfactvarias.codusu = " & vUsu.Codigo

        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        B = True
        While Not Rs.EOF 'And b
            If Rs.Fields(0) <> 0 Then ' a�adido pq en arigasol sino tiene tipo de iva pone ceros
                SQL = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (SQL), , adSearchForward
                If RSconta.EOF Then
                    B = False 'no encontrado
                    SQL = "No existe el tipo de IVA: " & Rs.Fields(0) & ". Revise."
                    MsgBox SQL, vbExclamation
                End If
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    
        If Not B Then
            ComprobarTiposIVA3 = False
        Else
            ComprobarTiposIVA3 = True
        End If
    
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function





Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim cad As String
Dim Rs As ADODB.Recordset
    
Dim vSeccion As CSeccion
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    
    If txtcodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            FechaFin = DateAdd("yyyy", 1, DBLet(Rs!FechaFin, "F"))
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtcodigo(ind).Text, FechaFin) Then
                 cad = "El per�odo de contabilizaci�n debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 cad = cad & "    Desde: " & FechaIni & vbCrLf
                 cad = cad & "    Hasta: " & FechaFin
                 MsgBox cad, vbExclamation
                 txtcodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
            
    '[Monica]20/06/2017: solo para el caso de Montifrut la fecha de recepcion es la de factura, en el resto es la de recepcion
    If ComprobarFechasConta Then
        If tabla = "fvarcabfactpro" Then
            ResultadoFechaContaOK = EsFechaOKConta(CDate(txtcodigo(15)))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                ComprobarFechasConta = False
            End If
        End If
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function

Private Function GenerarFacturasPrevio(cadTabla As String, cadWhere As String, NumError As Long, MensError As String) As Boolean

Dim SQL As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cad As String
Dim NumF As Long
Dim CabSql As String
Dim LinSql As String

Dim vTipoMov As CTiposMov
Dim NumFact As Long

Dim TipoIVA As String
Dim PorIva As String
Dim PorRec As String
Dim ImpoIva As Currency
Dim ImpoRec As Currency
Dim TotalFact As Currency

Dim Rs As ADODB.Recordset
Dim NomCuenta As String
Dim Existe As Boolean
Dim CodTipoMov As String

    On Error GoTo EContab


    GenerarFacturasPrevio = False


'    If TotalRegistrosConsulta("select * from tmpfactvarias where codusu = " & vUsu.Codigo) > 0 Then
'        If MsgBox("� Desea eliminar los registros anteriormente insertados ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'            Sql = "delete from tmpfactvarias where codusu = " & vUsu.Codigo
'            conn.Execute Sql
'        End If
'    End If

    
    
    SQL = "insert into tmpfactvarias (codusu,codconce,ampliaci,precio,cantidad,importe,codsoccli)  "
    SQL = SQL & " select " & vUsu.Codigo & "," & DBSet(txtcodigo(71).Text, "N") & ","
    SQL = SQL & DBSet(txtcodigo(66).Text, "T") & "," & DBSet(txtcodigo(69).Text, "N", "S") & "," & DBSet(txtcodigo(70).Text, "N", "S") & ","
    SQL = SQL & DBSet(txtcodigo(68).Text, "N") & ","
    If cadTabla = "rsocios" Then
        SQL = SQL & "codsocio "
    Else
        SQL = SQL & "codclien "
    End If

    SQL = SQL & " FROM " & cadTabla
    SQL = SQL & " WHERE " & cadWhere

    conn.Execute SQL

    GenerarFacturasPrevio = True
    Exit Function
EContab:
    If Err.Number <> 0 Then
        NumError = Err.Number
        MensError = "Generar Facturas Previo " & Err.Description
    End If
End Function






Private Sub ContabilizarCobros(ByRef NumError As Long)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String

    NumError = 1

    SQL = "GENFAC" 'contabilizar COBROS
    
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Cobros Facturas Varias. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2

     If txtcodigo(11).Text = "" Then
        txtcodigo(11).Text = vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
     End If


     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los par�metros
     If Not ComprobarFechasConta(11) Then Exit Sub

    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================

    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100


    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariagrorec
    '--------------------------------------------------------------------------
    Me.lblProgres(0).Caption = "Comprobando letras de serie ..."
    B = ComprobarLetraSerie("tmpfactvarias")
    
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    
    If Not B Then
        MsgBox "No existe la letra de serie XX1", vbExclamation
        Exit Sub
    End If

    'comprobar que todas las CUENTAS de los distintos socios que vamos a
    'contabilizar existen en la Conta: rsocios_seccion.codmaccli IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(0).Caption = "Comprobando Cuentas Contables socios en contabilidad ..."
    
    Dim Tipo As Byte
    Tipo = 1
    If cadTabla = "rsocios" Then Tipo = 0
    
    B = ComprobarCtaContable_new("tmpfactvarias", 1, Tipo, CInt(txtcodigo(52).Text))
    IncrementarProgres Me.Pb1, 30
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub


    'comprobar que todas las CUENTAS de conceptos que vamos a
    'contabilizar existen en la Conta: fvarconcep.codmacta  IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(0).Caption = "Comprobando Cuentas Contables Conceptos en contabilidad ..."
    B = ComprobarCtaContable_new("tmpfactvarias", 2, , CInt(txtcodigo(52).Text))
    IncrementarProgres Me.Pb1, 30
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub
    
'    'comprobar que todas las CUENTAS de gastos a pie de factura
'    b = ComprobarCtaContable_new(cadTabla, 12, tipo, CInt(txtcodigo(2).Text))
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub

    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: rfactsoc.tipoiva IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(0).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    B = ComprobarTiposIVA3()
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub


    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Cobros Facturas Varias: "
    CargarProgres Me.Pb1, 10


    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar Cobros Facturas Varias: " & vbCrLf & cadTabla & vbCrLf & vUsu.Codigo
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    '---- Crear tabla TEMP para los posible errores de facturas

    '---- Pasar las Facturas a la Contabilidad
    B = ProcesarCobros

    '---- Mostrar ListView de posibles errores (si hay)
    If B Then NumError = 0

End Sub


Private Function ProcesarCobros() As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim B As Boolean
Dim i As Integer
Dim numfactu As Integer
Dim Codigo1 As String

    On Error GoTo EPasarFac

    ProcesarCobros = False

    ConnConta.BeginTrans
    
    
    
    '---- Obtener el total de Facturas a Insertar en la contabilidad
    SQL = "SELECT count(*) "
    SQL = SQL & " FROM tmpfactvarias where codusu = " & vUsu.Codigo

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        numfactu = Rs.Fields(0)
    Else
        numfactu = 0
    End If
    Rs.Close
    Set Rs = Nothing


    'Modificacion como David
    '-----------------------------------------------------------
    ' Mosrtaremos para cada factura de PROVEEDOR
    ' que numregis le ha asignado
    SQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute SQL

    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If numfactu > 0 Then
        CargarProgres Me.Pb1, numfactu

        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        SQL = "SELECT * "
        SQL = SQL & " FROM tmpfactvarias where codusu = " & vUsu.Codigo

        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        B = True
        'pasar a contabilidad cada una de las facturas seleccionadas
        
        Dim MenError As String
        
        While Not Rs.EOF
            Dim Tipo As Integer
            
            SQL = "codsoccli = " & DBSet(Rs.Fields(1).Value, "N")
            
            If Option1(3).Value Then
                If InsertarEnTesoreriaNewFVAR(SQL, MenError) = False And B Then B = False
            Else
                If InsertarEnTesoreriaNewFVARPagos(SQL, MenError) = False And B Then B = False
            End If

            IncrementarProgres Me.Pb1, 1
            Me.lblProgres(1).Caption = "Insertando Cobros ...   (" & i & " de " & numfactu & ")"
            Me.Refresh
            DoEvents
            
            i = i + 1
            Rs.MoveNext
        Wend

        Rs.Close
        
        Set Rs = Nothing
        
    End If
    

EPasarFac:
    If Err.Number <> 0 Then B = False

    If B Then
         ConnConta.CommitTrans
         ProcesarCobros = True
    Else
        ProcesarCobros = False
        ConnConta.RollbackTrans
    End If
End Function



Private Function InsertarEnTesoreriaNewFVAR(cadWhere As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim B As Boolean
Dim SQL As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset
Dim rsVenci As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim i As Long
Dim DigConta As String
Dim CC As String

Dim Iban As String
Dim CodBanco As String
Dim CodSucur As String
Dim CuentaBa As String
Dim Codmacta As String



Dim vrefer As String
Dim letraser As String
Dim Rsx As ADODB.Recordset
Dim FecVenci As Date
Dim ImpVenci As Currency
Dim ImpVenci1 As Currency
Dim AcumIva As Currency
Dim PorcIva As String

Dim Rsx7 As ADODB.Recordset
Dim Sql7 As String
Dim cadena As String

Dim CadRegistro As String
Dim CadRegistro1 As String

Dim vSocio As cSocio
Dim vvIban As String
Dim TotalFac As Currency

Dim CodTipom As String
Dim CtaBan As String
Dim NumFact As Long
        
Dim vTipoMov As CTiposMov

Dim TipoIVA As String
Dim PorIva As String
Dim PorRec As String
Dim ImpoIva As Currency
Dim ImpoRec As Currency
Dim TotalFact As Currency


    On Error GoTo EInsertarTesoreriaNewFac

    B = False
    InsertarEnTesoreriaNewFVAR = False
    
    
    CtaBan = txtcodigo(12).Text
    
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    SQL = "select * from tmpfactvarias where codusu = " & vUsu.Codigo & " and " & cadWhere
    Rsx.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rsx.EOF Then
    
        CodTipom = "XX1"
    
        Set vTipoMov = New CTiposMov
        If vTipoMov.Leer(CodTipom) Then
            letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "codtipom", "codtipom", CodTipom, "T")
        
        
            NumFact = vTipoMov.ConseguirContador(CodTipom)
        
            If cadTabla = "rsocios" Then
                ' socio
                
                Dim vSoc As cSocio
                Set vSoc = New cSocio
                
                If vSoc.LeerDatos(DBLet(Rsx!CODSOCCLI, "N")) Then
                    If vSoc.LeerDatosSeccion(DBLet(Rsx!CODSOCCLI, "N"), txtcodigo(52).Text) Then
                        B = True
                                
                        CC = DBLet(vSoc.Digcontrol, "T")
                        If DBLet(vSoc.Digcontrol, "T") = "**" Then CC = "00"
            
                        Iban = vSoc.Iban
                        CodBanco = vSoc.Banco
                        CodSucur = vSoc.Sucursal
                        CuentaBa = vSoc.CuentaBan
                        Codmacta = vSoc.CtaClien
                    End If
                End If
            Else
                ' cliente
                Sql4 = "select codbanco, codsucur, digcontr, cuentaba, codmacta, iban, nomclien,domclien,pobclien,codpobla,proclien,cifclien  from clientes where codclien = " & DBLet(Rsx!CODSOCCLI, "N")
                Set Rs4 = New ADODB.Recordset
                
                Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not Rs4.EOF Then
                    B = True
                    
                    CC = DBLet(Rs4!digcontr, "T")
                    If DBLet(Rs4!digcontr, "T") = "**" Then CC = "00"
                    
                    Iban = DBLet(Rs4!Iban, "T")
                    CodBanco = DBLet(Rs4!CodBanco, "N")
                    CodSucur = DBLet(Rs4!CodSucur, "N")
                    CuentaBa = DBLet(Rs4!CuentaBa, "T")
                    Codmacta = DBLet(Rs4!Codmacta, "T")
                End If
            End If
                
            If B Then
                
                TipoIVA = ""
                PorIva = ""
                ImpoIva = 0
                TotalFact = 0
                
                TipoIVA = DevuelveDesdeBDNew(cAgro, "fvarconce", "tipoiva", "codconce", DBLet(Rsx!codConce, "N"), "N")
                If CodTipom = "FVG" Then
                    TipoIVA = vSeccion.TipIvaExento
                End If
                PorIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", TipoIVA, "N")
                PorRec = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", TipoIVA, "N")
                ImpoIva = Round2(DBLet(Rsx!Importe, "N") * ComprobarCero(PorIva) / 100, 2)
                ImpoRec = Round2(DBLet(Rsx!Importe, "N") * ComprobarCero(PorRec) / 100, 2)
                
                TotalFact = DBLet(Rsx!Importe, "N") + ImpoIva + ImpoRec
                
                
                
                Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(NumFact, "T") & " " & Format(txtcodigo(11).Text, "dd/mm/yy") & "'"
                Text41csb = "de " & DBSet(TotalFact, "N")
                
                'Obtener el N� de Vencimientos de la forma de pago
                SQL = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(txtcodigo(46).Text, "N")
                Set rsVenci = New ADODB.Recordset
                rsVenci.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                If Not rsVenci.EOF Then
                    If DBLet(rsVenci!numerove, "N") > 0 Then
                
                        CadValuesAux2 = "('" & Trim(letraser) & "', " & DBSet(NumFact, "N") & ", " & DBSet(txtcodigo(11).Text, "F") & ", "
                        '-------- Primer Vencimiento
                        i = 1
                        'FECHA VTO
                        FecVenci = txtcodigo(13).Text
                        '=== Laura 23/01/2007
                        'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                        FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                        '===
                        
                        CadValues2 = CadValuesAux2 & i & ", "
                        
                        '[Monica]03/07/2013: a�ado trim(codmacta)
                        CadValues2 = CadValues2 & DBSet(Trim(Codmacta), "T") & ", " & DBSet(txtcodigo(46).Text, "N") & ", " & DBSet(FecVenci, "F") & ", "
                        
                        'IMPORTE del Vencimiento
                        If rsVenci!numerove = 1 Then
                            ImpVenci = DBLet(TotalFact, "N")
                        Else
                            ImpVenci = Round2(TotalFact / rsVenci!numerove, 2)
                            'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                            If ImpVenci * rsVenci!numerove <> DBLet(TotalFact, "N") Then
                                ImpVenci = Round2(ImpVenci + (DBLet(TotalFact, "N") - ImpVenci * rsVenci!numerove), 2)
                            End If
                        End If
                        
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", "
                        
                        If Not vParamAplic.ContabilidadNueva Then
                            CadValues2 = CadValues2 & DBSet(CodBanco, "N", "S") & ", " & DBSet(CodSucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(CuentaBa, "T", "S") & ", "
                            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1" '),"
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & "),"
                            Else
                                CadValues2 = CadValues2 & "),"
                            End If
                        Else
                            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1"
                            
                            vvIban = MiFormat(Iban, "") & MiFormat(CodBanco, "0000") & MiFormat(CodSucur, "0000") & MiFormat(CC, "00") & MiFormat(CuentaBa, "0000000000")
                            
                            CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                            
                            If Tipo = 0 Then ' socio
                                CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & ","
                                CadValues2 = CadValues2 & DBSet(vSoc.CPostal, "T") & "," & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'),"
                            Else ' cliente
                                'nomclien,domclien,pobclien,codpobla,proclien,cifclien
                                CadValues2 = CadValues2 & DBSet(Rs4!nomclien, "T") & "," & DBSet(Rs4!domclien, "T") & "," & DBSet(Rs4!pobclien, "T") & ","
                                CadValues2 = CadValues2 & DBSet(Rs4!CodPobla, "T") & "," & DBSet(Rs4!proclien, "T") & "," & DBSet(Rs4!cifclien, "T") & ",'ES'),"
                            End If
                        End If
                        
                    
                        'Resto Vencimientos
                        '--------------------------------------------------------------------
                        For i = 2 To rsVenci!numerove
                           'FECHA Resto Vencimientos
                            '=== Laura 23/01/2007
                            'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                            FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                            '===
                                
                            CadValues2 = CadValues2 & CadValuesAux2 & i & ", " & DBSet(Trim(Rs4!Codmacta), "T") & ", " & DBSet(txtcodigo(46).Text, "N") & ", '" & Format(FecVenci, FormatoFecha) & "', "
                            
                            'IMPORTE Resto de Vendimientos
                            ImpVenci = Round2(TotalFact / rsVenci!numerove, 2)
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", "
                            
                            If Not vParamAplic.ContabilidadNueva Then
                                CadValues2 = CadValues2 & DBSet(Rs4!CodBanco, "N", "S") & ", " & DBSet(Rs4!CodSucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!CuentaBa, "T", "S") & ", "
                                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1" '),"
                                '[Monica]22/11/2013: Tema iban
                                If vEmpresa.HayNorma19_34Nueva = 1 Then
                                    CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & "),"
                                Else
                                    CadValues2 = CadValues2 & "),"
                                End If
                            Else
                                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1"
                                
                                vvIban = MiFormat(Iban, "") & MiFormat(DBLet(Rs4!CodBanco), "0000") & MiFormat(DBLet(Rs4!CodSucur), "0000") & MiFormat(CC, "00") & MiFormat(DBLet(Rs4!CuentaBa), "0000000000")
                                
                                CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                                
                                If Tipo = 0 Then ' socio
                                    CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & ","
                                    CadValues2 = CadValues2 & DBSet(vSoc.CPostal, "T") & "," & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'),"
                                Else ' cliente
                                    'nomclien,domclien,pobclien,codpobla,proclien,cifclien
                                    CadValues2 = CadValues2 & DBSet(Rs4!nomclien, "T") & "," & DBSet(Rs4!domclien, "T") & "," & DBSet(Rs4!pobclien, "T") & ","
                                    CadValues2 = CadValues2 & DBSet(Rs4!CodPobla, "T") & "," & DBSet(Rs4!proclien, "T") & "," & DBSet(Rs4!cifclien, "T") & ",'ES'),"
                                End If
                            End If
                        Next i
                        ' quitamos la ultima coma
                        CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)
                            
                        If vParamAplic.ContabilidadNueva Then
                            SQL = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                            SQL = SQL & "ctabanc1,  fecultco, impcobro, "
                            SQL = SQL & " text33csb, text41csb, agente, iban, " ') "
                            SQL = SQL & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
                            SQL = SQL & ") "
                        
                        Else
                            'Insertamos en la tabla scobro de la CONTA
                            SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                            SQL = SQL & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                            SQL = SQL & " text33csb, text41csb, agente" ') "
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                SQL = SQL & ", iban) "
                            Else
                                SQL = SQL & ") "
                            End If
                        End If
                        
                        SQL = SQL & " VALUES " & CadValues2
                        ConnConta.Execute SQL
                    
                    End If
                End If
            
                vTipoMov.IncrementarContador (CodTipom)
                Set vTipoMov = Nothing
            
                B = True
    
            End If
        End If
    End If
    
EInsertarTesoreriaNewFac:
    If Err.Number <> 0 Then
        B = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewFVAR = B
End Function




Private Function InsertarEnTesoreriaNewFVARPagos(cadWhere As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim B As Boolean
Dim SQL As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset
Dim rsVenci As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim i As Long
Dim DigConta As String
Dim CC As String

Dim Iban As String
Dim CodBanco As String
Dim CodSucur As String
Dim CuentaBa As String
Dim Codmacta As String
Dim codMacta2 As String



Dim vrefer As String
Dim letraser As String
Dim Rsx As ADODB.Recordset
Dim FecVenci As Date
Dim ImpVenci As Currency
Dim ImpVenci1 As Currency
Dim AcumIva As Currency
Dim PorcIva As String

Dim Rsx7 As ADODB.Recordset
Dim Sql7 As String
Dim cadena As String

Dim CadRegistro As String
Dim CadRegistro1 As String

Dim vSocio As cSocio
Dim vvIban As String
Dim TotalFac As Currency

Dim CodTipom As String
Dim CtaBan As String
Dim NumFact As Long
        
Dim vTipoMov As CTiposMov

Dim TipoIVA As String
Dim PorIva As String
Dim PorRec As String
Dim ImpoIva As Currency
Dim ImpoRec As Currency
Dim TotalFact As Currency

Dim vSoc As cSocio
Dim Text42csb As String

Dim Nombre As String
Dim Direccion As String
Dim Poblacion As String
Dim CodPobla As String
Dim Provincia As String
Dim nif As String


    On Error GoTo EInsertarTesoreriaNewFac

    B = False
    
    InsertarEnTesoreriaNewFVARPagos = False
    
    CtaBan = txtcodigo(12).Text
    
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    SQL = "select * from tmpfactvarias where codusu = " & vUsu.Codigo & " and " & cadWhere
    Rsx.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rsx.EOF Then
    
        If cadTabla = "rsocios" Then
            ' socio
            Set vSoc = New cSocio
            
            If vSoc.LeerDatos(DBLet(Rsx!CODSOCCLI, "N")) Then
                If vSoc.LeerDatosSeccion(DBLet(Rsx!CODSOCCLI, "N"), txtcodigo(52).Text) Then
                    B = True
                            
                    CC = DBLet(vSoc.Digcontrol, "T")
                    If DBLet(vSoc.Digcontrol, "T") = "**" Then CC = "00"
        
                    Iban = vSoc.Iban
                    CodBanco = vSoc.Banco
                    CodSucur = vSoc.Sucursal
                    CuentaBa = vSoc.CuentaBan
                    Codmacta = vSoc.CtaProv
                    codMacta2 = vSoc.CtaClien
                    
                    Nombre = vSoc.Nombre
                    Direccion = vSoc.Direccion
                    Poblacion = vSoc.Poblacion
                    CodPobla = vSoc.CPostal
                    Provincia = vSoc.Provincia
                    nif = vSoc.nif
                End If
            End If
        Else
            ' cliente
            Sql4 = "select codbanco, codsucur, digcontr, cuentaba, codmacta, iban, nomclien,domclien,pobclien,codpobla,proclien,cifclien  from clientes where codclien = " & DBLet(Rsx!CODSOCCLI, "N")
            Set Rs4 = New ADODB.Recordset
            
            Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs4.EOF Then
                B = True
                
                CC = DBLet(Rs4!digcontr, "T")
                If DBLet(Rs4!digcontr, "T") = "**" Then CC = "00"
                
                Iban = DBLet(Rs4!Iban, "T")
                CodBanco = DBLet(Rs4!CodBanco, "N")
                CodSucur = DBLet(Rs4!CodSucur, "N")
                CuentaBa = DBLet(Rs4!CuentaBa, "T")
                Codmacta = DBLet(Rs4!Codmacta, "T")
                codMacta2 = Codmacta
            
                Nombre = DBLet(Rs4!nomclien, "T")
                Direccion = DBLet(Rs4!domclien, "T")
                Poblacion = DBLet(Rs4!pobclien, "T")
                CodPobla = DBLet(Rs4!CodPobla, "T")
                Provincia = DBLet(Rs4!proclien, "T")
                nif = DBLet(Rs4!cifclien, "T")
            
            
            
            End If
        End If
    
        If B Then
            CodTipom = "XX1"
        
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(CodTipom) Then
                letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "codtipom", "codtipom", CodTipom, "T")
        
                NumFact = vTipoMov.ConseguirContador(CodTipom)
                
                If DBLet(Rsx!Importe, "N") > 0 Then ' se insertara en la cartera de pagos (spagop)
                    CadValues2 = ""
            
                    CadValuesAux2 = "("
                    If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & "'" & SerieFraPro & "',"
                    CadValuesAux2 = CadValuesAux2 & "'" & Trim(Codmacta) & "', " & DBSet(NumFact, "T") & ", '" & Format(txtcodigo(11).Text, FormatoFecha) & "', "
            
                    '------------------------------------------------------------
                    i = 1
                    CadValues2 = CadValuesAux2 & i
                    
                    CadValues2 = CadValues2 & ", " & DBSet(txtcodigo(46), "N") & ", '" & Format(txtcodigo(13).Text, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & DBSet(Rsx!Importe, "N") & ", " & DBSet(txtcodigo(12).Text, "T") & ","
                
                    If Not vParamAplic.ContabilidadNueva Then
                        'David. Para que ponga la cuenta bancaria (SI LA tiene)
                        CadValues2 = CadValues2 & DBSet(CodBanco, "T", "S") & "," & DBSet(CodSucur, "T", "S") & ","
                        CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CuentaBa, "T", "S") & ","
                    End If
                    
                    'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                    SQL = "Fact.: " & SerieFraPro & "-" & NumFact & "-" & Format(txtcodigo(11).Text, "dd/mm/yyyy")
                        
                    CadValues2 = CadValues2 & "'" & DevNombreSQL(SQL) & "',"
                    
                    SQL = ""
                    CadValues2 = CadValues2 & "'" & DevNombreSQL(SQL) & "'" ')"
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(Iban, "") & MiFormat(CStr(CodBanco), "0000") & MiFormat(CStr(CodSucur), "0000") & MiFormat(CC, "00") & MiFormat(CuentaBa, "0000000000")
                        
                        CadValues2 = CadValues2 & ", " & DBSet(vvIban, "T") & ","
                        'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                        CadValues2 = CadValues2 & DBSet(Nombre, "T") & "," & DBSet(Direccion, "T") & "," & DBSet(Poblacion, "T") & "," & DBSet(CodPobla, "T") & ","
                        CadValues2 = CadValues2 & DBSet(Provincia, "T") & "," & DBSet(nif, "T") & ",'ES') "
                    
                    Else
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & ") "
                        Else
                            CadValues2 = CadValues2 & ") "
                        End If
                    End If
                
                    'Grabar tabla spagop de la CONTABILIDAD
                    '-------------------------------------------------
                    If CadValues2 <> "" Then
                        'Insertamos en la tabla spagop de la CONTA
                        'David. Cuenta bancaria y descripcion textos
                        If vParamAplic.ContabilidadNueva Then
                            SQL = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                            SQL = SQL & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
                        Else
                            SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                SQL = SQL & ", iban) "
                            Else
                                SQL = SQL & ") "
                            End If
                        End If
                        SQL = SQL & " VALUES " & CadValues2
                        ConnConta.Execute SQL
                    End If
                
                Else
                    ' si es negativo se inserta en positivo en la cartera de cobros (scobro)
            
                    letraser = ""
                    letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(CodTipom, "T"))
            
            '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
            '        Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(numfactu, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
                    Text33csb = "'Factura:" & DBLet(NumFact, "T") & " " & Format(DBLet(txtcodigo(11).Text, "F"), "dd/mm/yy") & "'"
                    Text41csb = "de " & DBSet(Rsx!Importe, "N")
                    Text42csb = ""
            
                        
                    '[Monica]03/07/2013: a�ado trim(codmacta)
                    CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(NumFact, "N") & "," & DBSet(txtcodigo(11).Text, "F") & ", 1," & DBSet(Trim(codMacta2), "T") & ","
                    CadValues2 = CadValuesAux2 & DBSet(txtcodigo(12).Text, "N") & "," & DBSet(txtcodigo(13).Text, "F") & "," & DBSet(Rsx!Importe * (-1), "N") & ","
                    If Not vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(txtcodigo(12).Text, "T") & "," & DBSet(CodBanco, "N", "S") & "," & DBSet(CodSucur, "N", "S") & ","
                        CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CuentaBa, "T", "S") & ","
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" ')"
                    Else
                        CadValues2 = CadValues2 & DBSet(txtcodigo(12).Text, "T") & "," & ValorNulo & "," & ValorNulo & ","
                        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1"  ')"
                    End If
                    
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(Iban, "") & MiFormat(CStr(CodBanco), "0000") & MiFormat(CStr(CodSucur), "0000") & MiFormat(CC, "00") & MiFormat(CuentaBa, "0000000000")
                        
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                        'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                        CadValues2 = CadValues2 & DBSet(Nombre, "T") & "," & DBSet(Direccion, "T") & "," & DBSet(Poblacion, "T") & "," & DBSet(CodPobla, "T") & ","
                        CadValues2 = CadValues2 & DBSet(Provincia, "T") & "," & DBSet(nif, "T") & ",'ES') "
            
                        'Insertamos en la tabla scobro de la CONTA
                        SQL = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        SQL = SQL & "ctabanc1,  fecultco, impcobro, "
                        SQL = SQL & " text33csb, text41csb,  agente, iban, " ') "
                        SQL = SQL & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
                        SQL = SQL & ") "
                    
                    Else
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & "," & DBSet(Iban, "T", "S") & ") "
                        Else
                            CadValues2 = CadValues2 & ") "
                        End If
                        
                
                        'Insertamos en la tabla scobro de la CONTA
                        SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        SQL = SQL & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                        SQL = SQL & " text33csb, text41csb, text42csb, agente" ') "
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            SQL = SQL & ", iban) "
                        Else
                            SQL = SQL & ") "
                        End If
                    End If
                    
                    SQL = SQL & " VALUES " & CadValues2
                    ConnConta.Execute SQL
            
                End If
                
            End If
            
            vTipoMov.IncrementarContador (CodTipom)
            Set vTipoMov = Nothing
        
            
            
            B = True
        
        End If
    End If
                
                
'hasta aqui a�adido
' He quitado lo siguiente
        
'        Set vSoc = New cSocio
'
'        If vSoc.LeerDatos(Rsx!CODSOCCLI) Then
'            CodTipom = "XX1"
'
'            Set vTipoMov = New CTiposMov
'            If vTipoMov.Leer(CodTipom) Then
'                letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "codtipom", "codtipom", CodTipom, "T")
'
'                NumFact = vTipoMov.ConseguirContador(CodTipom)
'
'                If vSoc.LeerDatosSeccion(DBLet(Rsx!CODSOCCLI, "N"), txtCodigo(52).Text) Then
'                    b = True
'
'                    CC = DBLet(vSoc.Digcontrol, "T")
'                    If DBLet(vSoc.Digcontrol, "T") = "**" Then CC = "00"
'
'                    Codmacta = vSoc.CtaProv
'                End If
'
'                If DBLet(Rsx!Importe, "N") > 0 Then ' se insertara en la cartera de pagos (spagop)
'                    CadValues2 = ""
'
'                    CadValuesAux2 = "("
'                    If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & "'" & SerieFraPro & "',"
'                    CadValuesAux2 = CadValuesAux2 & "'" & Trim(vSoc.CtaProv) & "', " & DBSet(NumFact, "T") & ", '" & Format(txtCodigo(11).Text, FormatoFecha) & "', "
'
'                    '------------------------------------------------------------
'                    i = 1
'                    CadValues2 = CadValuesAux2 & i
'
'                    CadValues2 = CadValues2 & ", " & DBSet(txtCodigo(46), "N") & ", '" & Format(txtCodigo(13).Text, FormatoFecha) & "', "
'                    CadValues2 = CadValues2 & DBSet(Rsx!Importe, "N") & ", " & DBSet(txtCodigo(12).Text, "T") & ","
'
'                    If Not vParamAplic.ContabilidadNueva Then
'                        'David. Para que ponga la cuenta bancaria (SI LA tiene)
'                        CadValues2 = CadValues2 & DBSet(vSoc.Banco, "T", "S") & "," & DBSet(vSoc.Sucursal, "T", "S") & ","
'                        CadValues2 = CadValues2 & DBSet(vSoc.Digcontrol, "T", "S") & "," & DBSet(vSoc.CuentaBan, "T", "S") & ","
'                    End If
'
'                    'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
'                    Sql = "Fact.: " & SerieFraPro & "-" & NumFact & "-" & Format(txtCodigo(11).Text, "dd/mm/yyyy")
'
'                    CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
'
'                    Sql = ""
'                    CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
'                    If vParamAplic.ContabilidadNueva Then
'                        vvIban = MiFormat(vSoc.Iban, "") & MiFormat(CStr(vSoc.Banco), "0000") & MiFormat(CStr(vSoc.Sucursal), "0000") & MiFormat(vSoc.Digcontrol, "00") & MiFormat(vSoc.CuentaBan, "0000000000")
'
'                        CadValues2 = CadValues2 & ", " & DBSet(vvIban, "T") & ","
'                        'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
'                        CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
'                        CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES') "
'
'                    Else
'                        '[Monica]22/11/2013: Tema iban
'                        If vEmpresa.HayNorma19_34Nueva = 1 Then
'                            CadValues2 = CadValues2 & ", " & DBSet(vSoc.Iban, "T", "S") & ") "
'                        Else
'                            CadValues2 = CadValues2 & ") "
'                        End If
'                    End If
'
'                    'Grabar tabla spagop de la CONTABILIDAD
'                    '-------------------------------------------------
'                    If CadValues2 <> "" Then
'                        'Insertamos en la tabla spagop de la CONTA
'                        'David. Cuenta bancaria y descripcion textos
'                        If vParamAplic.ContabilidadNueva Then
'                            Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
'                            Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
'                        Else
'                            Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
'                            '[Monica]22/11/2013: Tema iban
'                            If vEmpresa.HayNorma19_34Nueva = 1 Then
'                                Sql = Sql & ", iban) "
'                            Else
'                                Sql = Sql & ") "
'                            End If
'                        End If
'                        Sql = Sql & " VALUES " & CadValues2
'                        ConnConta.Execute Sql
'                    End If
'
'                Else
'                    ' si es negativo se inserta en positivo en la cartera de cobros (scobro)
'
'                    letraser = ""
'                    letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(CodTipom, "T"))
'
'            '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
'            '        Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(numfactu, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
'                    Text33csb = "'Factura:" & DBLet(NumFact, "T") & " " & Format(DBLet(txtCodigo(11).Text, "F"), "dd/mm/yy") & "'"
'                    Text41csb = "de " & DBSet(Rsx!Importe, "N")
'                    Text42csb = ""
'
'                    CC = DBLet(vSoc.Digcontrol, "T")
'                    If DBLet(vSoc.Digcontrol, "T") = "**" Then CC = "00"
'
'                    '[Monica]03/07/2013: a�ado trim(codmacta)
'                    CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(NumFact, "N") & "," & DBSet(txtCodigo(11).Text, "F") & ", 1," & DBSet(Trim(vSoc.CtaProv), "T") & ","
'                    CadValues2 = CadValuesAux2 & DBSet(txtCodigo(12).Text, "N") & "," & DBSet(txtCodigo(13).Text, "F") & "," & DBSet(Rsx!Importe * (-1), "N") & ","
'                    If Not vParamAplic.ContabilidadNueva Then
'                        CadValues2 = CadValues2 & DBSet(txtCodigo(12).Text, "T") & "," & DBSet(vSoc.Banco, "N", "S") & "," & DBSet(vSoc.Sucursal, "N", "S") & ","
'                        CadValues2 = CadValues2 & DBSet(vSoc.Digcontrol, "T", "S") & "," & DBSet(vSoc.CuentaBan, "T", "S") & ","
'                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
'                        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" ')"
'                    Else
'                        CadValues2 = CadValues2 & DBSet(txtCodigo(12).Text, "T") & "," & ValorNulo & "," & ValorNulo & ","
'                        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1"  ')"
'                    End If
'
'                    If vParamAplic.ContabilidadNueva Then
'                        vvIban = MiFormat(vSoc.Iban, "") & MiFormat(CStr(vSoc.Banco), "0000") & MiFormat(CStr(vSoc.Sucursal), "0000") & MiFormat(vSoc.Digcontrol, "00") & MiFormat(vSoc.CuentaBan, "0000000000")
'
'                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
'                        'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
'                        CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
'                        CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES') "
'
'                        'Insertamos en la tabla scobro de la CONTA
'                        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
'                        Sql = Sql & "ctabanc1,  fecultco, impcobro, "
'                        Sql = Sql & " text33csb, text41csb,  agente, iban, " ') "
'                        Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
'                        Sql = Sql & ") "
'
'                    Else
'                        '[Monica]22/11/2013: Tema iban
'                        If vEmpresa.HayNorma19_34Nueva = 1 Then
'                            CadValues2 = CadValues2 & "," & DBSet(vSoc.Iban, "T", "S") & ") "
'                        Else
'                            CadValues2 = CadValues2 & ") "
'                        End If
'
'
'                        'Insertamos en la tabla scobro de la CONTA
'                        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
'                        Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
'                        Sql = Sql & " text33csb, text41csb, text42csb, agente" ') "
'                        '[Monica]22/11/2013: Tema iban
'                        If vEmpresa.HayNorma19_34Nueva = 1 Then
'                            Sql = Sql & ", iban) "
'                        Else
'                            Sql = Sql & ") "
'                        End If
'                    End If
'
'                    Sql = Sql & " VALUES " & CadValues2
'                    ConnConta.Execute Sql
'
'                End If
'
'            End If
'            b = True
'        End If
'    End If
    
    Set vSoc = Nothing
        
    
EInsertarTesoreriaNewFac:
    If Err.Number <> 0 Then
        B = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewFVARPagos = B
End Function



