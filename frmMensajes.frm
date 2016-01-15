VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14160
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameImpFrasPozos 
      Height          =   5790
      Left            =   0
      TabIndex        =   135
      Top             =   0
      Width           =   10260
      Begin VB.CommandButton CmdAcepImpFras 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   7080
         TabIndex        =   138
         Top             =   5130
         Width           =   1215
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   4
         Left            =   8520
         TabIndex        =   137
         Top             =   5130
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView19 
         Height          =   4155
         Left            =   240
         TabIndex        =   136
         Top             =   750
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label17 
         Caption         =   "Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   139
         Top             =   270
         Width           =   5145
      End
      Begin VB.Image imgCheck4 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmMensajes.frx":000C
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck4 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmMensajes.frx":0156
         Top             =   5160
         Width           =   240
      End
   End
   Begin VB.Frame FrameVisualizaEntradas 
      Height          =   5790
      Left            =   0
      TabIndex        =   160
      Top             =   0
      Width           =   10260
      Begin VB.CommandButton cmdcancelVisEnt 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   8520
         TabIndex        =   162
         Top             =   5130
         Width           =   1215
      End
      Begin VB.CommandButton CmdAcepVisEntr 
         Caption         =   "Continuar"
         Height          =   375
         Left            =   7080
         TabIndex        =   161
         Top             =   5130
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView20 
         Height          =   4155
         Left            =   240
         TabIndex        =   163
         Top             =   750
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label22 
         Caption         =   "Datos del fichero de Entradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   164
         Top             =   270
         Width           =   5145
      End
   End
   Begin VB.Frame FrameCreacionCampo 
      Height          =   4725
      Left            =   0
      TabIndex        =   140
      Top             =   0
      Width           =   6555
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   145
         Tag             =   "Poligono|N|N|0|999|rcampos|poligono|000||"
         Top             =   1710
         Width           =   825
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   146
         Tag             =   "Parcela|N|N|0|999999|rcampos|parcela|000000||"
         Top             =   2130
         Width           =   795
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   147
         Tag             =   "Subparcela|T|S|||rcampos|subparce|||"
         Top             =   2550
         Width           =   825
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   148
         Tag             =   "Partida|N|N|1|9999|rcampos|codparti|0000||"
         Top             =   3210
         Width           =   855
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2295
         MaxLength       =   30
         TabIndex        =   155
         Top             =   3210
         Width           =   3795
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   144
         Tag             =   "Variedad|N|N|1|9999|rcampos|codvarie|0000||"
         Top             =   1290
         Width           =   840
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2220
         MaxLength       =   30
         TabIndex        =   153
         Top             =   1320
         Width           =   3915
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   149
         Top             =   870
         Width           =   3915
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   143
         Tag             =   "Código Socio|N|N|1|999999|rcampos|codsocio|000000|N|"
         Top             =   870
         Width           =   825
      End
      Begin VB.CommandButton CmdCanCrearCampo 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   152
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepCrearCampo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   150
         Top             =   3780
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1050
         ToolTipText     =   "Buscar Socio"
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1050
         ToolTipText     =   "Buscar Variedad"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1050
         ToolTipText     =   "Buscar Partida"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Poligono"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   159
         Top             =   1740
         Width           =   705
      End
      Begin VB.Label Label20 
         Caption         =   "Parcela"
         Height          =   255
         Left            =   390
         TabIndex        =   158
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label Label19 
         Caption         =   "Subparcela"
         Height          =   255
         Left            =   390
         TabIndex        =   157
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label Label21 
         Caption         =   "Partida"
         Height          =   255
         Left            =   390
         TabIndex        =   156
         Top             =   3210
         Width           =   585
      End
      Begin VB.Label Label18 
         Caption         =   "Variedad"
         Height          =   255
         Left            =   390
         TabIndex        =   154
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Socio"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   151
         Top             =   900
         Width           =   600
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   142
         Top             =   3480
         Width           =   6195
      End
      Begin VB.Label Label1 
         Caption         =   "Datos para la creación del campo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   13
         Left            =   390
         TabIndex        =   141
         Top             =   330
         Width           =   5565
      End
   End
   Begin VB.Frame FramePago 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2235
      Left            =   0
      TabIndex        =   122
      Top             =   0
      Width           =   3945
      Begin VB.CheckBox Check1 
         Caption         =   "Contado"
         Height          =   225
         Left            =   570
         TabIndex        =   126
         Top             =   960
         Width           =   1485
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Banco"
         Height          =   225
         Left            =   2190
         TabIndex        =   125
         Top             =   960
         Width           =   1485
      End
      Begin VB.CommandButton CmdAcep 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1140
         TabIndex        =   124
         Top             =   1440
         Width           =   1005
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   2
         Left            =   2460
         TabIndex        =   123
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label15 
         Caption         =   "Tipo de Ticket"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   270
         TabIndex        =   127
         Top             =   330
         Width           =   3015
      End
   End
   Begin VB.Frame FrameMatriculas 
      Height          =   4620
      Left            =   0
      TabIndex        =   131
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Regresar"
         Height          =   375
         Index           =   3
         Left            =   4650
         TabIndex        =   132
         Top             =   4110
         Width           =   1005
      End
      Begin MSComctlLib.ListView ListView18 
         Height          =   3255
         Left            =   240
         TabIndex        =   133
         Top             =   660
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.Label Label4 
         Caption         =   "Matrículas del Transportista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   4
         Left            =   270
         TabIndex        =   134
         Top             =   240
         Width           =   5145
      End
   End
   Begin VB.Frame FrameEntradasConError 
      Height          =   4620
      Left            =   0
      TabIndex        =   84
      Top             =   60
      Width           =   8655
      Begin VB.CommandButton CmdSal 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7290
         TabIndex        =   85
         Top             =   4110
         Width           =   1005
      End
      Begin MSComctlLib.ListView ListView15 
         Height          =   3255
         Left            =   240
         TabIndex        =   86
         Top             =   660
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.Label Label4 
         Caption         =   "Entradas con error"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   87
         Top             =   210
         Width           =   5145
      End
   End
   Begin VB.Frame frameClaveAcceso 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1365
      Left            =   0
      TabIndex        =   128
      Top             =   0
      Width           =   3645
      Begin VB.TextBox Text7 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1290
         PasswordChar    =   "*"
         TabIndex        =   129
         Top             =   570
         Width           =   1665
      End
      Begin VB.Label Label16 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   130
         Top             =   600
         Width           =   945
      End
   End
   Begin VB.Frame FrameVariedades 
      Height          =   5790
      Left            =   0
      TabIndex        =   32
      Top             =   60
      Width           =   7050
      Begin MSComctlLib.ListView ListView6 
         Height          =   4155
         Left            =   225
         TabIndex        =   35
         Top             =   675
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.CommandButton cmdCanVariedades 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5520
         TabIndex        =   34
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdAcepVariedades 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   33
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   240
         Picture         =   "frmMensajes.frx":02A0
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   600
         Picture         =   "frmMensajes.frx":03EA
         Top             =   5160
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Variedades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   36
         Top             =   270
         Width           =   5145
      End
   End
   Begin VB.Frame FrameHidrantesSocio 
      Height          =   5790
      Left            =   2970
      TabIndex        =   67
      Top             =   90
      Width           =   7050
      Begin VB.CommandButton CmdAceptarPozos 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   70
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton CmdCan 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5520
         TabIndex        =   69
         Top             =   5160
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView13 
         Height          =   4155
         Left            =   210
         TabIndex        =   68
         Top             =   750
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "Contadores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   71
         Top             =   270
         Width           =   6495
      End
      Begin VB.Image imgCheck2 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmMensajes.frx":0534
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck2 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmMensajes.frx":067E
         Top             =   5160
         Width           =   240
      End
   End
   Begin VB.Frame FrameCambios 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   6000
      Left            =   0
      TabIndex        =   113
      Top             =   -60
      Width           =   8145
      Begin VB.TextBox txtAux 
         Height          =   585
         Index           =   0
         Left            =   480
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   121
         Tag             =   "Valor Anterior|T|S|||cambios|valoranterior|||"
         Top             =   4650
         Width           =   7365
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Regresar"
         Height          =   405
         Index           =   1
         Left            =   6540
         TabIndex        =   119
         Top             =   5310
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Height          =   1365
         Index           =   3
         Left            =   450
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   115
         Tag             =   "Cadena|T|N|||cambios|cadena|||"
         Top             =   1170
         Width           =   7365
      End
      Begin VB.TextBox txtAux 
         Height          =   1365
         Index           =   4
         Left            =   450
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   114
         Tag             =   "Valor Anterior|T|S|||cambios|valoranterior|||"
         Top             =   2910
         Width           =   7365
      End
      Begin VB.Label Label4 
         Caption         =   "CP"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   120
         Top             =   4410
         Width           =   465
      End
      Begin VB.Label Label14 
         Caption         =   "Datos Cambio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   450
         TabIndex        =   118
         Top             =   360
         Width           =   5145
      End
      Begin VB.Label Label1 
         Caption         =   "Cadena ejecutada"
         Height          =   255
         Index           =   12
         Left            =   450
         TabIndex        =   117
         Top             =   870
         Width           =   2115
      End
      Begin VB.Label Label4 
         Caption         =   "Valor Anterior"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   116
         Top             =   2640
         Width           =   1215
      End
   End
   Begin VB.Frame FrameCampos 
      Height          =   5610
      Left            =   0
      TabIndex        =   28
      Top             =   60
      Width           =   8935
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1140
         TabIndex        =   56
         Text            =   "Text4"
         Top             =   5010
         Visible         =   0   'False
         Width           =   1275
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   4155
         Left            =   225
         TabIndex        =   29
         Top             =   720
         Width           =   8455
         _ExtentX        =   14923
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Partida"
            Object.Width           =   3000
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "C.Pobla"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Poblacion"
            Object.Width           =   2735
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Polígono"
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Parcela"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Sp."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Nro."
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Hdas"
            Object.Width           =   1305
         EndProperty
      End
      Begin VB.CommandButton cmdCamposSocio 
         Caption         =   "Regresar"
         Height          =   375
         Index           =   2
         Left            =   7020
         TabIndex        =   30
         Top             =   4995
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Socio"
         Height          =   285
         Left            =   360
         TabIndex        =   55
         Top             =   5040
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Campos del Socio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   31
         Top             =   270
         Width           =   5145
      End
   End
   Begin VB.Frame FrameCamposSocio 
      Height          =   7455
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton cmdCamposSocio 
         Caption         =   "Imprimir"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   20
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCamposSocio 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6960
         TabIndex        =   19
         Top             =   6960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   6045
         Left            =   240
         TabIndex        =   18
         Top             =   810
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   10663
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Partida"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Polígono"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Parcela"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Nro."
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Campos del Socio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   27
         Top             =   270
         Width           =   5145
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmMensajes.frx":07C8
         Top             =   6960
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmMensajes.frx":0912
         Top             =   6960
         Width           =   240
      End
   End
   Begin VB.Frame FrameConsumoSocio 
      Height          =   4500
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   9855
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6420
         TabIndex        =   103
         Text            =   "Text2"
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2970
         TabIndex        =   101
         Text            =   "Text2"
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6420
         TabIndex        =   49
         Text            =   "Text2"
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2970
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton CmdSalir1 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   8460
         TabIndex        =   43
         Top             =   4020
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView8 
         Height          =   3055
         Left            =   150
         TabIndex        =   42
         Top             =   540
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   5398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin MSComctlLib.ListView ListView9 
         Height          =   3055
         Left            =   4950
         TabIndex        =   45
         Top             =   540
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   5398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.Label Label1 
         Caption         =   "Bodega:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   11
         Left            =   4950
         TabIndex        =   104
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Recolectados Almazara:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   10
         Left            =   210
         TabIndex        =   102
         Top             =   4080
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "Disponible:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   8
         Left            =   4950
         TabIndex        =   50
         Top             =   3720
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "Litros producidos de Almazara:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   7
         Left            =   210
         TabIndex        =   47
         Top             =   3720
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "Consumo Socio por Productos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   6
         Left            =   4950
         TabIndex        =   46
         Top             =   210
         Width           =   4065
      End
      Begin VB.Label Label1 
         Caption         =   "Consumo Socio por Variedades:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   4
         Left            =   150
         TabIndex        =   44
         Top             =   210
         Width           =   4065
      End
   End
   Begin VB.Frame FrameEntradasSinClasificar 
      Height          =   4620
      Left            =   0
      TabIndex        =   21
      Top             =   90
      Width           =   8655
      Begin VB.CommandButton CmdAceptarPal 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5910
         TabIndex        =   23
         Top             =   4005
         Width           =   975
      End
      Begin VB.CommandButton CmdCanPal 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7080
         TabIndex        =   22
         Top             =   4005
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   3135
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.Label Label1 
         Caption         =   "¿ Desea Continuar ?"
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   26
         Top             =   4050
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "Entradas sin clasificar o sin gastos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   3
         Left            =   135
         TabIndex        =   25
         Top             =   225
         Width           =   7215
      End
   End
   Begin VB.Frame FrameFacturas 
      Height          =   5610
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   6585
      Begin VB.CommandButton cmdFacturas 
         Caption         =   "Regresar"
         Height          =   375
         Left            =   4890
         TabIndex        =   59
         Top             =   4980
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView11 
         Height          =   4155
         Left            =   225
         TabIndex        =   58
         Top             =   720
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Partida"
            Object.Width           =   3000
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "C.Pobla"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Poblacion"
            Object.Width           =   2823
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Polígono"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Parcela"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Nro."
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Hdas"
            Object.Width           =   1305
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Anticipos de Venta Campo sin Entradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   60
         Top             =   270
         Width           =   5145
      End
   End
   Begin VB.Frame FrameHidrantesANoFacturar 
      Height          =   4500
      Left            =   0
      TabIndex        =   106
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton CmdSalir4 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   8490
         TabIndex        =   112
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton CmdContinuar 
         Caption         =   "&Continuar"
         Height          =   375
         Left            =   7380
         TabIndex        =   107
         Top             =   3840
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView16 
         Height          =   3055
         Left            =   150
         TabIndex        =   108
         Top             =   540
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   5398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin MSComctlLib.ListView ListView17 
         Height          =   3055
         Left            =   4950
         TabIndex        =   109
         Top             =   540
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   5398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.Label Label1 
         Caption         =   "con Consumo inferior al mínimo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   17
         Left            =   150
         TabIndex        =   111
         Top             =   210
         Width           =   4065
      End
      Begin VB.Label Label1 
         Caption         =   "con Consumo superior al máximo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   16
         Left            =   4950
         TabIndex        =   110
         Top             =   210
         Width           =   4065
      End
   End
   Begin VB.Frame FrameCobrosPtes 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton cmdCancelarCobros 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7080
         TabIndex        =   11
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtParam 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmMensajes.frx":0A5C
         Top             =   210
         Width           =   6615
      End
      Begin VB.CommandButton cmdAceptarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5970
         TabIndex        =   1
         Top             =   4440
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.Label Label1 
         Caption         =   "¿Desea continuar?"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   12
         Top             =   4440
         Width           =   7215
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame FrameOrdenListado 
      Height          =   2265
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   5865
      Begin VB.CommandButton CmdAcepOrden 
         Caption         =   "&Regresar"
         Height          =   375
         Left            =   4440
         TabIndex        =   77
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Frame Frame3 
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
         Height          =   825
         Left            =   150
         TabIndex        =   73
         Top             =   630
         Width           =   5385
         Begin VB.OptionButton Option2 
            Caption         =   "Contador"
            Height          =   225
            Index           =   0
            Left            =   390
            TabIndex        =   76
            Top             =   330
            Width           =   1545
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Código Socio"
            Height          =   225
            Index           =   1
            Left            =   2100
            TabIndex        =   75
            Top             =   330
            Width           =   1545
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Nombre Socio"
            Height          =   225
            Index           =   2
            Left            =   3750
            TabIndex        =   74
            Top             =   330
            Width           =   1545
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Orden del Listado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   180
         TabIndex        =   78
         Top             =   300
         Width           =   3495
      End
   End
   Begin VB.Frame FrameImgContador 
      Height          =   9720
      Left            =   0
      TabIndex        =   105
      Top             =   0
      Width           =   8325
      Begin VB.Image Image2 
         Height          =   9585
         Left            =   0
         Stretch         =   -1  'True
         Top             =   90
         Width           =   8280
      End
   End
   Begin VB.Frame FrameDiferencias 
      Height          =   1890
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Width           =   6915
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   495
         Left            =   240
         TabIndex        =   100
         Top             =   330
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton CmdCanDif 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5460
         TabIndex        =   98
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   285
         Left            =   270
         TabIndex        =   99
         Top             =   990
         Width           =   4155
      End
   End
   Begin VB.Frame FrameEmpresas 
      Height          =   5610
      Left            =   0
      TabIndex        =   88
      Top             =   0
      Width           =   6915
      Begin VB.CommandButton CmdAcepEmpresas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3930
         TabIndex        =   95
         Top             =   4980
         Width           =   1215
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1050
         TabIndex        =   91
         Top             =   660
         Width           =   1440
      End
      Begin VB.TextBox txtlargo 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2610
         TabIndex        =   90
         Top             =   660
         Width           =   3945
      End
      Begin VB.CommandButton CmdSalir3 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5340
         TabIndex        =   89
         Top             =   4980
         Width           =   1215
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   3600
         Left            =   210
         TabIndex        =   96
         Top             =   1170
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   6350
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4885
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   3381
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   210
         Picture         =   "frmMensajes.frx":0A62
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblLabels 
         Caption         =   "Empresas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   94
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Seleccione una de las empresas disponibles para el usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   930
         TabIndex        =   93
         Top             =   300
         Width           =   5595
      End
      Begin VB.Label lblLabels 
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   92
         Top             =   660
         Width           =   720
      End
   End
   Begin VB.Frame FrameAlbaranesLiquidados 
      Height          =   5610
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   6585
      Begin VB.CommandButton CmdAlbLiq 
         Caption         =   "Liquidar Todos"
         Height          =   375
         Index           =   0
         Left            =   1890
         TabIndex        =   66
         Top             =   4980
         Width           =   1425
      End
      Begin VB.CommandButton CmdAlbLiq 
         Caption         =   "Sólo Pendientes"
         Height          =   375
         Index           =   1
         Left            =   3450
         TabIndex        =   65
         Top             =   4980
         Width           =   1335
      End
      Begin VB.CommandButton CmdAlbLiq 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4890
         TabIndex        =   62
         Top             =   4980
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView12 
         Height          =   4155
         Left            =   225
         TabIndex        =   63
         Top             =   720
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Partida"
            Object.Width           =   3000
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "C.Pobla"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Poblacion"
            Object.Width           =   2823
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Polígono"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Parcela"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Nro."
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Hdas"
            Object.Width           =   1305
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Albaranes Liquidados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   64
         Top             =   270
         Width           =   5145
      End
   End
   Begin VB.Frame FrameArchivos 
      Height          =   5790
      Left            =   3000
      TabIndex        =   79
      Top             =   0
      Width           =   7050
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   81
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton CmdAcepArchivos 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   80
         Top             =   5160
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView14 
         Height          =   4155
         Left            =   210
         TabIndex        =   82
         Top             =   750
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Image imgCheck3 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmMensajes.frx":0EA4
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck3 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmMensajes.frx":0FEE
         Top             =   5160
         Width           =   240
      End
      Begin VB.Label Label12 
         Caption         =   "Archivos a adjuntar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   83
         Top             =   270
         Width           =   6495
      End
   End
   Begin VB.Frame FrameNSeries 
      Height          =   5000
      Left            =   0
      TabIndex        =   3
      Top             =   -45
      Width           =   6975
      Begin VB.CommandButton cmdSelTodos 
         Caption         =   "&Todos"
         Height          =   315
         Left            =   720
         TabIndex        =   8
         Top             =   4320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdDeselTodos 
         Caption         =   "&Ninguno"
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   4320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNSeries 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   4320
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   720
         TabIndex        =   5
         Top             =   840
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.Image imgCheck1 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmMensajes.frx":1138
         Top             =   4185
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgCheck1 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmMensajes.frx":1282
         Top             =   4185
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Frame FrameErrores 
      Height          =   5535
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   7080
         TabIndex        =   15
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   4335
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Text            =   "frmMensajes.frx":13CC
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame FrameEntradasSinCRFID 
      Height          =   4620
      Left            =   30
      TabIndex        =   37
      Top             =   210
      Width           =   8655
      Begin MSComctlLib.ListView ListView7 
         Height          =   3135
         Left            =   240
         TabIndex        =   39
         Top             =   540
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7230
         TabIndex        =   38
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Entradas sin CRFID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   5
         Left            =   270
         TabIndex        =   40
         Top             =   210
         Width           =   7215
      End
   End
   Begin VB.Frame FrameEntradasSinSalida 
      Height          =   4620
      Left            =   2610
      TabIndex        =   51
      Top             =   990
      Width           =   8655
      Begin VB.CommandButton CmdSalir2 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7290
         TabIndex        =   53
         Top             =   4110
         Width           =   1005
      End
      Begin MSComctlLib.ListView ListView10 
         Height          =   3255
         Left            =   240
         TabIndex        =   52
         Top             =   660
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin VB.Label Label4 
         Caption         =   "Entradas sin Tarar Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Top             =   210
         Width           =   5145
      End
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================== VBLES PUBLICAS ================================

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionMensaje As Byte
'======================================
'==== FACTURACION =====================
' 1 .- Mensaje de Cobros Pendientes
' 2 .- Mensaje de No hay suficiente Stock para pasar de Pedido a Albaran
' 3 .- Mensaje Acerca de...
' 4 .- Variedades dado un rango de productos
' 5 .- Consumo de socio por variedades y por producto (bodega)
' 6 .- Mostrar Campos de socio dados de alta
' 7 .- Mostrar Campos de socio dados de alta con la variedad (ADV)
' 8 .- Mostrar lista clientes para seleccionar los que queremos imprimir (Etiquetas)
' 9 .- Mostrar lista Proveedores para seleccionar los que queremos imprimir (Etiquetas)
'10 .- Mostrar lista de Errores de las facturas NO contabilizadas
'11 .- Mostrar lista lineas de factura a Rectificar para seleccionar las q queremos traer al Albaran de FAct. Rectificativa
'12 .- Mostrar Albaranes del Rango que no se van a Facturar. (Facturar Albaranes Venta)

'13 .- Mostrar Errores
'14 .- Mostrar Empresas existentes en el sistema



'15 .- Mostrar lista de campos del el socio
'16 .- Mostrar las variedades dado un rango de clases
'17 .- Etiquetas clientes. LO MISMO QUE EL 8 pero hecho por david
'18 .- Mantenimientos. paso ejercicio siguiente a actual
'19 .- Lista de Entradas que estan sin clasificar o sin gastos
'20 .- Entradas que no se han clasificado por variedad sin calidad de VC

'21 .- Entradas de bascula que no tienen CRFID
'22 .- Trabajadores a seleccionar de la cuadrilla
'23 .- Hidrantes de un Socio (POZOS)

'24 .- Albaranes de un tranportista pendientes de facturar
'25 .- Campos con el nro de orden incorrecto, solo tiene sentido para Alzira

'26 .- albaranes de un socio
'27 .- Plagas
'28 .- Notas de campo sin tara de salida (SOLO PARA QUATRETONDA)


'29 .- Campos de un socio para POZOS

'30 .- Albaranes de bodega sin tarar
'31 .- Facturas de anticipo venta campo sin kilos entrados
'32 .- Tipo de Aportaciones

'34 .- Campos para el informe de clasificacion

'35 .- Albaranes Liquidados

'36 .- Situaciones de los socios para el informe de socios por seccion

'37 .- Hidrantes de un socio para hacer factura de mantenimiento

'38 .- Orden del printnou

'39 .- Hidrantes de un campo para hacer el cambio de socio (desde mto de campos/cambio de socio)

'40 .- Archivos a incluir en el email
'41 .- Entradas importadas desde excel con error (Traspaso de clasificacion de Anna)


'42 .- Socios de los recibos de pozos seleccionados
'43 .- Muestra la campaña anterior y la actual para sacar los kilos de oliva de Aportaciones (Solo Moixent)

'44 .- Busqueda de diferencias de indefa Pozos Escalona
'45 .- imagen ampliada del contador de Escalona/Utxera
'46 .- imagen ampliada del documento

'47 .- Entradas de hco. de fruta que no estan asignadas a ninguna factura
'48 .- Anticipos sin descontar

'49 .- Campos sin precio en la zona para facturar talla

'50 .- Hidrantes a no facturar por consumo inferior/superior al minimo/maximo. POZOS

'51 .- Campos Agrupados por nro de campo
'53 .- Datos Cambios registros del log


'54 .- Transportistas a facturar
'55 .- Socios de rsocios_pozos a seleccionar (CASTELDUC)

'56 .- Fechas de anticipos pendientes de descontar que quieren descontar

'57 .- Campos para facturar a manta


'58 .- tipo de forma de pago para los recibos de manta POZOS
'59 .- Clave de acceso (password)

'60 .- Transportistas con el mismo nif

'61 .- Facturas de pozos a impresión

'62 .- Creacion de un campo en el traspaso de entradas de almazara para ABN
'63 .- Visualizacion previa del fichero de entradas

Private WithEvents frmPar As frmManPartidas 'partidas
Attribute frmPar.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1


Public cadWhere As String 'Cadena para pasarle la WHERE de la SELECT de los cobros pendientes o de Pedido(para comp. stock)
                          'o CodArtic para seleccionar los Nº Series
                          'para cargar el ListView
                          
Public cadWHERE2 As String
Public Campo As String
Public Cadena As String ' sql para cargar el listview
Public vCampos As String 'Articulo y cantidad Empipados para Nº de Series
                         'Tambien para pasar el nombre de la tabla de lineas (sliped, slirep,...)
                         'Dependiendo desde donde llamemos, de Pedidos o Reparaciones


'====================== VBLES LOCALES ================================

Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form
Dim PrimeraVez As Boolean

'Para los Nº de Serie
Dim TotalArray As Integer
Dim codArtic() As String
Dim cantidad() As Integer


Dim CadContadores As String

Dim nomColumna As String
Dim nomColumna2 As String
Dim columna As Integer
Dim Columna2 As Integer
Dim Orden As Integer
Dim Orden2 As Integer
Dim PrimerCampo As Integer

Private Sub CmdAcepArchivos_Click()
    'Cargo lor registros marcados
    Cadena = ""
    For NumRegElim = 1 To ListView14.ListItems.Count
        If ListView14.ListItems(NumRegElim).Checked Then
            Cadena = Cadena & ListView14.ListItems(NumRegElim).Text & ","
        End If
    Next NumRegElim
    ' quitamos la ultima coma
    If Cadena <> "" Then
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
    End If
    
    RaiseEvent DatoSeleccionado(Cadena)
    Unload Me

End Sub

Private Function DatosOk() As Boolean

    DatosOk = False

    If Text8(1).Text = "" Or Text9(1).Text = "" Then
        MsgBox "Socio no existe. Reintroduzca.", vbExclamation
        Exit Function
    End If
    If Text8(2).Text = "" Or Text9(2).Text = "" Then
        MsgBox "Variedad no existe. Reintroduzca.", vbExclamation
        Exit Function
    End If
    If Text8(4).Text = "" Then
        MsgBox "Tiene que introducir un Polígono", vbExclamation
        Exit Function
    End If
    If Text8(5).Text = "" Then
        MsgBox "Tiene que introducir una Parcela", vbExclamation
        Exit Function
    End If
    If Text8(6).Text = "" Then
        MsgBox "Tiene que introducir una Subparcela", vbExclamation
        Exit Function
    End If

    DatosOk = True

End Function



Private Sub CmdAcepCrearCampo_Click()
Dim NroCampo As Long
Dim SQL As String
Dim Codzona As String
Dim vSuperficie As Currency

    On Error GoTo ECrear

    If Not DatosOk Then Exit Sub


    Codzona = DevuelveValor("select codzonas from rpartida where codparti = " & DBSet(Text8(3).Text, "N"))
    vSuperficie = 0

    SQL = "select max(codcampo) from rcampos "
    NroCampo = DevuelveValor(SQL) + 1

    ' insertamos en la tabla de rhisfruta
    SQL = "insert into rcampos (codcampo, codsocio, codpropiet, codvarie, codparti, "
    SQL = SQL & "codzonas, fecaltas, supsigpa, supcoope, supcatas, supculti, codsitua, "
    SQL = SQL & "poligono, parcela, subparce, asegurado, tipoparc, recintos, nrocampo, recolect) VALUES ("
    SQL = SQL & DBSet(NroCampo, "N") & ","
    SQL = SQL & DBSet(Text8(1).Text, "N") & ","
    SQL = SQL & DBSet(Text8(1).Text, "N") & ","
    SQL = SQL & DBSet(Text8(2).Text, "N") & ","
    SQL = SQL & DBSet(Text8(3).Text, "N") & ","
    SQL = SQL & DBSet(Codzona, "N") & ","
    SQL = SQL & DBSet(Now, "F") & ","
    SQL = SQL & DBSet(vSuperficie, "N") & "," ' superficie en hectareas
    SQL = SQL & DBSet(vSuperficie, "N") & ","
    SQL = SQL & DBSet(vSuperficie, "N") & ","
    SQL = SQL & DBSet(vSuperficie, "N") & ","
    SQL = SQL & "0," ' situacion
    SQL = SQL & DBSet(Text8(4).Text, "N") & ","
    SQL = SQL & DBSet(Text8(5).Text, "N") & ","
    SQL = SQL & DBSet(Text8(6).Text, "T") & ","
    SQL = SQL & "0,0,0,"
    SQL = SQL & DBSet(NroCampo, "N") & ","
    SQL = SQL & "0)"
    
    conn.Execute SQL
    
    RaiseEvent DatoSeleccionado(CStr(NroCampo))
    Unload Me
    
    Exit Sub
    
ECrear:
    MuestraError Err.Number, "Crear Campo", Err.Description
End Sub

Private Sub CmdAcepEmpresas_Click()
    CadenaDesdeOtroForm = lw1.SelectedItem.Tag
    
    Cadena = RecuperaValor(lw1.SelectedItem.Tag, 1)
    
    RaiseEvent DatoSeleccionado(Cadena)
    Unload Me

End Sub

Private Sub CmdAcepImpFras_Click()
Dim SQL As String
Dim I As Integer

    For I = 1 To Me.ListView19.ListItems.Count
        If ListView19.ListItems(I).Checked Then
            SQL = "update rrecibpozos set imprimir = " & DBSet(vUsu.PC, "T")
            SQL = SQL & " where codtipom = " & DBSet(ListView19.ListItems(I).Text, "T")
            SQL = SQL & " and numfactu = " & DBSet(Me.ListView19.ListItems(I).SubItems(1), "N")
            SQL = SQL & " and fecfactu = " & DBSet(Me.ListView19.ListItems(I).SubItems(2), "F")
            
            conn.Execute SQL
        End If
    Next I
    Unload Me
End Sub

Private Sub CmdAcepOrden_Click()
Dim devuelve As String

    devuelve = "pOrden="
    
    If Me.Option2(0).Value Then devuelve = devuelve & "{rpozos.hidrante}|"
    If Me.Option2(1).Value Then devuelve = devuelve & "{rpozos.codsocio}|"
    If Me.Option2(2).Value Then devuelve = devuelve & "{rsocios.nomsocio}|"
    
    RaiseEvent DatoSeleccionado(devuelve)
    Unload Me
End Sub

Private Sub CmdAceptarCobros_Click()
    If OpcionMensaje = 12 Then vCampos = "1"
    If OpcionMensaje = 49 Then vCampos = "1"
    
    Unload Me
End Sub


'Private Sub cmdAceptarComp_Click()
''Boton Aceptar de Componentes del Mant. de Nº de Series en Reparaciones
'Dim h As Integer, w As Integer
'
'    ponerFrameComponentesVisible False, h, w
'    PonerFrameCobrosPtesVisible True, h, w
'    Me.Height = h + 350
'    Me.Width = w + 70
'
'    If Me.OptCompXMant.Value Then
'        'Mostrar Resumen de los Nº de Serie del Mantenimiento
'        Me.Caption = "Equipos del Mantenimiento"
'        CargarListaComponentes (1)
'    ElseIf Me.OptCompXDpto.Value Then
'        'Mostrar Resumen de los Nº de Serie del Departamento
'        Me.Caption = "Equipos del Departamento"
'        CargarListaComponentes (2)
'    ElseIf Me.OptCompXClien.Value Then
'        'Mostrar Resumen de los Nº de Serie del Cliente
'        Me.Caption = "Equipos del Cliente"
'        CargarListaComponentes (3)
'    End If
'    PonerFocoBtn Me.cmdAceptarCobros
'End Sub


Private Sub CmdAceptarPal_Click()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cad As String

    SQL = "select cast(group_concat(numnotac) as char) from tmpclasifica where codusu = " & vUsu.Codigo
    SQL = SQL & " and codclase = 0"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Cad = DBLet(RS.Fields(0).Value, "T")
    Else
        Cad = ""
    End If
    Set RS = Nothing
    
    RaiseEvent DatoSeleccionado(Cad)
       
    Unload Me
End Sub

Private Sub cmdAceptarNSeries_Click()
Dim I As Integer, J As Integer
Dim Seleccionados As Integer
Dim Cad As String, SQL As String
Dim articulo As String
Dim RS As ADODB.Recordset
Dim c1 As String * 10, c2 As String * 10, c3 As String * 10


    If OpcionMensaje = 4 Then
        'Comprobar que se han seleccionado el nº correcto de  Nº de Serie para cada Articulo
        Seleccionados = 0
        articulo = ""
      
        'Si se ha seleccionado la cantidad correcta de Nº de series, empiparlos y
        'devolverlos al form de Albaranes(facturacion)
        Cad = ""
        For J = 0 To TotalArray
            articulo = codArtic(J)
            Cad = Cad & articulo & "|"
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    If articulo = ListView2.ListItems(I).ListSubItems(1).Text Then
                        If Seleccionados < Abs(cantidad(J)) Then
                            Seleccionados = Seleccionados + 1
                            Cad = Cad & ListView2.ListItems(I).Text & "|"
                        End If
                   'cad = cad & Data1.Recordset.Fields(1) & "|"
                    End If
                End If
            Next I
            If Seleccionados < Abs(cantidad(J)) Then
                'Comprobar que si tiene Nºs de serie de ese articulos cargados seleccione los
                'que corresponden
                SQL = "SELECT count(sserie.numserie)"
                SQL = SQL & " FROM sserie " 'INNER JOIN sartic ON sserie.codartic=sartic.codartic "
                SQL = SQL & " WHERE sserie.codartic=" & DBSet(articulo, "T")
                SQL = SQL & " AND (isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='') "
                SQL = SQL & " ORDER BY sserie.codartic, numserie "
                Set RS = New ADODB.Recordset
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If RS.Fields(0).Value >= Abs(cantidad(J)) - Seleccionados Then
                    MsgBox "Debe seleccionar " & cantidad(J) & " Nº Series para el articulo " & codArtic(J), vbExclamation
                    Exit Sub
                Else
                    'No hay Nº Serie y Pedirlos
                End If
                RS.Close
                Set RS = Nothing
            
            End If
            Cad = Cad & "·"
            Seleccionados = 0
        Next J
                                                                                                 '[Monica]11/11/2013: castelduc
    ElseIf OpcionMensaje = 8 Or OpcionMensaje = 9 Or OpcionMensaje = 17 Or OpcionMensaje = 42 Or OpcionMensaje = 55 Then
        'concatenar todos los clientes seleccionados para imprimir etiquetas
        If OpcionMensaje = 17 Then
            
            '----------------------------------------------------------------
            Cad = "insert into tmpnlotes (codusu,numalbar,fechaalb,numlinea,codprove) values ("
            Cad = Cad & vUsu.Codigo & ",1,'2005-04-12',1,"
            
            
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    conn.Execute Cad & (ListView2.ListItems(I).Text) & ")"
                    NumRegElim = NumRegElim + 1
                End If
            Next I
            
            
            '----------------------------------------------------------------
            
        Else
            Cad = ""
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    Cad = Cad & Val(ListView2.ListItems(I).Text) & ","
                     'cad = cad & Data1.Recordset.Fields(1) & "|"
                End If
            Next I
            If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
        End If
    ElseIf OpcionMensaje = 11 Then
    'Lineas Factura a rectificar
        'cad = "(" & cadWHERE & ")"
        Cad = ""
        c1 = ""
        c2 = ""
        c3 = ""
        SQL = ""
        For I = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(I).Checked Then
                If SQL = "" Then
                    c1 = DBSet(ListView2.ListItems(I), "T", "N")
                    c2 = ListView2.ListItems(I).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                    Cad = "(codtipoa=" & Trim(c1) & " and numalbar=" & Val(c2) & " and numlinea IN (" & ListView2.ListItems(I).ListSubItems(2)

                Else
                    If Trim(DBSet(ListView2.ListItems(I), "T", "N")) = Trim(c1) And Trim(ListView2.ListItems(I).ListSubItems(1)) = Trim(c2) Then
                    'es el mismo albaran y concatenamos lineas
                        Cad = "," & ListView2.ListItems(I).ListSubItems(2)

                    Else
                        If Cad <> "" Then SQL = SQL & ")) "
                        c1 = DBSet(ListView2.ListItems(I), "T", "N")
                        c2 = ListView2.ListItems(I).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                        Cad = " or (codtipoa=" & Trim(c1) & " and numalbar=" & Val(c2) & " and numlinea IN (" & ListView2.ListItems(I).ListSubItems(2)
                        
'                       cad=cad &
                    End If
                End If
                SQL = SQL & Cad
'                If cad <> "" Then cad = cad & " OR "
'                cad = cad & "(codtipoa=" & DBSet(ListView2.ListItems(i), "T", "N") & " and numalbar=" & Val(ListView2.ListItems(i).ListSubItems(1)) & " and numlinea=" & ListView2.ListItems(i).ListSubItems(2) & ")"
            Else
'                cad = ""
            End If
        Next I
        If Cad <> "" Then
            SQL = SQL & "))"
            Cad = "(" & cadWhere & ") AND (" & SQL & ")"
        End If
'        If cad <> "" Then cad = "(" & cadWHERE & ") AND (" & cad & ")"
    ElseIf OpcionMensaje = 14 Then
        Cad = RegresarCargaEmpresas
    End If
    
    
    
     'Actualizar la tabla sseries asignando los valores correspondientes a los
      'campos: codclien, coddirec, tieneman, codtipom, numalbar, fechavta, numline1
      'y Salir (Volver a Mto Albaranes Clientes (Facturacion)
      PulsadoSalir = True
      'RaiseEvent CargarNumSeries
      RaiseEvent DatoSeleccionado(Cad)
      Unload Me
End Sub


Private Sub CmdAceptarPozos_Click()
Dim Cadena As String
    'Cargo lor registros marcados
    Cadena = ""
    For NumRegElim = 1 To ListView13.ListItems.Count
        If ListView13.ListItems(NumRegElim).Checked Then
            Cadena = Cadena & "'" & ListView13.ListItems(NumRegElim).Text & "',"
        End If
    Next NumRegElim
    ' quitamos la ultima coma
    If Cadena <> "" Then
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
    End If
    
    RaiseEvent DatoSeleccionado(Cadena)
    Unload Me

End Sub

Private Sub cmdacepVariedades_Click()
Dim Cadena As String
    'Cargo las variedades marcadas
    
    
    If OpcionMensaje = 48 Then
        Cadena = ""
        For NumRegElim = 1 To ListView6.ListItems.Count
            If ListView6.ListItems(NumRegElim).Checked Then
                Cadena = Cadena & "('" & ListView6.ListItems(NumRegElim).Text & "','" & Format(ListView6.ListItems(NumRegElim).SubItems(1), "yyyy-mm-dd") & "'),"
            End If
        Next NumRegElim
        ' quitamos la ultima coma
        If Cadena <> "" Then
            Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
        End If
        
        RaiseEvent DatoSeleccionado(Cadena)
        Unload Me
        
        Exit Sub
    Else
        If OpcionMensaje = 56 Then
            Cadena = ""
            
            For NumRegElim = 1 To ListView6.ListItems.Count
                If ListView6.ListItems(NumRegElim).Checked Then
                    Cadena = Cadena & DBSet(ListView6.ListItems(NumRegElim).Text, "F") & ","
                End If
            Next NumRegElim
            ' quitamos la ultima coma
            If Cadena <> "" Then
                Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
            End If
            
            RaiseEvent DatoSeleccionado(Cadena)
            Unload Me
        
            Exit Sub
        End If
    End If
    
    
    Cadena = ""
    For NumRegElim = 1 To ListView6.ListItems.Count
        If ListView6.ListItems(NumRegElim).Checked Then
            Cadena = Cadena & ListView6.ListItems(NumRegElim).Text & ","
        End If
    Next NumRegElim
    ' quitamos la ultima coma
    If Cadena <> "" Then
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
    End If
    
    RaiseEvent DatoSeleccionado(Cadena)
    Unload Me

End Sub

Private Sub CmdAcepVisEntr_Click()
    RaiseEvent DatoSeleccionado("OK")
    Unload Me
End Sub

Private Sub CmdAlbLiq_Click(Index As Integer)
    
    Cadena = Index
    
    RaiseEvent DatoSeleccionado(Cadena)
    Unload Me

End Sub

Private Sub CmdCan_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    If OpcionMensaje = 4 Then
        MsgBox "Debe introducir los nº de serie necesarios para el Albaran.", vbInformation
        Exit Sub
    End If
    PulsadoSalir = True
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdCancelarCobros_Click()
    vCampos = "0"
    Unload Me
End Sub

Private Sub cmdcancelVisEnt_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub CmdCanCrearCampo_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub CmdCanDif_Click()
    PulsadoSalir = True
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub CmdCanPal_Click()
    RaiseEvent DatoSeleccionado("0")
    Unload Me
End Sub

Private Sub cmdCanVariedades_Click()
    If OpcionMensaje = 56 Then
        RaiseEvent DatoSeleccionado("-1")
    Else
        RaiseEvent DatoSeleccionado("")
    End If
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub


Private Sub CmdContinuar_Click()
    Cadena = "1"
    RaiseEvent DatoSeleccionado(Cadena)
    Unload Me
End Sub

Private Sub cmdDeselTodos_Click()
Dim I As Long

    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = False
    Next I
End Sub




Private Sub cmdCamposSocio_Click(Index As Integer)
Dim Cadena As String
Dim It As ListItem

    Select Case Index
        Case 0
            NumRegElim = 0
            
            If OpcionMensaje = 34 Or OpcionMensaje = 51 Or OpcionMensaje = 57 Then
                Cadena = ""
                For NumRegElim = 1 To ListView3.ListItems.Count
                    If ListView3.ListItems(NumRegElim).Checked Then
                        Cadena = Cadena & ListView3.ListItems(NumRegElim).Text & ","
                    End If
                Next NumRegElim
                If Cadena <> "" Then Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
                RaiseEvent DatoSeleccionado(Cadena)
            End If
            
        Case 1
            'Cargo los campos marcados del socio
            Cadena = ""
            For NumRegElim = 1 To ListView3.ListItems.Count
                If ListView3.ListItems(NumRegElim).Checked Then
                    Cadena = Cadena & ListView3.ListItems(NumRegElim).Text & ","
                End If
            Next NumRegElim
            ' quitamos la ultima coma
            If Cadena <> "" Then
                Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
            End If
            
            RaiseEvent DatoSeleccionado(Cadena)
        Case 2
'            Set It = ListView4.ListItems.Item
            If ListView4.ListItems.Count <> 0 Then
                Cadena = ListView4.SelectedItem.Text & "|"
            End If
            
            If OpcionMensaje = 24 Then ' si son los albaranes del transportista devolvemos (numalbar,numnotac)
                Cadena = Cadena & ListView4.SelectedItem.SubItems(1) & "|"
                Cadena = Cadena & ListView4.SelectedItem.SubItems(2) & "|"
                Cadena = Cadena & ListView4.SelectedItem.SubItems(3) & "|"
                Cadena = Cadena & ListView4.SelectedItem.SubItems(4) & "|"
                Cadena = Cadena & ListView4.SelectedItem.SubItems(5) & "|"
                Cadena = Cadena & ListView4.SelectedItem.SubItems(6) & "|"
            End If
            '[Monica]20/02/2011: si hay mas de un campo seleccionado lo mandamos tb para observaciones del parte de adv
            If OpcionMensaje = 7 Then
                Cadena = ""
                For NumRegElim = 1 To ListView4.ListItems.Count
                    If ListView4.ListItems(NumRegElim).Checked Then
                        Cadena = Cadena & ListView4.ListItems(NumRegElim).Text & "|"
                    End If
                Next NumRegElim
            End If
            
            '[Monica]20/02/2011: devolvemos los albaranes del socio de VC que se van a recalcular los importes
            '                    prorrateados segun los kilos netos
            If OpcionMensaje = 26 Then
                Cadena = ""
                For NumRegElim = 1 To ListView4.ListItems.Count
                    If ListView4.ListItems(NumRegElim).Checked Then
                        Cadena = Cadena & ListView4.ListItems(NumRegElim).Text & ","
                    End If
                Next NumRegElim
                If Cadena <> "" Then Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
            End If
            
            
            RaiseEvent DatoSeleccionado(Cadena)
    End Select
    Unload Me
End Sub

Private Sub cmdFacturas_Click()
    Unload Me
End Sub

Private Sub cmdRegresar_Click()
    Unload Me
End Sub

Private Sub CmdSal_Click()
    Unload Me
End Sub

Private Sub CmdSalir1_Click()
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdSalir2_Click()
    Cadena = ListView10.SelectedItem.Text & "|"
    RaiseEvent DatoSeleccionado(Cadena)
    Unload Me
End Sub

Private Sub CmdSalir3_Click()
    Unload Me
End Sub

Private Sub CmdSalir4_Click()
    Cadena = "0"
    RaiseEvent DatoSeleccionado(Cadena)
    Unload Me
End Sub

Private Sub cmdSelTodos_Click()
    Dim I As Long

    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = True
    Next I
End Sub





Private Sub Form_Activate()
Dim OK As Boolean

    
    Select Case OpcionMensaje
        Case 4 ' variedades viene de un rango de productos
            CargarListaVariedades True
        
        Case 5 ' variedades que se ha llevado un socio de bodega
            CargarListaConsumo
            Label1(4).Caption = "Por Variedad:"
            Label1(6).Caption = "Por Producto:"
            Me.Caption = "Consumo del Socio"
        
        Case 6 ' mostrar los campos del socio
            CargarCamposSocio 1
            If Campo <> "" Then SituarCampoSocio CLng(Campo)
            
        Case 7 ' mostrar los campos del socio con la variedad
            CargarCamposSocio 2
            If Campo <> "" Then SituarCampoSocio CLng(Campo)
            
        Case 8, 9, 17, 42, 55 'Etiquetas de clientes/Proveedores/Socios
            CargarListaClientes
        
        Case 10 'Errores al contabilizar facturas
            CargarListaErrContab
        
        Case 11 'Lineas Factura a rectificar
            CargarListaLinFactu
            
        Case 14 'Mostrar Empresas del sistema
            CargarListaEmpresas
            
        Case 15 'Campos de socio
            CargarCamposSocio 0
            
        Case 16 'Variedades viene de un rango de clases
            CargarListaVariedades False
        
        Case 22 ' Trabajadores a seleccionar de la cuadrilla
            CargarListaTrabajadores Campo
        
        Case 23 ' Hidrantes de un socio
            CargarHidrantesSocio
            
        Case 24 ' Entradas de un transportista pendientes de facturar
            CargarAlbaranes
            
        Case 25 ' Campos con el nro de orden incorrecto
            CargarCamposSocio 3
            
        Case 26 ' Albaranes de un socio
            CargarAlbaranesSocio
            
        Case 27 ' Plagas
            CargarPlagas
            
        Case 28
            CargarNotasSinTaraSalida
        
        Case 29 ' mostrar los campos del socio
            CargarCamposSocio 1
            If Campo <> "" Then SituarCampoSocio CLng(Campo)
            
            Label6.visible = True
            Text4.visible = True
        
        '[Monica]12/09/2013:
        Case 52 ' nro de orden de recoleccion por nro y fecha saco el nro de campo y socio
            CargarCamposSocio 6
        
        Case 30
            CargarAlbaranesBodegaSinTarar
    
        Case 31
            CargarFacturasVCsinEntradas
            
        Case 32 ' Tipos de aportaciones
            CargarAportaciones
            
        Case 33
            CargarFacturasVCsinEntradas
        
        Case 34 ' Campos de socios
            CargarCamposSocio 4
            
            
        Case 35 ' Albaranes ya liquidados
            CargarAlbaranesLiquidados
        
        Case 36
            CargarListaSituaciones
            
        Case 37 ' Hidrantes de un socio para facturar
            CargarHidrantesSocioFacturar
            
        Case 38 ' orden del printnou
            Me.Option2(0).Value = True
            Me.Option2(0).SetFocus
            
        Case 39 ' Hidrantes de un Campo para cambiar el socio
            CargarHidrantesCampo
            
        Case 40 ' Archivos a adjuntar
            CargarArchivos
            
        Case 41 ' Entradas con error
            CargarEntradasConError
            
        Case 43 ' Muestra la campaña anterior y actual (Carga de kilos de Moixent)
            txtUser.Text = vUsu.Login
            txtlargo.Text = vUsu.Nombre
            
            CargaEmpresas
            
        Case 44 ' diferencias con indefa
            If PrimeraVez Then
                PulsadoSalir = False
                PrimeraVez = False
            End If
            BuscarDiferencias
            
        Case 45, 46 ' visualizar imagen de contador de escalona/utxera
                    '46  visualizar la imagen del documento
            If Cadena <> "" Then
                Me.Image2.Picture = LoadPicture(Cadena)
            End If
            
        Case 47 ' Cargar las entradas del socio
            CargarAlbaranesPdtesFacturar
         
        Case 48 ' Facturas de Anticipos sin descontar
            CargarAnticiposSinDescontar
            
        Case 49
            PonerFocoBtn Me.cmdCancelarCobros
            
        Case 50 ' contadores a no facturar (POZOS)
            CargarContadoresANoFacturar
            
            Me.Caption = "Contadores a no facturar"
        
    
        Case 51 ' Campos de socios agrupados por nro de campo
            CargarCamposSocio 5
    
    
        Case 53 'asignamos los valores
            txtAux(3).Text = Cadena
            txtAux(4).Text = cadWhere
            txtAux(0).Text = Campo
            
            
        Case 54 ' transportistas
            CargarListaTransportistas
            
        Case 56 ' fechas sin descontar
            CargarFechasSinDescontar
            
        Case 57 ' campos para facturar a manta
            CargarCamposSocio 7
        
        Case 60 ' matriculas dde transportistas
            CargarMatriculas
            
        Case 61 ' facturas de pozos
            nomColumna = ""
            nomColumna2 = ""
            columna = 1
            Orden = 0
            CargarFacturasPozos "importe1", nomColumna2
            
        Case 62 ' creacion de campo de entrada de almazara de abn
            Text8(1).Text = RecuperaValor(Cadena, 1)
            Text9(1).Text = DevuelveValor("select nomsocio from rsocios where codsocio = " & DBSet(Text8(1).Text, "N"))
            Text8(2).Text = RecuperaValor(Cadena, 2)
            Text9(2).Text = DevuelveValor("select nomvarie from variedades where codvarie = " & DBSet(Text8(2).Text, "N"))
            Text8(4).Text = RecuperaValor(Cadena, 3)
            Text8(5).Text = RecuperaValor(Cadena, 4)
            Text8(6).Text = RecuperaValor(Cadena, 5)
        
            PonerFoco Text8(1)
            
        Case 63 ' cargar la previsualizacion de entradas
            CargarPrevisualizacion
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim Cad As String
On Error Resume Next

    Me.FrameCobrosPtes.visible = False
    Me.FrameNSeries.visible = False
    Me.FrameErrores.visible = False
    Me.FrameCamposSocio.visible = False
    Me.FrameEntradasSinClasificar.visible = False
    Me.FrameCampos.visible = False
    Me.FrameVariedades.visible = False
    Me.FrameEntradasSinCRFID.visible = False
    Me.FrameConsumoSocio.visible = False
    Me.FrameEntradasSinSalida.visible = False
    Me.FrameFacturas.visible = False
    Me.FrameAlbaranesLiquidados.visible = False
    Me.FrameHidrantesSocio.visible = False
    Me.FrameOrdenListado.visible = False
    Me.FrameArchivos.visible = False
    Me.FrameEntradasConError.visible = False
    Me.FrameEmpresas.visible = False
    Me.FrameDiferencias.visible = False
    Me.FrameImgContador.visible = False
    Me.FrameHidrantesANoFacturar.visible = False
    Me.FrameCambios.visible = False
    Me.FramePago.visible = False
    Me.frameClaveAcceso.visible = False
    Me.FrameMatriculas.visible = False
    Me.FrameImpFrasPozos.visible = False
    Me.FrameCreacionCampo.visible = False
    Me.FrameVisualizaEntradas.visible = False
    PulsadoSalir = True
    PrimeraVez = True
    
    For H = 1 To 3
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    
    Select Case OpcionMensaje
        Case 1 'Mensaje de Cobros Pendientes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaCobrosPtes
            Me.Caption = "Cobros Pendientes"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 2 'Mensaje de no hay suficiente Stock
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaArtSinStock (vCampos)
            Me.Caption = "Artículos sin stock suficiente"
            PonerFocoBtn Me.cmdAceptarCobros
            
        
'        Case 4 'Listado Nº Series Articulo
'            PonerFrameNSeriesVisible True, H, W
'            Me.Caption = "Nº Serie"
'            Me.Label7(1).Caption = "Seleccione los Nº de serie para el Albaran."
'            Me.Label7(1).FontSize = 12
'            PulsadoSalir = False
            
        Case 5 'Consumo de variedades por socio (bodega)
            H = FrameConsumoSocio.Height
            W = FrameConsumoSocio.Width
            PonerFrameVisible FrameConsumoSocio, True, H, W
            Label1(4).Caption = "por Variedad"
            Label1(6).Caption = "por Producto"
            frmMensajes.Caption = "Consumo del Socio"
            
            
        Case 6, 7  ' 6 = campos del socio para entrada
                   ' 7 = campos del socio para entrada de adv
            H = FrameCampos.Height
            W = FrameCampos.Width
            PonerFrameVisible FrameCampos, True, H, W
        
            
        Case 8, 17 'Etiquetas de Clientes
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Clientes"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            
        Case 9, 42, 55  'Etiquetas de Socios
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Socios"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = False
            Me.cmdDeselTodos.visible = False
            Me.imgCheck1(0).visible = True
            Me.imgCheck1(1).visible = True
            
            Me.cmdAceptarNSeries.Left = 5960
            Me.cmdCancelar.Left = 7040
        
        Case 10 'Errores al contabilizar facturas
            PonerFrameCobrosPtesVisible True, H, W
            Me.Caption = "Facturas NO contabilizadas: "
            PonerFocoBtn Me.cmdAceptarCobros
        
        Case 11 'Lineas Factura a Rectificar
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Lineas Factura a Rectificar"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            Me.cmdAceptarNSeries.Left = Me.cmdAceptarNSeries.Left + 1000
            Me.cmdCancelar.Left = Me.cmdCancelar.Left + 1000
        
        Case 12 'Mensaje Partes de ADV que no se van a Facturar
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaAlbaranes
            Me.Caption = "Facturación Partes ADV"
            Me.Label1(0).Caption = "Existen Partes que NO se van a Facturar:"
            Me.Label1(0).Top = 260
            Me.Label1(0).Left = 480
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 13 'Muestra Errores
            H = 6000
            W = 8800
            PonerFrameVisible Me.FrameErrores, True, H, W
            Me.Text1.Text = vCampos
            Me.Caption = "Errores"
        
        Case 14 'Muestra Empresas del sistema
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Selección"
            CargarListaEmpresas
            
        Case 15
            H = FrameCamposSocio.Height
            W = FrameCamposSocio.Width
            PonerFrameVisible FrameCamposSocio, True, H, W
            
        Case 4, 16 'variedades
            H = FrameVariedades.Height
            W = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, H, W
            

        Case 19 'Entradas sin clasificar
            PonerFrameEntradasSinClasificarVisible True, H, W
            CargarListaEntradas
            Me.Caption = "Entradas Erróneas"
            PonerFocoBtn Me.CmdAceptarPal
        
        Case 20 'Entradas sin clasificar
            PonerFrameEntradasSinClasificarVisible True, H, W
            Label1(2).visible = False
            CmdCanPal.visible = False
            CmdCanPal.Enabled = False
            CargarListaEntradasErr
            Me.Label1(3).Caption = "Entradas Sin Clasificar: "
            PonerFocoBtn Me.CmdCanPal
            
        Case 21 ' Entradas de Bascula sin CRFID
            PonerFrameEntradasSinCRFIDVisible True, H, W
            CargarListaEntradasSinCRFID Cadena
            Me.Label1(3).Caption = "Entradas Sin CRFID: "
            PonerFocoBtn Me.CmdSalir
        
        Case 22 ' Trabajadores de la cuadrilla
            H = FrameVariedades.Height
            W = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, H, W
            Label3.Caption = "Trabajadores"
    
        Case 23 ' Hidrantes de un socio (POZOS)
            H = FrameCampos.Height
            W = FrameCampos.Width
            PonerFrameVisible FrameCampos, True, H, W
                   
        Case 24 ' albaranes de transportista
            H = FrameCampos.Height
            W = FrameCampos.Width
            PonerFrameVisible FrameCampos, True, H, W
            
            Label2.Caption = "Albaranes Pendientes del Transportista"
            
        Case 25
            H = FrameCamposSocio.Height
            W = FrameCamposSocio.Width
            Label5.Caption = "Campos con el Nro.Orden incorrecto"
            cmdCamposSocio(1).Caption = "Corregir"
            PonerFrameVisible FrameCamposSocio, True, H, W
        
        Case 26 ' albaranes de un socio
            H = FrameCampos.Height
            W = FrameCampos.Width
            PonerFrameVisible FrameCampos, True, H, W
            
            Label2.Caption = "Albaranes Venta Campo"
        
        Case 27 ' incidencias
            H = FrameVariedades.Height
            W = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, H, W
            
            Label3.Caption = "Plagas"
        
        Case 28 ' Entradas de Bascula sin tara de salida
            H = FrameEntradasSinSalida.Height
            W = FrameEntradasSinSalida.Width
            PonerFrameVisible FrameEntradasSinSalida, True, H, W
            
        Case 29, 52
            H = FrameCampos.Height
            W = FrameCampos.Width
            PonerFrameVisible FrameCampos, True, H, W
        
            If OpcionMensaje = 52 Then Label2.Caption = "Nro.Ordenes Recolección"
        
            Text4.Text = Campo
    
        Case 30 ' Entradas de Bodega sin tarar
            H = FrameEntradasSinSalida.Height
            W = FrameEntradasSinSalida.Width
            PonerFrameVisible FrameEntradasSinSalida, True, H, W
    
        Case 31 ' Facturas de anticipo venta campo sin entradas
            H = FrameFacturas.Height
            W = FrameFacturas.Width
            PonerFrameVisible FrameFacturas, True, H, W
    
        Case 32 ' Tipo de aportaciones
            H = FrameVariedades.Height
            W = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, H, W
            
            Label3.Caption = "Tipo de Aportaciones"
        
        Case 33 ' Facturas de un socio
            Label9.Caption = "Facturas del Socio "
            
            H = FrameFacturas.Height
            W = FrameFacturas.Width
            PonerFrameVisible FrameFacturas, True, H, W
    
        Case 34, 51, 57 ' relacion de campos
            H = FrameCamposSocio.Height
            W = FrameCamposSocio.Width
            Label5.Caption = "Campos "
            cmdCamposSocio(1).Enabled = False
            cmdCamposSocio(1).visible = False
            PonerFrameVisible FrameCamposSocio, True, H, W
            cmdCamposSocio(0).Caption = "Regresar"
    
    
        Case 35 ' relacion de campos
            H = FrameAlbaranesLiquidados.Height
            W = FrameAlbaranesLiquidados.Width
            
            PonerFrameVisible FrameAlbaranesLiquidados, True, H, W
    
        Case 36 ' situaciones
            H = FrameVariedades.Height
            W = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, H, W
            Label3.Caption = "Situaciones"
            
        Case 37 ' hidrantes de un socio para facturar
            H = FrameHidrantesSocio.Height
            W = FrameHidrantesSocio.Width
            PonerFrameVisible FrameHidrantesSocio, True, H, W
            Label10.Caption = "Hidrantes Socio para Facturar"
            
        Case 38 ' orden del printnou
            H = FrameOrdenListado.Height
            W = FrameOrdenListado.Width
            PonerFrameVisible FrameOrdenListado, True, H + 90, W
        
        Case 39 ' hidrantes de un campo para cambiar el socio
            H = FrameHidrantesSocio.Height
            W = FrameHidrantesSocio.Width
            PonerFrameVisible FrameHidrantesSocio, True, H, W
            Label10.Caption = "Contadores del Campo" ' a modificar"
            '[Monica]30/10/2013: quitamos lo de seleccionar
            ListView13.Checkboxes = False
            
            imgCheck2(0).Enabled = False
            imgCheck2(0).visible = False
            imgCheck2(1).Enabled = False
            imgCheck2(1).visible = False
            
            
        Case 40 ' Archivos a incluir en el email
            H = FrameArchivos.Height
            W = FrameArchivos.Width
            PonerFrameVisible FrameArchivos, True, H, W
            Label10.Caption = "Archivos a adjuntar"
            
        Case 41 ' Entradas clasificadas con error.
            H = Me.FrameEntradasConError.Height
            W = FrameEntradasConError.Width
            PonerFrameVisible FrameEntradasConError, True, H, W
            Label4(1).Caption = "Entradas con Error"
    
        Case 43 ' Empresas para carga de kilos
            H = Me.FrameEmpresas.Height
            W = FrameEmpresas.Width
            PonerFrameVisible FrameEmpresas, True, H, W
    
        Case 44 ' Diferencias con Indefa
            H = Me.FrameDiferencias.Height
            W = FrameDiferencias.Width
            PonerFrameVisible FrameDiferencias, True, H, W
        
        Case 45, 46 ' Imagen de Indefa del contador
            H = Me.FrameImgContador.Height + 500
            W = Me.FrameImgContador.Width + 500
            
            FrameImgContador.visible = True
            FrameImgContador.Top = -90
            FrameImgContador.Width = W
            FrameImgContador.Height = H + 90
            Me.Image2.Height = H + 90
            Me.Image2.Width = W
        
        Case 47 ' albaranes del socio
            H = FrameVariedades.Height
            W = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, H, W
            Label3.Caption = "Albaranes Pdtes del Socio"
        
        
        Case 48 ' anticipos sin descontar
            H = FrameVariedades.Height
            W = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, H, W
            Label3.Caption = "Anticipos sin descontar"
        
        
        Case 49 ' campos con precios de zona a 0
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaCamposSinPrecioZona
            Me.Caption = "Zonas sin precio €/Hda:"
            PonerFocoBtn Me.cmdCancelarCobros
        
        
        Case 50 'Contadores con Consumo inferior al minimo y superior al maximo que no se van a facturar (POZOS)
            H = Me.FrameHidrantesANoFacturar.Height
            W = Me.FrameHidrantesANoFacturar.Width
            PonerFrameVisible FrameHidrantesANoFacturar, True, H, W
            frmMensajes.Caption = "Contadores a no facturar"
        
            Label1(17).Caption = Label1(17).Caption & " " & vParamAplic.ConsumoMinPOZ
            Label1(16).Caption = Label1(16).Caption & " " & vParamAplic.ConsumoMaxPOZ
        
        Case 53 'Frame de cambios
            H = Me.FrameCambios.Height + 150
            W = Me.FrameCambios.Width
            PonerFrameVisible FrameCambios, True, H, W
        
        Case 54 'transportistas
            H = FrameVariedades.Height
            W = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, H, W
            Me.Label3.Caption = "Transportistas"
        
        Case 56 ' fechas sin descontar
            H = FrameVariedades.Height
            W = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, H, W
            Label3.Caption = "Fechas de Anticipos a descontar"
        
        Case 58 ' tipo de recibo
            H = FramePago.Height
            W = FramePago.Width
            PonerFrameVisible FramePago, True, H, W
        
        Case 59 ' clave de acceso
            H = frameClaveAcceso.Height
            W = frameClaveAcceso.Width
            PonerFrameVisible frameClaveAcceso, True, H, W
        
        Case 60 ' transportistas
            H = FrameMatriculas.Height
            W = FrameMatriculas.Width
            PonerFrameVisible FrameMatriculas, True, H, W
        
        Case 61 ' facturas de pozos
            H = FrameImpFrasPozos.Height
            W = FrameImpFrasPozos.Width
            PonerFrameVisible FrameImpFrasPozos, True, H, W
        
        Case 62 ' creacion de un campo de una entrada de abn
            H = Me.FrameCreacionCampo.Height
            W = Me.FrameCreacionCampo.Width
            PonerFrameVisible FrameCreacionCampo, True, H, W
        
        Case 63 ' visualizacion de entradas de almazara previa a insertar (ABN)
            H = Me.FrameVisualizaEntradas.Height
            W = Me.FrameVisualizaEntradas.Width
            PonerFrameVisible FrameVisualizaEntradas, True, H, W
        
    End Select
    'Me.cmdCancel(indFrame).Cancel = True
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerFrameCobrosPtesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 4600
        
    Select Case OpcionMensaje
        Case 1
            H = 5000
            W = 8600
            Me.Label1(0).Caption = "SOCIO: " & vCampos
        Case 2
            W = 8800
            Me.cmdAceptarCobros.Top = 4000
            Me.cmdAceptarCobros.Left = 4200
        Case 5 'Componentes
            W = 6000
            H = 5000
            Me.cmdAceptarCobros.Left = 4000

        Case 6, 7 'Prefacturar Albaranes
            W = 7000
            H = 6000
            Me.cmdAceptarCobros.Top = 5400
            Me.cmdAceptarCobros.Left = 4600

        Case 10, 12 'Errores al contabilizar facturas
            H = 6000
            W = 8400
            Me.cmdAceptarCobros.Top = 5300
            Me.cmdAceptarCobros.Left = 4900
            If OpcionMensaje = 12 Then
                Me.cmdCancelarCobros.Top = 5300
                Me.cmdCancelarCobros.Left = 4600
                Me.cmdAceptarCobros.Left = 3300
                Me.Label1(1).Top = 4800
                Me.Label1(1).Left = 3400
                Me.cmdAceptarCobros.Caption = "&SI"
                Me.cmdCancelarCobros.Caption = "&NO"
            End If
            
        Case 49
            H = 6000
            W = 8400
            Me.cmdAceptarCobros.Top = 5300
            Me.cmdAceptarCobros.Left = 4900
            
            Me.cmdCancelarCobros.Top = 5300
            Me.cmdCancelarCobros.Left = 4600
            Me.cmdAceptarCobros.Left = 3300
            Me.Label1(1).Top = 4800
            Me.Label1(1).Left = 3400
            Me.cmdAceptarCobros.Caption = "&SI"
            Me.cmdCancelarCobros.Caption = "&NO"
            Me.Label1(0).Caption = ""
    End Select
            
    PonerFrameVisible Me.FrameCobrosPtes, visible, H, W

    If visible = True Then
        Me.txtParam.visible = (OpcionMensaje = 6 Or OpcionMensaje = 7)
        Me.Label1(0).visible = (OpcionMensaje = 1) Or (OpcionMensaje = 5) Or (OpcionMensaje = 12) Or (OpcionMensaje = 49)
        Me.cmdCancelarCobros.visible = (OpcionMensaje = 12) Or (OpcionMensaje = 49)
        Me.Label1(1).visible = (OpcionMensaje = 12) Or (OpcionMensaje = 49)
    End If
End Sub

Private Sub PonerFrameEntradasSinClasificarVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

        
    H = 6000
    W = 8400
    Me.CmdAceptarPal.Top = 5300
    Me.CmdAceptarPal.Left = 4900
    Me.CmdCanPal.Top = 5300
    Me.CmdCanPal.Left = 4600
    Me.CmdAceptarPal.Left = 3300
    Me.Label1(2).Top = 4800
    Me.Label1(2).Left = 3400
    Me.CmdAceptarPal.Caption = "&Continuar"
    Me.CmdCanPal.Caption = "&Salir"
        
    PonerFrameVisible Me.FrameEntradasSinClasificar, visible, H, W

End Sub

Private Sub PonerFrameEntradasSinCRFIDVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

        
    H = 4600
    W = 8655
        
    PonerFrameVisible Me.FrameEntradasSinCRFID, visible, H, W

End Sub




Private Sub PonerFrameNSeriesVisible(visible As Boolean, H As Integer, W As Integer)
'Pone el Frame de Nº Serie Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 5000
   
    If OpcionMensaje = 11 Then 'Lineas Factura a Rectificar
        W = 10900
    ElseIf OpcionMensaje = 14 Then
        W = 6500
        Me.Label7(1).visible = True
    Else
        W = 8500
        Me.Label7(1).visible = False
    End If
    PonerFrameVisible Me.FrameNSeries, visible, H, W
End Sub


'Private Sub ponerFrameComponentesVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
''Pone el Frame de Componentes Visible y Ajustado al Formulario, y visualiza los controles
''necesario para el Informe
'
''    Me.FrameComponentes.visible = visible
'    Me.FrameComponentes2.visible = visible
'
'    h = 4000
'    w = 5300
'    PonerFrameVisible Me.FrameComponentes, visible, h, w
'
'    If visible = True Then
'        'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
'        If vParamAplic.Departamento Then
'            Me.OptCompXDpto.Caption = "Departemento"
'        Else
'            Me.OptCompXDpto.Caption = "Dirección"
'        End If
'    End If
'End Sub


Private Sub CargarListaCobrosPtes()
'Muestra la lista Detallada de cobros en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    SQL = "SELECT numserie, codfaccl, fecfaccl, fecvenci, impvenci, impcobro "
    SQL = SQL & " FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
    SQL = SQL & cadWhere
    SQL = SQL & " and (ImpVenci + if(Gastos is null,0,gastos) - if(impcobro is null, 0, impcobro)) <> 0 "
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    ListView1.Top = 900
    ListView1.Height = 3250
    ListView1.Width = 8100
    ListView1.Left = 160
    
    'Los encabezados
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Nº Serie", 760
    ListView1.ColumnHeaders.Add , , "Nº Factura", 1100, 1
    ListView1.ColumnHeaders.Add , , "Fecha Factura", 1250, 2
    ListView1.ColumnHeaders.Add , , "Fecha Venci.", 1200, 2
    ListView1.ColumnHeaders.Add , , "Imp. Venci.(€)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Imp. Cobro(€)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Pte. Cobro(€)", 1250, 1
    
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = RS.Fields(0).Value 'Nº Serie
        ItmX.SubItems(1) = RS.Fields(1).Value 'Nº Factura
        ItmX.SubItems(2) = RS.Fields(2).Value 'Fecha Factura
        ItmX.SubItems(3) = RS.Fields(3).Value 'Fecha Vencimiento
        ItmX.SubItems(4) = RS.Fields(4).Value 'Importe Vencido
        ItmX.SubItems(5) = DBLet(RS.Fields(5).Value, "N") 'Importe Cobrado
        ItmX.SubItems(6) = RS.Fields(4).Value - DBLet(RS.Fields(5).Value, "N") 'Pendiente de cobro
        If ItmX.SubItems(6) > 0 Then
            ItmX.ListSubItems(6).ForeColor = vbRed
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Sub CargarListaArtSinStock(NomTabla As String)
'Muestra la lista Detallada de Articulos que no tienen stock suficiente en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    SQL = "SELECT " & NomTabla & ".codalmac," & NomTabla & ".codartic, " & NomTabla & ".nomartic, salmac.canstock as canstock, SUM(cantidad) as cantidad, canstock-SUM(cantidad) as disp "
    SQL = SQL & "FROM ((" & NomTabla & " INNER JOIN sartic ON " & NomTabla & ".codartic=sartic.codartic) INNER JOIN sfamia ON sartic.codfamia=sfamia.codfamia) "
    SQL = SQL & "INNER JOIN salmac ON " & NomTabla & ".codalmac=salmac.codalmac and " & NomTabla & ".codartic=salmac.codartic "
    SQL = SQL & cadWhere 'Where numpedcl = 2 And sfamia.instalac = 0
    SQL = SQL & "GROUP by " & NomTabla & ".codalmac, " & NomTabla & ".codartic "
    

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    Me.ListView1.Top = 500
     
    'Los encabezados
    ListView1.Width = 8400
    ListView1.Height = 3150
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Alm.", 500
    ListView1.ColumnHeaders.Add , , "Articulo", 1800, 2
    ListView1.ColumnHeaders.Add , , "Dec. Artic", 3300
    ListView1.ColumnHeaders.Add , , "Stock", 950, 2
    ListView1.ColumnHeaders.Add , , "Cantidad", 900, 2
    ListView1.ColumnHeaders.Add , , "No Disp.", 900, 2
    
    While Not RS.EOF
        If RS!disp < 0 Then
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(RS.Fields(0).Value, "000") 'Cod Almacen
            ItmX.SubItems(1) = RS.Fields(1).Value 'Cod Artic
            ItmX.SubItems(2) = RS.Fields(2).Value 'Nom Artic
            ItmX.SubItems(3) = RS.Fields(3).Value 'Stock
            ItmX.SubItems(4) = RS.Fields(4).Value 'Cantidad
            ItmX.SubItems(5) = RS.Fields(5).Value 'No Disp
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


'Private Sub CargarListaNSeries()
''Carga las lista con todos los Nº de serie encontrados en la tabla:sserie
''para el articulo pasado como parametro en la cadwhere: "codartic='00012'"
''y que esten disponibles: numfactu y numalbar no tengan valor
'Dim Rs As ADODB.Recordset
'Dim ItmX As ListItem
'Dim sql As String
'Dim cadLista As String
'Dim Dif As Single
'
'    On Error GoTo ECargarLista
'
'    If cadWHERE2 = "" Then
'        'Mostramos los nº serie libres para seleccionar la cantidad
'        sql = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
'        sql = sql & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
'        sql = sql & cadWHERE 'Where codartic='000012'
'        'seleccionamos los que no esten asignados a ninguna factura ni albaran
'        sql = sql & " AND ((isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='')) "
'        sql = sql & " ORDER BY sserie.codartic, numserie "
'
'    Else 'venimos de modificar la cantidad y seleccionamos los ya asignados
'        If InStr(1, cadWHERE2, "|") > 0 Then
'            Dif = CSng(RecuperaValor(cadWHERE2, 1))
'            cadWHERE2 = RecuperaValor(cadWHERE2, 2)
'
'            'seleccionamos nº serie del albaran que modificamos
'            sql = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
'            sql = sql & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
'            sql = sql & cadWHERE2
'
'
'            If Dif < 0 Then
'                'Si la diferencia de cantidad es < 0, mostrar en la lista los nº serie que
'                'tiene la linea de albaran asignado con todos marcados y desmarcar el que no queremos
'
'            Else
'                'si la diferencia de cantidad es > 0, mostrar en la lista los nº de serie que
'                'ya tenia asignados la linea del albaran más los libres para seleccionar los que añadimos de mas
'                cadLista = ""
'                Set Rs = New ADODB.Recordset
'                Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                While Not Rs.EOF
'                    cadLista = cadLista & ", " & Rs!numserie
'                    Rs.MoveNext
'                Wend
'                Rs.Close
'                Set Rs = Nothing
'
'                'mostrar tambien los nº serie sin asignar
'                sql = sql & " OR (" & Replace(cadWHERE, "WHERE", "") & " and (numalbar=''or isnull(numalbar)))"
'            End If
'        Else
'            'viene de una factura rectificativa, seleccionamos los nº de serie de
'            'esa factura y marcamos los que queremos quitar
'            sql = cadWHERE2
'        End If
'    End If
'
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    'Los encabezados
'    ListView2.Width = 7400
'    Me.ListView2.Height = 3100
'    Me.ListView2.Left = 650
'    ListView2.ColumnHeaders.Clear
'
'    ListView2.ColumnHeaders.Add , , "Nº Serie", 1800
'    ListView2.ColumnHeaders.Add , , "Articulo", 1800
'    ListView2.ColumnHeaders.Add , , "Desc. Artic", 3650
'
'    If Rs.EOF Then Unload Me
'
'    While Not Rs.EOF
'         Set ItmX = ListView2.ListItems.Add
'         ItmX.Text = Rs.Fields(0).Value 'num serie
'         If Dif < 0 Then
'            ItmX.Checked = True
'         ElseIf Dif > 0 Then
'            If InStr(1, cadLista, CStr(Rs!numserie)) > 0 Then
'                ItmX.Checked = True
'            Else
'                ItmX.Checked = False
'            End If
'         Else
'            ItmX.Checked = False
'         End If
'         ItmX.SubItems(1) = Rs.Fields(1).Value 'Desc Artic
'         ItmX.SubItems(2) = Rs.Fields(2).Value 'Nom Artic
'         Rs.MoveNext
'    Wend
'    Rs.Close
'    Set Rs = Nothing
'
'
'ECargarLista:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar Nº Series", Err.Description
'End Sub
'

'Private Sub CargarListaComponentes(opt As Byte)
''Muestra la lista Detallada de cobros en un ListView
''Carga los valores de la tabla scobro de la Contabilidad
'Dim RS As ADODB.Recordset
'Dim ItmX As ListItem
'Dim SQL As String
'Dim Codigo As String, cadCodigo As String
'
'    Select Case opt
'        Case 1 'Mantenimiento
'            Codigo = RecuperaValor(vCampos, 1)
'            If Codigo = "" Then
'                cadCodigo = " isnull(nummante) "
'            Else
'                cadCodigo = " nummante=" & DBSet(Codigo, "T")
'            End If
'            SQL = ObtenerSQLcomponentes(cadWHERE & " and " & cadCodigo)
'            Me.Label1(0).Caption = "Mantenimiento: " & Codigo
'
'        Case 2 'Departamento
'            Codigo = RecuperaValor(vCampos, 2)
'            If Codigo = "" Then
'                cadCodigo = "isnull(coddirec)"
'            Else
'                cadCodigo = " coddirec=" & Codigo
'            End If
'            SQL = ObtenerSQLcomponentes(cadWHERE & " and " & cadCodigo)
'            If vParamAplic.Departamento Then
'                Me.Caption = "Equipos del Departamento"
'                Me.Label1(0).Caption = " Departamento: " & RecuperaValor(vCampos, 3)
'            Else
'                Me.Caption = "Equipos de la Dirección"
'                Me.Label1(0).Caption = " Dirección: " & Codigo & " " & RecuperaValor(vCampos, 3)
'            End If
'
'        Case 3 'Cliente
'            SQL = ObtenerSQLcomponentes(cadWHERE)
'            Me.Caption = "Equipos del Cliente"
'            Me.Label1(0).Caption = "Cliente: " & RecuperaValor(vCampos, 4)
'    End Select
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    'Los encabezados
'    ListView1.Top = 800
'    ListView1.Left = 280
'    ListView1.Width = 4900
'    ListView1.Height = 3250
'    ListView1.ColumnHeaders.Clear
'
'    ListView1.ColumnHeaders.Add , , "TA", 760
'    ListView1.ColumnHeaders.Add , , "Tipo Articulo", 2800
'    ListView1.ColumnHeaders.Add , , "Cantidad", 1280, 2
'
'    If Not RS.EOF Then
'        While Not RS.EOF
'            Set ItmX = ListView1.ListItems.Add
'            ItmX.Text = RS.Fields(0).Value 'TA
'            ItmX.SubItems(1) = RS.Fields(1).Value 'Tipo Articulo
'            ItmX.SubItems(2) = RS.Fields(2).Value 'Cantidad
'            RS.MoveNext
'        Wend
'    End If
'    RS.Close
'    Set RS = Nothing
'End Sub




Private Sub CargarListaClientes()
'Carga las lista con todos los clientes seleccionados en la tabla:sclien
'para imprimir etiquetas, pasando como parametro la cadwhere
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String, Men As String

    On Error GoTo ECargarLista

    Select Case OpcionMensaje
    Case 8
        'CLIENTES
        SQL = "SELECT codclien,nomclien,cifclien "
        SQL = SQL & "FROM clientes "
        If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
        SQL = SQL & " ORDER BY codclien "
        Men = "Cliente"
    Case 9
        'SOCIOS
        SQL = "SELECT distinct rsocios.codsocio,nomsocio,nifsocio "
        SQL = SQL & "FROM rsocios inner join rsocios_seccion on rsocios.codsocio = rsocios_seccion.codsocio "
        If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Or vParamAplic.Cooperativa = 12 Then
            SQL = SQL & " ORDER BY rsocios.nomsocio "
        Else
            SQL = SQL & " ORDER BY rsocios.codsocio "
        End If
        Men = "Socio"
    Case 17
        'CLIENTES MANTENIMIENTO
        SQL = cadWhere
    
    Case 42
        SQL = "SELECT distinct rsocios.codsocio,nomsocio,sum(rrecibpozos.totalfact) totalfact "
        SQL = SQL & "FROM rsocios inner join rrecibpozos on rsocios.codsocio = rrecibpozos.codsocio "
        If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
        SQL = SQL & " GROUP BY 1,2 "
        SQL = SQL & " ORDER BY rsocios.codsocio "
        Men = "Socio"
    
    Case 55
        SQL = cadWhere
    
    
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'Los encabezados
        ListView2.Width = 7000
        ListView2.Top = 500
        ListView2.Height = 3620
        ListView2.ColumnHeaders.Clear
        
        ListView2.ColumnHeaders.Add , , Men, 1350
        ListView2.ColumnHeaders.Add , , "Nombre", 4000
        
        If OpcionMensaje <> 42 Then
            ListView2.ColumnHeaders.Add , , "NIF", 1330
        Else
            ListView2.ColumnHeaders.Add , , "Importe", 1330
        End If
        
        While Not RS.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = Format(RS.Fields(0).Value, "000000") 'cod clien/prove
             ItmX.Checked = False
             ItmX.SubItems(1) = RS.Fields(1).Value 'Nom clien/prove
             ItmX.SubItems(2) = RS.Fields(2).Value 'NIF clien/prove
             RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar " & Men, Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub



Private Sub CargarListaErrContab()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = " SELECT  * "
    SQL = SQL & " FROM tmpErrFac "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        ListView1.Height = 4500
        ListView1.Width = 7400
        ListView1.Left = 500
        ListView1.Top = 500

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        If RS.Fields(0).Name = "codprove" Then
            'Facturas de Compra
             ListView1.ColumnHeaders.Add , , "Prove.", 700
        Else 'Facturas de Venta
            ListView1.ColumnHeaders.Add , , "Tipo", 600
        End If
        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = RS.Fields(0).Value
            ItmX.SubItems(1) = Format(RS!numfactu, "0000000")
            ItmX.SubItems(2) = RS!fecfactu
            ItmX.SubItems(3) = RS!Error
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub


Private Sub CargarListaLinFactu()
'Carga las lista con todas las lineas de la factura que estamos rectificando
'seleccionamos las que nos queremos llevar al Albaran de rectificacion
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarLista

    SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre"
    SQL = SQL & " FROM slifac "
    If cadWhere <> "" Then SQL = SQL & " WHERE " & cadWhere
    SQL = SQL & " ORDER BY codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        
        ListView2.Top = 500
        ListView2.Left = 380
        ListView2.Width = 10100
        ListView2.Height = 3620
        
        'Los encabezados
        ListView2.ColumnHeaders.Clear
    
        ListView2.ColumnHeaders.Add , , "T.Alb", 660
        ListView2.ColumnHeaders.Add , , "Nº Alb", 840
        ListView2.ColumnHeaders.Add , , "Lin.", 450
         ListView2.ColumnHeaders.item(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Alm", 460
        ListView2.ColumnHeaders.Add , , "Artic", 1380
        ListView2.ColumnHeaders.Add , , "Desc. Artic.", 2500
        ListView2.ColumnHeaders.Add , , "Cant.", 600
        ListView2.ColumnHeaders.item(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Precio", 960
        ListView2.ColumnHeaders.item(8).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 1", 600
        ListView2.ColumnHeaders.item(9).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 2", 600
        ListView2.ColumnHeaders.item(10).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Importe", 950
        ListView2.ColumnHeaders.item(11).Alignment = lvwColumnRight
    
        While Not RS.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = RS!codtipoa 'cod tipo alb
             ItmX.Checked = False
             ItmX.SubItems(1) = Format(RS!Numalbar, "0000000") 'Nº Albaran
             ItmX.SubItems(2) = RS!numlinea 'linea Albaran
             ItmX.SubItems(3) = Format(RS!codAlmac, "000") 'cod almacen
             ItmX.SubItems(4) = RS!codArtic 'Cod Articulo
             ItmX.SubItems(5) = RS!NomArtic 'Nombre del Articulo
             ItmX.SubItems(6) = RS!cantidad
             ItmX.SubItems(7) = Format(RS!precioar, FormatoPrecio)
             ItmX.SubItems(8) = RS!dtoline1
             ItmX.SubItems(9) = RS!dtoline2
             ItmX.SubItems(10) = Format(RS!ImporteL, FormatoImporte)
             RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
    'si aparece la barra de desplazamiento ajustar el ancho
    If Me.ListView2.ListItems.Count > 11 Then
        Me.ListView2.ColumnHeaders(5).Width = 1200 'codartic
        Me.ListView2.ColumnHeaders(8).Width = 920  'precio
    End If
   
    
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Lineas Factura", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub




Private Sub CargarListaAlbaranes()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = cadWhere 'cadwhere ya le pasamos toda la SQL
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        ListView1.Height = 3900
        ListView1.Width = 7200
        ListView1.Left = 500
        ListView1.Top = 700

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        ListView1.ColumnHeaders.Add , , "Nº Parte", 900
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Tratto.", 800
        ListView1.ColumnHeaders.Add , , "Socio", 900
        ListView1.ColumnHeaders.Add , , "Nombre", 2500
        ListView1.ColumnHeaders.Add , , "Campo", 900
    
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(RS!Numparte, "0000000")
            ItmX.SubItems(1) = RS!Fechapar
            ItmX.SubItems(2) = RS!codtrata
            ItmX.SubItems(3) = Format(RS!Codsocio, "000000")
            ItmX.SubItems(4) = RS!nomsocio
            ItmX.SubItems(5) = RS!codcampo
            
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

    Set ItmX = Nothing
    
ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub CargarListaEntradas()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = "select numnotac, tmpclasifica.codsocio, nomsocio, case codclase when 0 then 'Sin Clasificar' when 1 then 'Gastos Erróneos' when 2 then 'Nota Duplicada' end from tmpclasifica, rsocios where codusu = " & vUsu.Codigo
    SQL = SQL & " and tmpclasifica.codsocio = rsocios.codsocio "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        ListView5.Height = 3900
        ListView5.Width = 7200
        ListView5.Left = 500
        ListView5.Top = 700

        'Los encabezados
        ListView5.ColumnHeaders.Clear

        ListView5.ColumnHeaders.Add , , "Nº Nota", 1000
        ListView5.ColumnHeaders.Add , , "Socio", 1000, 2
        ListView5.ColumnHeaders.Add , , "Nombre", 3000, 0
        ListView5.ColumnHeaders.Add , , "Tipo Error", 2000, 0
    
        While Not RS.EOF
            Set ItmX = ListView5.ListItems.Add
            ItmX.Text = Format(RS!numnotac, "000000")
            ItmX.SubItems(1) = Format(RS!Codsocio, "000000")
            ItmX.SubItems(2) = RS.Fields(2).Value
            ItmX.SubItems(3) = RS.Fields(3).Value
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub CargarListaEntradasErr()
'Muestra la lista Detallada de entradas que no se clasificaron
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = "select numnotac, tmperrent.codvarie, variedades.nomvarie from tmperrent, variedades where  "
    SQL = SQL & " tmperrent.codvarie = variedades.codvarie "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        ListView5.Height = 3900
        ListView5.Width = 7200
        ListView5.Left = 500
        ListView5.Top = 700

        'Los encabezados
        ListView5.ColumnHeaders.Clear

        ListView5.ColumnHeaders.Add , , "Nº Nota", 1000
        ListView5.ColumnHeaders.Add , , "Código", 1000, 2
        ListView5.ColumnHeaders.Add , , "Variedad", 1800, 0
        ListView5.ColumnHeaders.Add , , "", 3200, 0
    
        While Not RS.EOF
            Set ItmX = ListView5.ListItems.Add
            ItmX.Text = Format(RS!numnotac, "000000")
            ItmX.SubItems(1) = Format(RS!codvarie, "000000")
            ItmX.SubItems(2) = RS.Fields(2).Value
            ItmX.SubItems(3) = "Variedad sin calidad venta campo"
            
            'en campo traemos "/retirada"
            If Campo <> "" Then ItmX.SubItems(3) = ItmX.SubItems(3) & Campo

            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub CargarListaEntradasSinCRFID(SQL As String)
'Muestra la lista Detallada de entradas que no tienen CRFID
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem

    On Error GoTo ECargarList


    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        'Los encabezados
        ListView7.ColumnHeaders.Clear

        ListView7.ColumnHeaders.Add , , "Nº Nota", 1000
        ListView7.ColumnHeaders.Add , , "Código", 1000, 2
        ListView7.ColumnHeaders.Add , , "Variedad", 2000, 0
        ListView7.ColumnHeaders.Add , , "", 3000, 0
    
        While Not RS.EOF
            Set ItmX = ListView7.ListItems.Add
            ItmX.Text = Format(RS!numnotac, "000000")
            ItmX.SubItems(1) = Format(RS!codvarie, "000000")
            ItmX.SubItems(2) = RS.Fields(2).Value
            ItmX.SubItems(3) = "Entradas sin CRFID."
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub CargarListaEmpresas()
'Carga las lista con todas las empresas que hay en el sistema
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String
Dim I As Integer

Dim Prohibidas As String

    On Error GoTo ECargarLista

    VerEmresasProhibidas Prohibidas
    
    SQL = "Select * from usuarios.empresasariagro order by codempre"
    Set ListView2.SmallIcons = frmPpal.ImageListB
    ListView2.Width = 5000
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Empresa", 4900
    ListView2.HideColumnHeaders = True
    ListView2.GridLines = False
    ListView2.ListItems.Clear
    
    Set RS = New ADODB.Recordset
    I = -1
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        SQL = "|" & RS!codempre & "|"
        If InStr(1, Prohibidas, SQL) = 0 Then
            Set ItmX = ListView2.ListItems.Add(, , RS!nomempre, , 5)
            ItmX.Tag = RS!codempre
            If ItmX.Tag = vEmpresa.codempre Then
                ItmX.Checked = True
                I = ItmX.Index
            End If
            ItmX.ToolTipText = RS!Ariagro
        End If
        RS.MoveNext
    Wend
    RS.Close
    If I > 0 Then Set ListView2.SelectedItem = ListView2.ListItems(I)

    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargando datos empresas", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub


Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)
Dim SQL As String
Dim RS As ADODB.Recordset

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    SQL = "Select codempre from usuarios.usuarioempresasariagro WHERE codusu = " & (vUsu.Codigo Mod 1000)
    SQL = SQL & " order by codempre"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
          VarProhibidas = VarProhibidas & RS!codempre & "|"
          RS.MoveNext
    Wend
    RS.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte técnico"
    Set RS = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PulsadoSalir = False Then Cancel = 1
    If OpcionMensaje = 49 And vCampos <> "1" Then vCampos = "0"

End Sub



Private Function ObtenerTamanyosArray() As Boolean
'Para el frame de los Nº de Serie de los Articulos
'En cada indice pone en CodArtic(i) el codigo del articulo
'y en Cantidad(i) la cantidad solicitada de cada codartic
Dim I As Integer, J As Integer

    ObtenerTamanyosArray = False
    'Primero a los campos de la tabla
    TotalArray = -1
    J = 0
    Do
        I = J + 1
        J = InStr(I, vCampos, "·")
        If J > 0 Then TotalArray = TotalArray + 1
    Loop Until J = 0
    
    If TotalArray < 0 Then Exit Function
    
    'Las redimensionaremos
    ReDim codArtic(TotalArray)
    ReDim cantidad(TotalArray)
    
    ObtenerTamanyosArray = True
End Function


Private Function SeparaCampos() As Boolean
'Para el frame de los Nº de Serie de los Articulos
Dim Grupo As String
Dim I As Integer
Dim J As Integer
Dim C As Integer 'Contador dentro del array

    SeparaCampos = False
    I = 0
    C = 0
    Do
        J = I + 1
        I = InStr(J, vCampos, "·")
        If I > 0 Then
            Grupo = Mid(vCampos, J, I - J)
            'Y en la martriz
            InsertaGrupo Grupo, C
            C = C + 1
        End If
    Loop Until I = 0
    SeparaCampos = True
End Function


Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim J As Integer
Dim Cad As String

    J = 0
    Cad = ""
    
    'Cod Artic
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    codArtic(Contador) = Cad
    
    'Cantidad
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        Cad = Grupo
        Grupo = ""
    End If
    cantidad(Contador) = Cad
End Sub








Private Sub frmPar_DatoSeleccionado(CadenaSeleccion As String)
    Text8(3).Text = RecuperaValor(CadenaSeleccion, 1) 'partida
    FormateaCampo Text8(3)
    Text9(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomparti
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text8(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo Text8(1)
    Text9(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    Text8(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    FormateaCampo Text8(2)
    Text9(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear

    Select Case Index
       Case 1 'Socios
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text8(1)
    
    
       Case 2 'Variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text8(2)
       
       
       Case 3 'Partidas
            Set frmPar = New frmManPartidas
            frmPar.DeConsulta = True
            frmPar.DatosADevolverBusqueda = "0|1|2|3|4|5|"
            frmPar.CodigoActual = Text8(3).Text
            frmPar.Show vbModal
            Set frmPar = Nothing
            PonerFoco Text8(3)
    End Select
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim b As Boolean
    If Index < 2 Then
        'En el listview3
        b = Index = 1
        For TotalArray = 1 To ListView3.ListItems.Count
            ListView3.ListItems(TotalArray).Checked = b
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
    Else
        'En el listview6
        b = Index = 2
        For TotalArray = 1 To ListView6.ListItems.Count
            ListView6.ListItems(TotalArray).Checked = b
            If (TotalArray Mod 50) = 0 Then DoEvents
        Next TotalArray
    End If
End Sub



Private Sub imgCheck2_Click(Index As Integer)
Dim b As Boolean
    'En el listview33
    b = Index = 1
    For TotalArray = 1 To ListView13.ListItems.Count
        ListView13.ListItems(TotalArray).Checked = b
        If (TotalArray Mod 50) = 0 Then DoEvents
    Next TotalArray
End Sub





Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub



Private Function RegresarCargaEmpresas() As String
Dim SQL As String
Dim Parametros As String
Dim I As Integer

    CadenaDesdeOtroForm = ""
    
        SQL = ""
        Parametros = ""
        For I = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(I).Checked Then
                SQL = SQL & Me.ListView2.ListItems(I).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next I
        CadenaDesdeOtroForm = Len(Parametros) & "|" & SQL
        'Vemos las conta
        SQL = ""
        For I = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(I).Checked Then
                SQL = SQL & Me.ListView2.ListItems(I).Tag & "|"
            End If
        Next I
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
    
    
        RegresarCargaEmpresas = CadenaDesdeOtroForm

End Function



Private Sub CargarCamposSocio(Opcion As Integer)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    Select Case Opcion
    Case 0, 2
        SQL = "select rcampos.codcampo, rcampos.codvarie, variedades.nomvarie, rpartida.nomparti, "
        SQL = SQL & " rcampos.poligono, rcampos.parcela, rcampos.nrocampo  from rcampos, variedades, rpartida where "
        SQL = SQL & " rcampos.codvarie = variedades.codvarie and rcampos.codparti = rpartida.codparti "
    Case 1
        SQL = "select rcampos.codcampo, rcampos.codparti, rpartida.nomparti, rpartida.codpobla, rpueblos.despobla, "
        SQL = SQL & " rcampos.poligono, rcampos.parcela, rcampos.nrocampo, round(rcampos.supcoope / "
        SQL = SQL & DBSet(vParamAplic.Faneca, "N") & " ,2) hdas, rcampos.subparce from rcampos, rpartida, rpueblos where "
        SQL = SQL & " rcampos.codparti = rpartida.codparti and rpartida.codpobla = rpueblos.codpobla "
    
    Case 3
        SQL = "select rcampos.codcampo, rcampos.nrocampo, rpartida.nomparti, variedades.nomvarie,  "
        SQL = SQL & " rsocios.nomsocio  from rcampos, variedades, rsocios, rpartida where "
        SQL = SQL & " rcampos.codvarie = variedades.codvarie and rcampos.codsocio = rsocios.codsocio and rcampos.codparti = rpartida.codparti "
    
    Case 4
        SQL = "select rcampos.codcampo, rcampos.nrocampo, rpartida.nomparti, variedades.nomvarie,  "
        SQL = SQL & " rsocios.nomsocio  from rcampos, variedades, rsocios, rpartida where "
        SQL = SQL & " rcampos.codvarie = variedades.codvarie and rcampos.codsocio = rsocios.codsocio and rcampos.codparti = rpartida.codparti "
    
    
    Case 5
        SQL = "select rcampos.nrocampo, rpartida.nomparti, variedades.nomvarie,  "
        SQL = SQL & " rsocios.nomsocio  from rcampos, variedades, rsocios, rpartida where "
        SQL = SQL & " rcampos.codvarie = variedades.codvarie and rcampos.codsocio = rsocios.codsocio and rcampos.codparti = rpartida.codparti "
    
    Case 6
        SQL = "select distinct rcampos_ordrec.nroorden, rcampos_ordrec.fecimpre, rcampos.nrocampo, rpartida.nomparti, variedades.nomvarie, rsocios.nomsocio  "
        SQL = SQL & " from rcampos, rcampos_ordrec, variedades, rpartida, rsocios where rcampos.codcampo = rcampos_ordrec.codcampo and "
        SQL = SQL & " rcampos.codvarie = variedades.codvarie and rcampos.codsocio = rsocios.codsocio and rcampos.codparti = rpartida.codparti "
        
    Case 7
        SQL = "select rcampos.codcampo, rpartida.nomparti, rcampos.poligono, rcampos.parcela, rcampos.subparce, variedades.nomvarie, "
        SQL = SQL & " rsocios.nomsocio  from rcampos, variedades, rsocios, rpartida where "
        SQL = SQL & " rcampos.codvarie = variedades.codvarie and rcampos.codsocio = rsocios.codsocio and rcampos.codparti = rpartida.codparti "
    
    End Select
    
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    If Opcion = 4 Then
        SQL = SQL & " order by rcampos.codvarie, rcampos.codsocio "
    End If
    If Opcion = 5 Then
        SQL = SQL & " group by 1,2,3,4 "
        '[Monica]30/09/2013: antes el orden era 1,2,3,4
        SQL = SQL & " order by 2,3,4"
    End If
    If Opcion = 6 Then
        SQL = SQL & " order by 1,2 "
    End If
    
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Select Case Opcion
        Case 0
            ListView3.ColumnHeaders.Clear
        
            ListView3.ColumnHeaders.Add , , "Campo", 1200
            ListView3.ColumnHeaders.Add , , "Codigo", 800, 1
            ListView3.ColumnHeaders.Add , , "Variedad", 1500, 1
            ListView3.ColumnHeaders.Add , , "Partida", 1800
            ListView3.ColumnHeaders.Add , , "Poligono", 800
            ListView3.ColumnHeaders.Add , , "Parcela", 600
'            ListView3.ColumnHeaders.Add , , "SbP", 600
            ListView3.ColumnHeaders.Add , , "Nro.", 700
        Case 2
            ListView4.ColumnHeaders.Clear
        
            ListView4.ColumnHeaders.Add , , "Campo", 1200
            ListView4.ColumnHeaders.Add , , "Codigo", 800, 1
            ListView4.ColumnHeaders.Add , , "Variedad", 1500, 1
            ListView4.ColumnHeaders.Add , , "Partida", 1800
            ListView4.ColumnHeaders.Add , , "Poligono", 800
            ListView4.ColumnHeaders.Add , , "Parcela", 600
'            ListView4.ColumnHeaders.Add , , "SbP", 600
            ListView4.ColumnHeaders.Add , , "Nro.", 600
        
            '[Monica]20/02/2011: en partes de adv puedo seleccionar mas de un campo que va a observaciones
            If OpcionMensaje = 7 Then ListView4.Checkboxes = True
        Case 3
            ListView3.ColumnHeaders.Clear
        
            ListView3.ColumnHeaders.Add , , "Campo", 1200
            ListView3.ColumnHeaders.Add , , "NºOrden", 900, 1
            ListView3.ColumnHeaders.Add , , "Partida", 1800
            ListView3.ColumnHeaders.Add , , "Variedad", 1500
            ListView3.ColumnHeaders.Add , , "Socio", 2400
            
        Case 4 ' campos de un socio
            ListView3.ColumnHeaders.Clear
        
            ListView3.ColumnHeaders.Add , , "Campo", 1200
            ListView3.ColumnHeaders.Add , , "NºOrden", 900, 1
            ListView3.ColumnHeaders.Add , , "Partida", 1800
            ListView3.ColumnHeaders.Add , , "Variedad", 1500
            ListView3.ColumnHeaders.Add , , "Socio", 2400
   
        Case 5 ' campos de un socio
            ListView3.ColumnHeaders.Clear
        
            ListView3.ColumnHeaders.Add , , "NºOrden", 900
            ListView3.ColumnHeaders.Add , , "Partida", 1800
            ListView3.ColumnHeaders.Add , , "Variedad", 1500
            ListView3.ColumnHeaders.Add , , "Socio", 3400
   
        Case 6 ' ordenes de recoleccion emitidas
            ListView4.ColumnHeaders.Clear
        
            ListView4.ColumnHeaders.Add , , "NºOrden", 1000
            ListView4.ColumnHeaders.Add , , "Fecha", 1100
            ListView4.ColumnHeaders.Add , , "NºCampo", 1000
            ListView4.ColumnHeaders.Add , , "Partida", 1500
            ListView4.ColumnHeaders.Add , , "Variedad", 1100
            ListView4.ColumnHeaders.Add , , "Socio", 2800
   
        Case 7
            ListView3.ColumnHeaders.Clear
        
            ListView3.ColumnHeaders.Add , , "Campo", 1200
            ListView3.ColumnHeaders.Add , , "Partida", 1500
            ListView3.ColumnHeaders.Add , , "Poligono", 800
            ListView3.ColumnHeaders.Add , , "Parcela", 600
            ListView3.ColumnHeaders.Add , , "SbP", 600
            ListView3.ColumnHeaders.Add , , "Variedad", 1500
            ListView3.ColumnHeaders.Add , , "Socio", 1500
   
   End Select
    
    TotalArray = 0
    While Not RS.EOF
        Select Case Opcion
            Case 0
                Set It = ListView3.ListItems.Add
            Case 1, 2
                Set It = ListView4.ListItems.Add
            Case 3
                Set It = ListView3.ListItems.Add
            Case 4
                Set It = ListView3.ListItems.Add
            Case 5
                Set It = ListView3.ListItems.Add
            Case 6
                Set It = ListView4.ListItems.Add
            Case 7
                Set It = ListView3.ListItems.Add
        End Select
            
        'It.Tag = DevNombreSQL(RS!codCampo)
        If Opcion = 6 Then
            It.Text = DBLet(RS!nroorden, "N")
        Else
            If Opcion = 7 Then
                It.Text = DBLet(RS!codcampo, "N")
            Else
                If Opcion <> 5 Then
                    It.Text = DBLet(RS!codcampo, "N")
                Else
                    It.Text = DBLet(RS!NroCampo, "N")
                End If
            End If
        End If
        
        If Opcion = 0 Or Opcion = 2 Then
            It.SubItems(1) = Format(RS!codvarie, "000000")
            It.SubItems(2) = RS!nomvarie
            It.SubItems(3) = RS!nomparti
            It.SubItems(4) = RS!Poligono
            It.SubItems(5) = RS!Parcela
            It.SubItems(6) = RS!NroCampo
        Else
            If Opcion = 3 Or Opcion = 4 Then
                It.SubItems(1) = RS!NroCampo
                It.SubItems(2) = RS!nomparti
                It.SubItems(3) = RS!nomvarie
                It.SubItems(4) = RS!nomsocio
            Else
                If Opcion = 5 Then
                    It.SubItems(1) = RS!nomparti
                    It.SubItems(2) = RS!nomvarie
                    It.SubItems(3) = RS!nomsocio
                Else
                    If Opcion = 6 Then
                        It.SubItems(1) = RS!fecimpre
                        It.SubItems(2) = DBLet(RS!NroCampo, "N")
                        It.SubItems(3) = RS!nomparti
                        It.SubItems(4) = RS!nomvarie
                        It.SubItems(5) = RS!nomsocio
                    Else
                        If Opcion = 7 Then
                            It.SubItems(1) = RS!nomparti
                            It.SubItems(2) = DBLet(RS!Poligono, "N")
                            It.SubItems(3) = DBLet(RS!Parcela, "N")
                            It.SubItems(4) = DBLet(RS!SubParce, "T")
                            It.SubItems(5) = RS!nomvarie
                            It.SubItems(6) = RS!nomsocio
                        Else
                            It.SubItems(1) = RS!nomparti
                            It.SubItems(2) = DBLet(RS!CodPobla, "T")
                            It.SubItems(3) = DBLet(RS!desPobla, "T")
                            It.SubItems(4) = RS!Poligono
                            It.SubItems(5) = RS!Parcela
                            It.SubItems(6) = DBLet(RS!SubParce, "T")
                            It.SubItems(7) = RS!NroCampo
                            It.SubItems(8) = RS!Hdas
                        End If
                    End If
                End If
            End If
        End If
        
        If Opcion = 0 Then It.Checked = False
        If Opcion = 4 Or Opcion = 5 Then It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1

'       [Monica]20/02/2011: en Alzira en partes de adv pueden introducir mas de un campo que irá a observaciones
'                           Sólo marcamos el primero
' Manolo no lo quiere
'        If OpcionMensaje = 7 And Opcion = 2 And TotalArray = 1 Then It.Checked = True
        
        
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
    If (Opcion = 1 Or Opcion = 2) And Campo <> "" Then SituarCampoSocio CLng(Campo)
End Sub

Private Sub CargarListaVariedades(DadoProducto As Boolean)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    If DadoProducto Then ' viene de un rango de productos
        SQL = "select variedades.codvarie, variedades.nomvarie, variedades.codprodu, productos.nomprodu from variedades, productos "
        SQL = SQL & " where variedades.codprodu = productos.codprodu "
    Else ' viene de un rango de clases
        SQL = "select variedades.codvarie, variedades.nomvarie, variedades.codclase, clases.nomclase from variedades, clases "
        SQL = SQL & " where variedades.codclase = clases.codclase "
    End If
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.ColumnHeaders.Clear
    
    If DadoProducto Then
        ListView6.ColumnHeaders.Add , , "Código", 1000.0631
        ListView6.ColumnHeaders.Add , , "Variedad", 2200.2522, 1
        ListView6.ColumnHeaders.Add , , "Producto", 799.9371, 1
        ListView6.ColumnHeaders.Add , , "Descripción", 2101.0396
    Else
        ListView6.ColumnHeaders.Add , , "Código", 1000.0631
        ListView6.ColumnHeaders.Add , , "Variedad", 2200.2522, 1
        ListView6.ColumnHeaders.Add , , "Clase", 799.9371, 1
        ListView6.ColumnHeaders.Add , , "Descripción", 2101.0396
    End If
    
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView6.ListItems.Add
            
        It.Text = Format(DBLet(RS!codvarie, "N"), "000000")
        It.SubItems(1) = DBLet(RS!nomvarie, "T")
        If DadoProducto Then
            It.SubItems(2) = Format(DBLet(RS!codprodu, "N"), "000")
            It.SubItems(3) = DBLet(RS!nomprodu, "T")
        Else
            It.SubItems(2) = Format(DBLet(RS!codclase, "N"), "000")
            It.SubItems(3) = DBLet(RS!nomclase, "T")
        End If
        It.Checked = False
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub


Private Sub CargarListaConsumo()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim Consumido As Currency

    'CONSUMO DEL SOCIO POR VARIEDAD

    SQL = "select rbodalbaran_variedad.codvarie, variedades.nomvarie,sum(rbodalbaran_variedad.unidades) as unidades, sum(rbodalbaran_variedad.cantidad) as cantidad "
    SQL = SQL & " from variedades, rbodalbaran_variedad, rbodalbaran "
    SQL = SQL & " where variedades.codvarie = rbodalbaran_variedad.codvarie "
    SQL = SQL & " and rbodalbaran_variedad.numalbar = rbodalbaran.numalbar "
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    SQL = SQL & " group by 1,2 order by 1,2 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView8.ColumnHeaders.Clear
    
    ListView8.ColumnHeaders.Add , , "Código", 700
    ListView8.ColumnHeaders.Add , , "Variedad", 1740 ' 1450 '1740
    ListView8.ColumnHeaders.Add , , "Unidades", 900, 1
    ListView8.ColumnHeaders.Add , , "Cantidad", 900, 1
    
    
    While Not RS.EOF
        Set It = ListView8.ListItems.Add
            
        It.Text = Format(DBLet(RS!codvarie, "N"), "000000")
        It.SubItems(1) = DBLet(RS!nomvarie, "T")
        It.SubItems(2) = Format(DBLet(RS!Unidades, "N"), "###,##0.00")
        It.SubItems(3) = Format(DBLet(RS!cantidad, "N"), "###,##0.00")
        It.Checked = False
        
        RS.MoveNext
    Wend
    RS.Close
    
    
    'CONSUMO DEL SOCIO POR PRODUCTO

    SQL = "select productos.codprodu, productos.nomprodu,sum(rbodalbaran_variedad.unidades) as unidades, sum(rbodalbaran_variedad.cantidad) as cantidad "
    SQL = SQL & " from variedades, rbodalbaran_variedad, rbodalbaran, productos "
    SQL = SQL & " where variedades.codvarie = rbodalbaran_variedad.codvarie "
    SQL = SQL & " and rbodalbaran_variedad.numalbar = rbodalbaran.numalbar "
    SQL = SQL & " and variedades.codprodu = productos.codprodu "
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    SQL = SQL & " group by 1,2 order by 1,2 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView9.ColumnHeaders.Clear
    
    ListView9.ColumnHeaders.Add , , "Código", 700
    ListView9.ColumnHeaders.Add , , "Producto", 1740
    ListView9.ColumnHeaders.Add , , "Unidades", 900, 1
    ListView9.ColumnHeaders.Add , , "Cantidad", 900, 1
    
    
    While Not RS.EOF
        Set It = ListView9.ListItems.Add
            
        It.Text = Format(DBLet(RS!codprodu, "N"), "000000")
        It.SubItems(1) = DBLet(RS!nomprodu, "T")
        It.SubItems(2) = Format(DBLet(RS!Unidades, "N"), "###,##0.00")
        It.SubItems(3) = Format(DBLet(RS!cantidad, "N"), "###,##0.00")
        It.Checked = False
        
        RS.MoveNext
    Wend
    RS.Close
    
    
    'LITROS ACEITE
    
    SQL = "select sum(round(rhisfruta.prestimado * rhisfruta.kilosnet / 100, 0)) "
    SQL = SQL & " from variedades, rhisfruta, productos"
    SQL = SQL & " where rhisfruta.codvarie = variedades.codvarie "
    SQL = SQL & " and variedades.codprodu = productos.codprodu "
    SQL = SQL & " and productos.codgrupo = 5 "
    If cadWhere <> "" Then SQL = SQL & Replace(cadWhere, "rbodalbaran", "rhisfruta")
    
    Text2.Text = Format(CCur(DevuelveValor(SQL)), "###,###,##0.00")


    ' DISPONIBLE

    SQL = "select sum(rbodalbaran_variedad.cantidad) as cantidad "
    SQL = SQL & " from variedades, rbodalbaran_variedad, rbodalbaran, productos  "
    SQL = SQL & " where variedades.codvarie = rbodalbaran_variedad.codvarie "
    SQL = SQL & " and rbodalbaran_variedad.numalbar = rbodalbaran.numalbar "
    SQL = SQL & " and variedades.codprodu = productos.codprodu "
    SQL = SQL & " and productos.codgrupo = 5 "
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    Consumido = CCur(DevuelveValor(SQL))
    
    Text3.Text = Format(CCur(ImporteFormateado(Text2.Text)) - Consumido, "###,###,##0.00")
    
    ' KILOS RECOLECTADOS DE ALMAZARA
    SQL = "select sum(rhisfruta.kilosnet) "
    SQL = SQL & " from variedades, rhisfruta, productos"
    SQL = SQL & " where rhisfruta.codvarie = variedades.codvarie "
    SQL = SQL & " and variedades.codprodu = productos.codprodu "
    SQL = SQL & " and productos.codgrupo = 5 "
    If cadWhere <> "" Then SQL = SQL & Replace(cadWhere, "rbodalbaran", "rhisfruta")
    
    Text5.Text = Format(CCur(DevuelveValor(SQL)), "###,###,##0.00")
    
    ' KILOS RECOLECTADOS DE BODEGA
    SQL = "select sum(rhisfruta.kilosnet) "
    SQL = SQL & " from variedades, rhisfruta, productos"
    SQL = SQL & " where rhisfruta.codvarie = variedades.codvarie "
    SQL = SQL & " and variedades.codprodu = productos.codprodu "
    SQL = SQL & " and productos.codgrupo = 6 "
    If cadWhere <> "" Then SQL = SQL & Replace(cadWhere, "rbodalbaran", "rhisfruta")
    
    Text6.Text = Format(CCur(DevuelveValor(SQL)), "###,###,##0.00")
    
End Sub

Private Sub CargarHidrantesSocio()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select rpozos.hidrante, rpozos.codparti, rpartida.nomparti, rpozos.poligono, rpozos.parcelas from rpozos, rpartida where "
    SQL = SQL & " rpozos.codparti = rpartida.codparti "
    
    
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView4.ColumnHeaders.Clear

    ListView4.ColumnHeaders.Add , , "Hidrante", 1200
    ListView4.ColumnHeaders.Add , , "Codigo", 800, 1
    ListView4.ColumnHeaders.Add , , "Partida", 2000
    ListView4.ColumnHeaders.Add , , "Poligono", 1200
    ListView4.ColumnHeaders.Add , , "Parcela", 2500
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView4.ListItems.Add
            
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Hidrante, "T")
        It.SubItems(1) = Format(RS!codparti, "000000")
        It.SubItems(2) = RS!nomparti
        It.SubItems(3) = DBLet(RS!Poligono, "T")
        It.SubItems(4) = DBLet(RS!parcelas, "T")
        It.Checked = False
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
    If Campo <> "" Then SituarHidranteSocio Campo
End Sub


Private Sub CargarHidrantesSocioFacturar()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select rpozos.hidrante, rpozos.codparti, rpartida.nomparti, rpozos.poligono, rpozos.parcelas from rpozos, rpartida where "
    SQL = SQL & " rpozos.codparti = rpartida.codparti and "
    SQL = SQL & " (rpozos.fechabaja is null or rpozos.fechabaja = '')"
    
    
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView13.ColumnHeaders.Clear

    ListView13.ColumnHeaders.Add , , "Hidrante", 1200
    ListView13.ColumnHeaders.Add , , "Codigo", 800, 1
    ListView13.ColumnHeaders.Add , , "Partida", 2000
    ListView13.ColumnHeaders.Add , , "Poligono", 1000
    ListView13.ColumnHeaders.Add , , "Parcela", 1500
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView13.ListItems.Add
            
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Hidrante, "T")
        It.SubItems(1) = Format(RS!codparti, "000000")
        It.SubItems(2) = RS!nomparti
        It.SubItems(3) = DBLet(RS!Poligono, "T")
        It.SubItems(4) = DBLet(RS!parcelas, "T")
        It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub

Private Sub CargarHidrantesCampo()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select rpozos.hidrante, rpozos.codsocio, rsocios.nomsocio, rpartida.nomparti, rpozos.poligono, rpozos.parcelas from rpozos, rpartida, rsocios where "
    SQL = SQL & " rpozos.codparti = rpartida.codparti and "
    SQL = SQL & " rpozos.codsocio = rsocios.codsocio "
    
    '[Monica]30/10/2013: he añadido esto para que no me mire la fecha de baja del contador
    If cadWHERE2 <> "1" Then
        SQL = SQL & " and (rpozos.fechabaja is null or rpozos.fechabaja = '')"
    End If
    
    
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView13.ColumnHeaders.Clear

    ListView13.ColumnHeaders.Add , , "Contador", 1000
    ListView13.ColumnHeaders.Add , , "Codigo", 800, 1
    ListView13.ColumnHeaders.Add , , "Socio", 1800
    ListView13.ColumnHeaders.Add , , "Partida", 1200
    ListView13.ColumnHeaders.Add , , "Poligono", 800
    ListView13.ColumnHeaders.Add , , "Parcela", 1000
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView13.ListItems.Add
            
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Hidrante, "T")
        It.SubItems(1) = Format(RS!Codsocio, "000000")
        It.SubItems(2) = RS!nomsocio
        It.SubItems(3) = RS!nomparti
        It.SubItems(4) = DBLet(RS!Poligono, "T")
        It.SubItems(5) = DBLet(RS!parcelas, "T")
        It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub

Private Sub CargarArchivos()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

Dim NomFic As String


    
    ListView14.ColumnHeaders.Clear

    ListView14.ColumnHeaders.Add , , "Nombre de Archivo", 2000
    ListView14.ColumnHeaders.Add , , "Path", 5500
    
    TotalArray = 0
    
    ' cargamos las cartas
    NomFic = Dir(App.Path & "\cartas\")  ' Recupera la primera entrada.

    Do While NomFic <> ""   ' Inicia el bucle.
       ' Ignora el directorio actual y el que lo abarca.
       If NomFic <> "." And NomFic <> ".." Then
             Set It = ListView14.ListItems.Add
                 
             It.Text = NomFic
             It.SubItems(1) = App.Path & "\cartas\" & NomFic
             It.Checked = False
             
             TotalArray = TotalArray + 1
             If TotalArray > 300 Then
                 TotalArray = 0
                 DoEvents
             End If
       End If
       NomFic = Dir   ' Obtiene siguiente entrada.
    Loop
    
End Sub



Private Sub CargarEntradasConError()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

Dim NomFic As String

    SQL = "select * from tmpexcel where codusu = " & vUsu.Codigo
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView15.ColumnHeaders.Clear

    ListView15.ColumnHeaders.Add , , "Nota", 1100
    ListView15.ColumnHeaders.Add , , "Variedad", 1000, 1
    ListView15.ColumnHeaders.Add , , "Socio", 1000
    ListView15.ColumnHeaders.Add , , "Campo", 1200
    ListView15.ColumnHeaders.Add , , "Error", 3600
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView15.ListItems.Add
            
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Numalbar, "T")
        It.SubItems(1) = Format(RS!codvarie, "000000")
        It.SubItems(2) = Format(RS!Codsocio, "000000")
        It.SubItems(3) = Format(RS!codcampo, "00000000")
        Select Case DBLet(RS!TipoEntr, "N")
            Case 0
                It.SubItems(4) = "Ya existe la nota de campo en el histórico"
            Case 1
                It.SubItems(4) = "No existe la variedad"
            Case 2
                It.SubItems(4) = "No existe el socio"
            Case 3
                It.SubItems(4) = "No existe el campo para el socio variedad"
        End Select
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
    
End Sub










Private Sub txtMante_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Public Function ObtenerSQLcomponentes(cadWhere As String) As String
'Obtiene la consulta SQL que selecciona los articulos con nº de serie
'agrupados por tipo de articulo
Dim SQL As String

    SQL = "Select distinct sserie.codtipar, nomtipar, count(numserie) as cantidad "
    SQL = SQL & "FROM sserie INNER JOIN stipar ON sserie.codtipar=stipar.codtipar "
    SQL = SQL & cadWhere
    SQL = SQL & " GROUP by codtipar "
    
    ObtenerSQLcomponentes = SQL
End Function


Private Sub SituarCampoSocio(Campo As Long)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim I As Integer
Dim ItmX As ListItem
    
' es lo mismo que lo de abajo para otro caso
'    Set itmX = ListView4.FindItem(CStr(Campo), lvwText, , lvwPartial)
'    If Not itmX Is Nothing Then
'        itmX.Checked = True
'        itmX.Selected = True
'        itmX.EnsureVisible
'        ListView4.SetFocus
'    End If
'
    For I = 1 To ListView4.ListItems.Count
        If Val(ListView4.ListItems(I).Text) = Val(Campo) Then
            ListView4.ListItems(I).Selected = True
            ListView4.ListItems(I).EnsureVisible
            ListView4.SetFocus
            Exit Sub
        End If
    Next I
    
End Sub


Private Sub SituarHidranteSocio(Campo As String)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim I As Integer
Dim ItmX As ListItem
    
    For I = 1 To ListView4.ListItems.Count
        If Val(ListView4.ListItems(I).Text) = Val(Campo) Then
            ListView4.ListItems(I).Selected = True
            ListView4.ListItems(I).EnsureVisible
            ListView4.SetFocus
            Exit Sub
        End If
    Next I
    
End Sub




Private Sub imgCheck1_Click(Index As Integer)
    If Index = 0 Then
        cmdDeselTodos_Click
    Else
        cmdSelTodos_Click
    End If
End Sub


Private Sub imgCheck3_Click(Index As Integer)
Dim b As Boolean
    'En el listview33
    b = Index = 1
    For TotalArray = 1 To ListView14.ListItems.Count
        ListView14.ListItems(TotalArray).Checked = b
        If (TotalArray Mod 50) = 0 Then DoEvents
    Next TotalArray
End Sub

Private Sub imgCheck4_Click(Index As Integer)
Dim I As Long

    If Index = 0 Then
        For I = 1 To ListView19.ListItems.Count
            ListView19.ListItems(I).Checked = False
        Next I
    Else
        For I = 1 To ListView19.ListItems.Count
            ListView19.ListItems(I).Checked = True
        Next I
    End If

End Sub


Private Sub ListView4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        cmdCamposSocio_Click (2)
    ElseIf KeyAscii = 27 Then 'ESC
        Unload Me
    End If
End Sub


Private Sub CargarListaTrabajadores(Cuadrilla As String)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select rcuadrilla_trabajador.codtraba, straba.nomtraba from rcuadrilla_trabajador, straba "
    SQL = SQL & " where rcuadrilla_trabajador.codcuadrilla = " & DBSet(Cuadrilla, "N")
    SQL = SQL & " and rcuadrilla_trabajador.codtraba = straba.codtraba "
    '[Monica]28/10/2015: cuando seleccionamos los trabajadores de la cuadrilla solo los que no tienen fecha de baja
    SQL = SQL & " and straba.fechabaja is null "
    
    SQL = SQL & " order by rcuadrilla_trabajador.numlinea"
   
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.ColumnHeaders.Clear
    
    ListView6.ColumnHeaders.Add , , "Código", 1000.0631
    ListView6.ColumnHeaders.Add , , "Trabajador", 4200.2522, 0
    
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView6.ListItems.Add
            
        It.Text = Format(DBLet(RS!CodTraba, "N"), "000000")
        It.SubItems(1) = DBLet(RS!nomtraba, "T")
        It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub



Private Sub CargarAlbaranes()

Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, "
    SQL = SQL & " rhisfruta_entradas.numnotac, rhisfruta_entradas.kilosnet, rhisfruta_entradas.imptrans "
    SQL = SQL & "  from rhisfruta, rhisfruta_entradas, variedades where "
    SQL = SQL & " rhisfruta.numalbar = rhisfruta_entradas.numalbar and "
    SQL = SQL & " rhisfruta.codvarie = variedades.codvarie "
    
    If cadWhere <> "" Then SQL = SQL & " and " & cadWhere
    
    SQL = SQL & " ORDER BY 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView4.ColumnHeaders.Clear
    
    ListView4.ColumnHeaders.Add , , "Albarán", 900
    ListView4.ColumnHeaders.Add , , "Fecha", 1100
    ListView4.ColumnHeaders.Add , , "Variedad", 1500
    ListView4.ColumnHeaders.Add , , "Campo", 1000
    ListView4.ColumnHeaders.Add , , "Nota", 1000, 1
    ListView4.ColumnHeaders.Add , , "Kilos", 1200, 1
    ListView4.ColumnHeaders.Add , , "Importe", 1200, 1
        
        
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView4.ListItems.Add
        
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Numalbar, "N")
        It.SubItems(1) = DBLet(RS!Fecalbar, "F")
        It.SubItems(2) = RS!nomvarie
        It.SubItems(3) = Format(DBLet(RS!codcampo, "N"), "00000000")
        It.SubItems(4) = Format(DBLet(RS!numnotac, "N"), "0000000")
        It.SubItems(5) = DBLet(RS!KilosNet, "N")
        It.SubItems(6) = DBLet(RS!ImpTrans, "N")
        
        It.Checked = False
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close

End Sub


Private Sub CargarAlbaranesSocio()

Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem



    SQL = "select rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, variedades.nomvarie, rcampos.nrocampo, "
    SQL = SQL & "  rpartida.nomparti, rcampos.poligono, rcampos.parcela, rhisfruta.kilosnet "
    SQL = SQL & "  from rhisfruta, variedades, rpartida, rcampos where "
    SQL = SQL & " rhisfruta.codvarie = variedades.codvarie "
    SQL = SQL & " and rhisfruta.codcampo = rcampos.codcampo "
    SQL = SQL & " and rcampos.codparti = rpartida.codparti "
    
    If cadWhere <> "" Then SQL = SQL & " and " & cadWhere
    
    SQL = SQL & " ORDER BY 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView4.ColumnHeaders.Clear
    
    ListView4.ColumnHeaders.Add , , "Albarán", 900
    ListView4.ColumnHeaders.Add , , "Fecha", 1100
    ListView4.ColumnHeaders.Add , , "Variedad", 1500
    ListView4.ColumnHeaders.Add , , "Campo", 800
    ListView4.ColumnHeaders.Add , , "Partida", 1500
    ListView4.ColumnHeaders.Add , , "Pol.", 600
    ListView4.ColumnHeaders.Add , , "Parc", 600
    ListView4.ColumnHeaders.Add , , "Kilos", 1000, 1
    
    ListView4.Checkboxes = True
        
        
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView4.ListItems.Add
        
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Numalbar, "N")
        It.SubItems(1) = DBLet(RS!Fecalbar, "F")
        It.SubItems(2) = RS!nomvarie
        It.SubItems(3) = Format(DBLet(RS!NroCampo, "N"), "000000")
        It.SubItems(4) = RS!nomparti
        It.SubItems(5) = RS!Poligono
        It.SubItems(6) = RS!Parcela
        It.SubItems(7) = DBLet(RS!KilosNet, "N")
        
        It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close

End Sub

Private Sub CargarAlbaranesLiquidados()

Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, "
    SQL = SQL & " rhisfruta.kilosnet "
    SQL = SQL & "  from rhisfruta, variedades where "
    SQL = SQL & " rhisfruta.codvarie = variedades.codvarie "
    
    If cadWhere <> "" Then SQL = SQL & " and " & cadWhere
    
    SQL = SQL & " ORDER BY 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView12.ColumnHeaders.Clear
    
    ListView12.ColumnHeaders.Add , , "Albarán", 900
    ListView12.ColumnHeaders.Add , , "Fecha", 1100
    ListView12.ColumnHeaders.Add , , "Variedad", 1500
    ListView12.ColumnHeaders.Add , , "Campo", 1000
    ListView12.ColumnHeaders.Add , , "Kilos", 1200, 1
        
        
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView12.ListItems.Add
        
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Numalbar, "N")
        It.SubItems(1) = DBLet(RS!Fecalbar, "F")
        It.SubItems(2) = RS!nomvarie
        It.SubItems(3) = Format(DBLet(RS!codcampo, "N"), "00000000")
        It.SubItems(4) = DBLet(RS!KilosNet, "N")
        
        It.Checked = False
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close

End Sub



Private Sub CargarPlagas()

Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem



    SQL = "select rincidencia.codincid,nomincid,case tipincid when 0 then ""LEVE"" when 1 then ""GRAVE"" when 2 then ""MUY GRAVE"" end as tipoincid"
    SQL = SQL & "  from rincidencia "
    
    If cadWhere <> "" Then SQL = SQL & " where (1=1)" & cadWhere
    
    SQL = SQL & " ORDER BY 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.ColumnHeaders.Clear
    
    ListView6.ColumnHeaders.Add , , "Código", 1000
    ListView6.ColumnHeaders.Add , , "Plaga", 2500
    ListView6.ColumnHeaders.Add , , "Tipo", 2500
    
    ListView6.Checkboxes = True
        
        
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView6.ListItems.Add
        
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!codincid, "N")
        It.SubItems(1) = RS!nomincid
        It.SubItems(2) = RS!tipoincid
        
        It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close

End Sub


Private Sub CargarAportaciones()

Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem



    SQL = "select rtipoapor.codaport,nomaport "
    SQL = SQL & "  from rtipoapor "
    
    If cadWhere <> "" Then SQL = SQL & " where (1=1)" & cadWhere
    
    SQL = SQL & " ORDER BY 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.ColumnHeaders.Clear
    
    ListView6.ColumnHeaders.Add , , "Código", 1000
    ListView6.ColumnHeaders.Add , , "Descripción", 2500
    
    ListView6.Checkboxes = True
        
        
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView6.ListItems.Add
        
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Codaport, "N")
        It.SubItems(1) = RS!nomaport
        
        It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close

End Sub







Private Sub CargarNotasSinTaraSalida()

Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem



    SQL = "select rentradas.numnotac, rentradas.fechaent, rentradas.horaentr, rentradas.codvarie, variedades.nomvarie, rentradas.codsocio, rsocios.nomsocio "
    SQL = SQL & "  from rentradas, variedades, rsocios "
    
    ' siempre hay cadwhere pq sino no entro en mostrar entradas
    SQL = SQL & " where rentradas.codsocio = rsocios.codsocio and rentradas.codvarie = variedades.codvarie and " & cadWhere
    
    SQL = SQL & " ORDER BY 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView10.ColumnHeaders.Clear
    
    ListView10.ColumnHeaders.Add , , "Nota", 900
    ListView10.ColumnHeaders.Add , , "Fecha", 1100
    ListView10.ColumnHeaders.Add , , "Hora", 900
    ListView10.ColumnHeaders.Add , , "Código", 750
    ListView10.ColumnHeaders.Add , , "Variedad", 1300
    ListView10.ColumnHeaders.Add , , "Código", 750
    ListView10.ColumnHeaders.Add , , "Socio", 2050
    
    
    ListView10.Checkboxes = False
    
        
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView10.ListItems.Add
        
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!numnotac, "N")
        It.SubItems(1) = Format(RS!FechaEnt, "dd/mm/yyyy")
        It.SubItems(2) = Format(RS!horaentr, "hh:mm:ss")
        It.SubItems(3) = Format(RS!codvarie, "000000")
        It.SubItems(4) = RS!nomvarie
        It.SubItems(5) = Format(RS!Codsocio, "000000")
        It.SubItems(6) = RS!nomsocio
        
        It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close

End Sub


Private Sub lw1_DblClick()
    CmdAcepEmpresas_Click
End Sub



Private Sub Text4_LostFocus()
    If Not PerderFocoGnral(Text1, 3) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    If Text4.Text <> "" Then
        Text4.Text = Format(Text4.Text, "000000")
        cadWhere = " and rcampos.codsocio = " & DBSet(Text4.Text, "N")
        ListView4.ListItems.Clear
        CargarCamposSocio 1
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text4_GotFocus()
    ConseguirFoco Text4, 3
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub CargarAlbaranesBodegaSinTarar()

Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem



    SQL = "select rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta_entradas.horaentr, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codsocio, rsocios.nomsocio "
    SQL = SQL & "  from rhisfruta, rhisfruta_entradas, variedades, rsocios "
    
    ' siempre hay cadwhere pq sino no entro en mostrar entradas
    SQL = SQL & " where rhisfruta.codsocio = rsocios.codsocio and rhisfruta.codvarie = variedades.codvarie and rhisfruta.numalbar = rhisfruta_entradas.numalbar and " & cadWhere
    
    SQL = SQL & " ORDER BY 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView10.ColumnHeaders.Clear
    
    ListView10.ColumnHeaders.Add , , "Albarán", 900
    ListView10.ColumnHeaders.Add , , "Fecha", 1100
    ListView10.ColumnHeaders.Add , , "Hora", 900
    ListView10.ColumnHeaders.Add , , "Código", 750
    ListView10.ColumnHeaders.Add , , "Variedad", 1300
    ListView10.ColumnHeaders.Add , , "Código", 750
    ListView10.ColumnHeaders.Add , , "Socio", 2050
    
    
    ListView10.Checkboxes = False
    
        
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView10.ListItems.Add
        
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Numalbar, "N")
        It.SubItems(1) = Format(RS!Fecalbar, "dd/mm/yyyy")
        It.SubItems(2) = Format(RS!horaentr, "hh:mm:ss")
        It.SubItems(3) = Format(RS!codvarie, "000000")
        It.SubItems(4) = RS!nomvarie
        It.SubItems(5) = Format(RS!Codsocio, "000000")
        It.SubItems(6) = RS!nomsocio
        
        It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close

End Sub

Private Sub CargarFacturasVCsinEntradas()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem


    SQL = "select rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu, rfactsoc.baseimpo  "
    SQL = SQL & "  from rfactsoc "
    
    ' siempre hay cadwhere pq sino no entro en mostrar entradas
    SQL = SQL & " where  " & cadWhere
    
    SQL = SQL & " ORDER BY 1, 2, 3 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView11.ColumnHeaders.Clear
    
    ListView11.ColumnHeaders.Add , , "Tipo", 900
    ListView11.ColumnHeaders.Add , , "Fecha", 1100
    ListView11.ColumnHeaders.Add , , "Factura", 1100
    ListView11.ColumnHeaders.Add , , "Base Imponible Factura", 2200, 1
    
    
    ListView11.Checkboxes = False
    
        
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView11.ListItems.Add
        
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!CodTipom, "T")
        It.SubItems(1) = Format(RS!fecfactu, "dd/mm/yyyy")
        It.SubItems(2) = Format(RS!numfactu, "0000000")
        It.SubItems(3) = Format(RS!baseimpo, "###,###,##0.00")
        
        It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close


End Sub



Private Sub CargarListaSituaciones()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select rsituacion.codsitua, rsituacion.nomsitua from rsituacion where (1=1) "
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.ColumnHeaders.Clear
    
    ListView6.ColumnHeaders.Add , , "Código", 1500.0631
    ListView6.ColumnHeaders.Add , , "Situación", 4700.2522
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView6.ListItems.Add
            
        It.Text = Format(DBLet(RS!codsitua, "N"), "000")
        It.SubItems(1) = DBLet(RS!nomsitua, "T")
        It.Checked = False
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub


Private Sub CargaEmpresas()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim It As ListItem
Dim Cad As String
Dim Encontrado
Dim NomFic As String
Dim Sql1 As String



    'Cargamos las prohibidas
'    Prohibidas = DevuelveProhibidas
    
    lw1.ColumnHeaders.Clear

    lw1.ColumnHeaders.Add , , "", 2100
    lw1.ColumnHeaders.Add , , "", 4000
    
    
    'Cargamos las empresas
    Set RS = New ADODB.Recordset
    
    ' Primero meto la campaña actual
    SQL = "select * from usuarios.empresasariagro where ariagro = " & DBSet(vUsu.CadenaConexion, "T")
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        Cad = "|" & RS!codempre & "|"
        Cad = RS!nomempre
        Set It = lw1.ListItems.Add()
        
        It.Text = Cad
        It.SubItems(1) = RS!nomresum
        Cad = RS!Ariagro & "|" & RS!nomresum & "|" & RS!Usuario & "|" & RS!Pass & "|"
        It.Tag = Cad
        It.ToolTipText = RS!Ariagro
    End If
    Set RS = Nothing
    
    
    ' Ahora busco cual es la campaña anterior
    Set RS = New ADODB.Recordset
    
    SQL = "Select * from usuarios.empresasariagro where ariagro <> " & DBSet(vUsu.CadenaConexion, "T")
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Encontrado = False
    
    While Not RS.EOF And Not Encontrado
        If AbrirConexionCampAnterior(DBLet(RS!Ariagro, "T")) Then
            Sql1 = "select * from empresas "
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql1, ConnCAnt, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            If DBLet(RS1!FechaFin, "F") = (CDate(vParam.FecIniCam) - 1) Then
                    Encontrado = True
                    Cad = "|" & RS!codempre & "|"
                    Cad = RS!nomempre
                    Set It = lw1.ListItems.Add()
                    
                    It.Text = Cad
                    It.SubItems(1) = RS!nomresum
                    Cad = RS!Ariagro & "|" & RS!nomresum & "|" & RS!Usuario & "|" & RS!Pass & "|"
                    It.Tag = Cad
                    It.ToolTipText = RS!Ariagro
            End If
'        It.SmallIcon = 1
        End If
        CerrarConexionCampAnterior
        
        RS.MoveNext
    Wend
    RS.Close

   
End Sub


Private Sub BuscarDiferencias()
Dim SQL As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim Nregs As Integer

    On Error GoTo eBuscarDiferencias


    '[Monica]30/10/2013: cogemos tambien el nro de orden para comprobar la toma con indefa


    SQL = "select hidrante, poligono, parcelas, hanegada, codsocio, nroorden from rpozos where length(hidrante) = 6 and cast(hidrante as unsigned) "
    SQL = SQL & " and fechabaja is null order by 1 "
    
    Nregs = TotalRegistrosConsulta(SQL)
    CargarProgres Pb1, Nregs

    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadContadores = ""
    
    While Not RS.EOF And Not PulsadoSalir
        IncrementarProgres Pb1, 1
        
        Label13.Caption = "Procesando contador: " & DBLet(RS!Hidrante, "T")
        DoEvents
    
'        If RS!Hidrante = "023506" Then
'
'            MsgBox ""
'        End If
    
    
        Sql2 = "select poligono, parcelas, hanegadas, socio_revisado, toma from rae_visitas_hidtomas where sector = '" & CInt(Mid(RS!Hidrante, 1, 2))
        '[Monica]18/07/2013: cambio toma por salida_tch
                                                                                    '[Monica]27/01/2014: lo cambio a numerico
        Sql2 = Sql2 & "' and hidrante = '" & CInt(Mid(RS!Hidrante, 3, 2)) & "' and salida_tch = " & CInt(Mid(RS!Hidrante, 5, 2)) & ""
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs2.EOF Then
            If Trim(DBLet(RS!Poligono, "T")) <> Trim(DBLet(Rs2!Poligono, "T")) Or _
               Mid(Trim(DBLet(RS!parcelas, "T")), 1, 25) <> Mid(Trim(DBLet(Rs2!parcelas, "T")), 1, 25) Or _
               Int(ComprobarCero(DBLet(RS!hanegada, "N"))) <> Int(Round2(ComprobarCero(DBLet(Rs2!Hanegadas, "N")), 4)) Or _
               (DBLet(RS!Codsocio, "N") <> ComprobarCero(DBLet(Rs2!socio_revisado, "N")) And DBLet(Rs2!socio_revisado, "N") <> 0) Or _
               (ComprobarCero(DBLet(RS!nroorden, "N")) Mod 100) <> ComprobarCero(DBLet(Rs2!toma, "N")) Then
                CadContadores = CadContadores & DBSet(RS!Hidrante, "T") & ","
            End If
        End If
        Set Rs2 = Nothing
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    If PulsadoSalir Then CadContadores = ""
    
    RaiseEvent DatoSeleccionado(CadContadores)
    CmdCanDif_Click
    
eBuscarDiferencias:
    If Err.Number <> 0 Then
        CadContadores = ""
        MuestraError Err.Number, "Buscar Diferencias", Err.Description
    End If
End Sub



Private Sub CargarAlbaranesPdtesFacturar()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim I As Integer


    SQL = "select rhisfruta.numalbar, rhisfruta.fecalbar, variedades.nomvarie, rhisfruta.kilosnet from rhisfruta, variedades  where rhisfruta.codvarie = variedades.codvarie "
    If cadWhere <> "" Then SQL = SQL & " and " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.ColumnHeaders.Clear
    
    ListView6.ColumnHeaders.Add , , "Código", 1500.0631
    ListView6.ColumnHeaders.Add , , "Fecha", 1200.2522
    ListView6.ColumnHeaders.Add , , "Variedad", 2000.2522
    ListView6.ColumnHeaders.Add , , "Kilos", 1700.2522, 1
    
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView6.ListItems.Add
            
        It.Text = Format(DBLet(RS!Numalbar, "N"), "000000")
        It.SubItems(1) = DBLet(RS!Fecalbar, "F")
        It.SubItems(2) = DBLet(RS!nomvarie, "F")
        It.SubItems(3) = DBLet(RS!KilosNet, "N")
        It.Checked = False
        
        If EstaFacturado(RS!Numalbar) Then
            It.ForeColor = vbRed
            For I = 1 To 3
                It.ListSubItems(I).ForeColor = vbRed
            Next I
        End If
        
        
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub



Private Sub CargarAnticiposSinDescontar()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim I As Integer


    SQL = cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.ColumnHeaders.Clear
    
    ListView6.ColumnHeaders.Add , , "Número", 1500.0631
    ListView6.ColumnHeaders.Add , , "Fecha", 1200.2522
    ListView6.ColumnHeaders.Add , , "Variedad", 2000.2522
    ListView6.ColumnHeaders.Add , , "Base Imponible", 1700.2522, 1
    
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView6.ListItems.Add
            
        It.Text = DBLet(RS.Fields(0).Value, "T")
        It.SubItems(1) = DBLet(RS.Fields(1).Value, "F")
        It.SubItems(2) = DBLet(RS.Fields(2).Value, "F")
        It.SubItems(3) = DBLet(RS.Fields(3).Value, "N")
        It.Checked = False
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub

Private Sub CargarFechasSinDescontar()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim I As Integer


    SQL = "select distinct fecfactu from rfactsoc_variedad where fecfactu in (" & cadWhere & ") order by fecfactu "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.ColumnHeaders.Clear
    
    ListView6.ColumnHeaders.Add , , "Fecha", 1500.2522
    
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView6.ListItems.Add
            
        It.Text = DBLet(RS.Fields(0).Value, "F")
        It.Checked = True
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
    
    Set RS = Nothing
End Sub






Private Sub CargarListaCamposSinPrecioZona()
'Muestra la lista de campos en zonas sin precio
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = Cadena 'cadwhere ya le pasamos toda la SQL
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        ListView1.Height = 3900
        ListView1.Width = 7200
        ListView1.Left = 500
        ListView1.Top = 700

        'Los encabezados
        ListView1.ColumnHeaders.Clear

'        ListView1.ColumnHeaders.Add , , "Nº Campo", 1000
'        ListView1.ColumnHeaders.Add , , "Socio", 1100, 1
'        ListView1.ColumnHeaders.Add , , "Nombre", 3800

        ListView1.ColumnHeaders.Add , , "Zona", 1300
        ListView1.ColumnHeaders.Add , , "Nombre", 5800
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(RS!codzonas, "000")
            ItmX.SubItems(1) = RS!nomzonas
            
'            ItmX.SubItems(1) = Format(RS!Codsocio, "000000")
'            ItmX.SubItems(2) = RS!nomsocio
'            ItmX.SubItems(3) = Format(RS!codzonas, "000")
            
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

    Set ItmX = Nothing
    
ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub CargarContadoresANoFacturar()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim Consumido As Currency

    'CONTADORES CON CONSUMO INFERIOR AL MINIMO

    SQL = "select rpozos.hidrante, rsocios.nomsocio, rpozos.consumo as consumo "
    SQL = SQL & " from rpozos inner join rsocios on rpozos.codsocio = rsocios.codsocio "
    SQL = SQL & " where consumo < " & DBSet(vParamAplic.ConsumoMinPOZ, "N")
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    SQL = SQL & " order by 1,2 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView16.ColumnHeaders.Clear
    
    ListView16.ColumnHeaders.Add , , "Contador", 850
    ListView16.ColumnHeaders.Add , , "Socio", 2640 ' 1450 '1740
    ListView16.ColumnHeaders.Add , , "Consumo", 1000, 1
    
    
    While Not RS.EOF
        Set It = ListView16.ListItems.Add
            
        It.Text = DBLet(RS!Hidrante, "T")
        It.SubItems(1) = DBLet(RS!nomsocio, "T")
        It.SubItems(2) = Format(DBLet(RS!Consumo, "N"), "########0")
        It.Checked = False
        
        RS.MoveNext
    Wend
    RS.Close
    
    
    'CONTADORES CON CONSUMO SUPERIOR AL MAXIMO

    SQL = "select rpozos.hidrante, rsocios.nomsocio, rpozos.consumo as consumo "
    SQL = SQL & " from rpozos inner join rsocios on rpozos.codsocio = rsocios.codsocio "
    SQL = SQL & " where consumo > " & DBSet(vParamAplic.ConsumoMaxPOZ, "N")
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    SQL = SQL & " order by 1,2 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView17.ColumnHeaders.Clear
    
    ListView17.ColumnHeaders.Add , , "Contador", 850
    ListView17.ColumnHeaders.Add , , "Socio", 2640 ' 1450 '1740
    ListView17.ColumnHeaders.Add , , "Consumo", 1000, 1
    
    While Not RS.EOF
        Set It = ListView17.ListItems.Add
            
        It.Text = DBLet(RS!Hidrante, "T")
        It.SubItems(1) = DBLet(RS!nomsocio, "T")
        It.SubItems(2) = Format(DBLet(RS!Consumo, "N"), "########0")
        It.Checked = False
        
        RS.MoveNext
    Wend
    RS.Close
    
End Sub


Private Sub CargarListaTransportistas()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select rtransporte.codtrans, rtransporte.nomtrans, rtransporte.matricula from rtransporte "
    SQL = SQL & " where (1=1) "
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.ColumnHeaders.Clear
    
    ListView6.ColumnHeaders.Add , , "Código", 1000.0631
    ListView6.ColumnHeaders.Add , , "Nombre", 3200.2522, 0
    ListView6.ColumnHeaders.Add , , "Matrícula", 1799.9371, 0
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView6.ListItems.Add
            
        It.Text = DBLet(RS!codTrans, "T")
        It.SubItems(1) = DBLet(RS!nomtrans, "T")
        It.SubItems(2) = DBLet(RS!Matricula, "T")
        
        It.Checked = False
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub

Private Sub CmdAcep_Click()
    If Check1.Value = 0 And Check2.Value = 0 Then
        MsgBox "Debe seleccionar una forma de pago.", vbExclamation
        PonerFocoChk Check1
        Exit Sub
    Else
        Cadena = Check1.Value
        RaiseEvent DatoSeleccionado(Cadena)
        Unload Me
    
    End If
End Sub

Private Sub Text7_LostFocus()
    RaiseEvent DatoSeleccionado(Text7.Text)
    Unload Me
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        Text7_LostFocus
    ElseIf KeyAscii = 27 Then 'ESC
            Text7_LostFocus
    End If
End Sub

Private Sub CargarMatriculas()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select rtransporte.matricula, rtransporte.contador from rtransporte where "
    
    If cadWhere <> "" Then SQL = SQL & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView18.ColumnHeaders.Clear

    ListView18.ColumnHeaders.Add , , "Matrícula", 2200
    ListView18.ColumnHeaders.Add , , "Contador", 2000, 1
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView18.ListItems.Add
            
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Matricula, "T")
        It.SubItems(1) = Format(RS!Contador, "0000000")
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub



Private Sub CargarFacturasPozos(sColumna1 As String, sColumna2 As String)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select * from tmpinformes where codusu =" & vUsu.Codigo
    If sColumna1 <> "" Or sColumna2 <> "" Then
        SQL = SQL & " order by "
        If sColumna1 <> "" Then SQL = SQL & sColumna1
        If sColumna2 <> "" Then SQL = SQL & "," & sColumna2
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView19.ColumnHeaders.Clear

    ListView19.ColumnHeaders.Add , , "Tipo", 800
    If columna = 1 Then
        If Orden = 0 Then
            ListView19.ColumnHeaders.Add , , "NºRecibo", 1000, 1
        Else
            ListView19.ColumnHeaders.Add , , "NºRecibo v", 1000, 1
        End If
    Else
        If Columna2 = 1 Then
            If Orden2 = 0 Then
                ListView19.ColumnHeaders.Add , , "NºRecibo", 1000, 1
            Else
                ListView19.ColumnHeaders.Add , , "NºRecibo v", 1000, 1
            End If
        Else
            ListView19.ColumnHeaders.Add , , "NºRecibo", 1000, 1
        End If
    End If
    If columna = 2 Then
        If Orden = 0 Then
            ListView19.ColumnHeaders.Add , , "Fecha", 1200, 0
        Else
            ListView19.ColumnHeaders.Add , , "Fecha v", 1200, 0
        End If
    Else
        If Columna2 = 2 Then
            If Orden2 = 0 Then
                ListView19.ColumnHeaders.Add , , "Fecha", 1200, 0
            Else
                ListView19.ColumnHeaders.Add , , "Fecha v", 1200, 0
            End If
        Else
            ListView19.ColumnHeaders.Add , , "Fecha", 1200, 0
        End If
    End If
    If columna = 3 Then
        If Orden = 0 Then
            ListView19.ColumnHeaders.Add , , "Socio", 900, 1
        Else
            ListView19.ColumnHeaders.Add , , "Socio v", 900, 1
        End If
    Else
        If Columna2 = 3 Then
            If Orden2 = 0 Then
                ListView19.ColumnHeaders.Add , , "Socio", 900, 1
            Else
                ListView19.ColumnHeaders.Add , , "Socio v", 900, 1
            End If
        Else
            ListView19.ColumnHeaders.Add , , "Socio", 900, 1
        End If
    End If
    If columna = 4 Then
        If Orden = 0 Then
            ListView19.ColumnHeaders.Add , , "Nombre", 3000, 0
        Else
            ListView19.ColumnHeaders.Add , , "Nombre v", 3000, 0
        End If
    Else
        If Columna2 = 4 Then
            If Orden2 = 0 Then
                ListView19.ColumnHeaders.Add , , "Nombre", 3000, 0
            Else
                ListView19.ColumnHeaders.Add , , "Nombre v", 3000, 0
            End If
        Else
            ListView19.ColumnHeaders.Add , , "Nombre", 3000, 0
        End If
    End If
    If columna = 5 Then
        If Orden = 0 Then
            ListView19.ColumnHeaders.Add , , "Total", 1500, 1
        Else
            ListView19.ColumnHeaders.Add , , "Total v", 1500, 1
        End If
    Else
        If Columna2 = 5 Then
            If Orden2 = 0 Then
                ListView19.ColumnHeaders.Add , , "Total", 1500, 1
            Else
                ListView19.ColumnHeaders.Add , , "Total v", 1500, 1
            End If
        Else
            ListView19.ColumnHeaders.Add , , "Total", 1500, 1
        End If
    End If
    If columna = 6 Then
        If Orden = 0 Then
            ListView19.ColumnHeaders.Add , , "Cobrado", 800, 0
        Else
            ListView19.ColumnHeaders.Add , , "Cobrado v", 800, 0
        End If
    Else
        If Columna2 = 6 Then
            If Orden2 = 0 Then
                ListView19.ColumnHeaders.Add , , "Cobrado", 800, 0
            Else
                ListView19.ColumnHeaders.Add , , "Cobrado v", 800, 0
            End If
        Else
            ListView19.ColumnHeaders.Add , , "Cobrado", 800, 0
        End If
    End If
    
    ListView19.ListItems.Clear
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView19.ListItems.Add
            
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!Nombre1, "T")
        It.SubItems(1) = Format(DBLet(RS!importe1), "0000000")
        It.SubItems(2) = DBLet(RS!fecha1, "F")
        It.SubItems(3) = Format(RS!Codigo1, "000000")
        It.SubItems(4) = DBLet(RS!Nombre2)
        It.SubItems(5) = DBLet(RS!importe2, "###,###,##0.00")
        If DBLet(RS!campo1) = 1 Then
            It.SubItems(6) = "Cobrado"
        Else
            It.SubItems(6) = " "
        End If
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub

Private Sub ListView19_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim campo2 As Integer

    Select Case ColumnHeader
        Case "NºRecibo", "NºRecibo v"
            campo2 = 1
        Case "Fecha", "Fecha v"
            campo2 = 2
        Case "Socio", "Socio v"
            campo2 = 3
        Case "Nombre", "Nombre v"
            campo2 = 4
        Case "Total", "Total v"
            campo2 = 5
        Case "Cobrado", "Cobrado v"
            campo2 = 6
    End Select




    If nomColumna = "" Or PrimerCampo = campo2 Then
        Select Case ColumnHeader
            Case "NºRecibo", "NºRecibo v"
                nomColumna = "importe1"
                campo2 = 1
            Case "Fecha", "Fecha v"
                nomColumna = "fecha1"
                campo2 = 2
            Case "Socio", "Socio v"
                nomColumna = "codigo1"
                campo2 = 3
            Case "Nombre", "Nombre v"
                nomColumna = "nombre2"
                campo2 = 4
            Case "Total", "Total v"
                nomColumna = "importe2"
                campo2 = 5
            Case "Cobrado", "Cobrado v"
                nomColumna = "campo1"
                campo2 = 6
        End Select
        If PrimerCampo = 0 Then PrimerCampo = campo2
        
        If campo2 = columna Then
            If Orden = lvwAscending Then
                nomColumna = nomColumna & " DESC"
                Orden = lvwDescending
            Else
                Orden = lvwAscending
            End If
'        Else
'            nomColumna = nomColumna & " DESC"
'            Orden = lvwDescending
        End If
    
        Select Case ColumnHeader
            Case "NºRecibo", "NºRecibo v"
                columna = 1
            Case "Fecha", "Fecha v"
                columna = 2
            Case "Socio", "Socio v"
                columna = 3
            Case "Nombre", "Nombre v"
                columna = 4
            Case "Total", "Total v"
                columna = 5
            Case "Cobrado", "Cobrado v"
                columna = 6
        End Select
    Else
        Select Case ColumnHeader
            Case "NºRecibo", "NºRecibo v"
                nomColumna2 = "importe1"
                campo2 = 1
            Case "Fecha", "Fecha v"
                nomColumna2 = "fecha1"
                campo2 = 2
            Case "Socio", "Socio v"
                nomColumna2 = "codigo1"
                campo2 = 3
            Case "Nombre", "Nombre v"
                nomColumna2 = "nombre2"
                campo2 = 4
            Case "Total", "Total v"
                nomColumna2 = "importe2"
                campo2 = 5
            Case "Cobrado", "Cobrado v"
                nomColumna2 = "campo1"
                campo2 = 6
        End Select
        
        If campo2 = Columna2 Then
            If Orden2 = lvwAscending Then
                nomColumna2 = nomColumna2 & " DESC"
                Orden2 = lvwDescending
            Else
                Orden2 = lvwAscending
            End If
'        Else
'            nomColumna2 = nomColumna2 & " DESC"
'            Orden2 = lvwDescending
        End If
    
        Select Case ColumnHeader
            Case "NºRecibo", "NºRecibo v"
                Columna2 = 1
            Case "Fecha", "Fecha v"
                Columna2 = 2
            Case "Socio", "Socio v"
                Columna2 = 3
            Case "Nombre", "Nombre v"
                Columna2 = 4
            Case "Total", "Total v"
                Columna2 = 5
            Case "Cobrado", "Cobrado v"
                Columna2 = 6
        End Select
    
    
    End If
    CargarFacturasPozos nomColumna, nomColumna2

End Sub


'*******************
Private Sub Text8_GotFocus(Index As Integer)
    ConseguirFoco Text8(Index), 3
End Sub

Private Sub Text8_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(Text8(Index), 3) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Select Case Index
        Case 1 'SOCIO
            If PonerFormatoEntero(Text8(Index)) Then
                Text9(Index).Text = PonerNombreDeCod(Text8(Index), "rsocios", "nomsocio")
                If Text9(Index).Text = "" Then
                    cadMen = "No existe el Socio " & Text8(Index).Text & ". Reintroduzca."
                    MsgBox cadMen, vbExclamation
                    PonerFoco Text8(Index)
                End If
            Else
                Text9(Index).Text = ""
            End If
            
        
        Case 2 'VARIEDAD
            If PonerFormatoEntero(Text8(Index)) Then
                Text9(Index).Text = PonerNombreDeCod(Text8(Index), "variedades", "nomvarie")
                If Text9(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text8(Index).Text
                    MsgBox cadMen, vbExclamation
                    PonerFoco Text8(Index)
                End If
            Else
                Text9(Index).Text = ""
            End If
                
        Case 3 'PARTIDA
            If PonerFormatoEntero(Text8(Index)) Then
                Text9(Index).Text = PonerNombreDeCod(Text8(Index), "rpartida", "nomparti", "codparti", "N")
                If Text9(Index).Text = "" Then
                    cadMen = "No existe la Partida: " & Text8(Index).Text
                    MsgBox cadMen, vbExclamation
                    PonerFoco Text8(Index)
                End If
            Else
                Text9(Index).Text = ""
            End If
            
        Case 4, 5 'poligono y parcela
            PonerFormatoEntero Text8(Index)
        
                
    End Select
End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 1: KEYBusqueda KeyAscii, 1 'socio
            Case 2: KEYBusqueda KeyAscii, 2 'variedad
            Case 3: KEYBusqueda KeyAscii, 3 'partida
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text8_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub



Private Sub CargarPrevisualizacion()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim It As ListItem

    SQL = "select * from tmpinformes where codusu =" & vUsu.Codigo
    SQL = SQL & " order by importe1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView20.ColumnHeaders.Clear

    ListView20.ColumnHeaders.Add , , "Nota", 1000
    ListView20.ColumnHeaders.Add , , "Fecha", 1300, 1
    ListView20.ColumnHeaders.Add , , "Socio", 1300, 1
    ListView20.ColumnHeaders.Add , , "Nombre", 4000
    ListView20.ColumnHeaders.Add , , "Neto", 1500, 1
    
    
'    ListView20.ColumnHeaders.Add , , "Variedad", 1300, 1
'    ListView20.ColumnHeaders.Add , , "Poligono", 1500, 1
'    ListView20.ColumnHeaders.Add , , "Parcela", 1200, 1
'    ListView20.ColumnHeaders.Add , , "SubParcela", 1500, 1
    
    
    ListView20.ListItems.Clear
    
    TotalArray = 0
    While Not RS.EOF
        Set It = ListView20.ListItems.Add
            
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(RS!importe1, "N")
        It.SubItems(1) = DBLet(RS!fecha1, "F")
        It.SubItems(2) = Format(RS!importe2, "000000")
        It.SubItems(3) = DevuelveValor("select nomsocio from rsocios where codsocio = " & DBSet(RS!importe2, "N"))
        It.SubItems(4) = Format(RS!importeb1, "###,##0")
        
'        It.SubItems(3) = Format(DBLet(Rs!importe3), "000000")
'        It.SubItems(4) = DBLet(Rs!importe4)
'        It.SubItems(5) = DBLet(Rs!importe5)
'        It.SubItems(6) = DBLet(Rs!Nombre1)
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    
End Sub

