VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmListAnticipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6750
   Icon            =   "frmListAnticipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAnticipos 
      Height          =   6630
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Frame FrameTipo 
         Height          =   705
         Left            =   3300
         TabIndex        =   334
         Top             =   270
         Visible         =   0   'False
         Width           =   2715
         Begin VB.CheckBox Check1 
            Caption         =   "Agroseguro"
            Height          =   195
            Index           =   26
            Left            =   1410
            TabIndex        =   338
            Top             =   420
            Value           =   1  'Checked
            Width           =   1245
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VCampo"
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   337
            Top             =   420
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox Check1 
            Caption         =   "PIntegrado"
            Height          =   195
            Index           =   24
            Left            =   1410
            TabIndex        =   336
            Top             =   180
            Value           =   1  'Checked
            Width           =   1155
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Normales"
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   335
            Top             =   180
            Value           =   1  'Checked
            Width           =   1035
         End
      End
      Begin VB.TextBox txtcodigo 
         Enabled         =   0   'False
         Height          =   435
         Index           =   68
         Left            =   1650
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   5160
         Visible         =   0   'False
         Width           =   4275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "VR"
         Enabled         =   0   'False
         Height          =   195
         Index           =   22
         Left            =   1710
         TabIndex        =   332
         Top             =   4320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No permitir Facturas Negativas"
         Enabled         =   0   'False
         Height          =   195
         Index           =   21
         Left            =   3390
         TabIndex        =   304
         Top             =   4830
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Normales"
         Enabled         =   0   'False
         Height          =   195
         Index           =   16
         Left            =   5100
         TabIndex        =   300
         Top             =   4290
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Descontar Fras.Varias"
         Height          =   195
         Index           =   14
         Left            =   3390
         TabIndex        =   298
         Top             =   4560
         Width           =   1995
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Con Comisión Campo"
         Enabled         =   0   'False
         Height          =   195
         Index           =   13
         Left            =   4710
         TabIndex        =   279
         Top             =   3990
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   59
         Left            =   1650
         MaxLength       =   11
         TabIndex        =   8
         Tag             =   "Kilos Retirados|N|S|||rcampos|canaforo|###,###,###||"
         Top             =   4620
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Retirada"
         Enabled         =   0   'False
         Height          =   195
         Index           =   12
         Left            =   690
         TabIndex        =   277
         Top             =   4320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdAceptarAntGene 
         Caption         =   "&AcepGene"
         Height          =   375
         Left            =   1050
         TabIndex        =   10
         Top             =   6030
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Complementaria"
         Height          =   195
         Index           =   5
         Left            =   3390
         TabIndex        =   276
         Top             =   3690
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Terceros"
         Height          =   195
         Index           =   11
         Left            =   5100
         TabIndex        =   275
         Top             =   3690
         Width           =   1035
      End
      Begin VB.CommandButton cmdAceptarLiqIndustria 
         Caption         =   "&AcepIndus"
         Height          =   375
         Left            =   3030
         TabIndex        =   12
         Top             =   6030
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame FrameAgrupado 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   3120
         TabIndex        =   174
         Top             =   3180
         Width           =   2835
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   175
            Tag             =   "Recolectado|N|N|0|1|rhisfruta|recolect||N|"
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label11 
            Caption         =   "Agrupado por"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   0
            Left            =   270
            TabIndex        =   176
            Top             =   90
            Width           =   1035
         End
      End
      Begin VB.Frame FrameRecolectado 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   3240
         TabIndex        =   169
         Top             =   2730
         Width           =   2865
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   170
            Tag             =   "Recolectado|N|N|0|1|rhisfruta|recolect||N|"
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label Label11 
            Caption         =   "Recolectado"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   171
            Top             =   150
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdAceptarAntGastos 
         Caption         =   "&AcepGast"
         Height          =   375
         Left            =   2010
         TabIndex        =   11
         Top             =   6030
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame FrameOpciones 
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   3120
         TabIndex        =   106
         Top             =   3870
         Width           =   2115
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   2
            Left            =   270
            TabIndex        =   108
            Top             =   90
            Width           =   1965
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Factura"
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   107
            Top             =   390
            Width           =   1995
         End
      End
      Begin VB.Frame FrameFechaAnt 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   525
         Left            =   390
         TabIndex        =   31
         Top             =   3630
         Width           =   3045
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   960
            Picture         =   "frmListAnticipos.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Anticipo"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   25
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListAnticipos.frx":0097
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListAnticipos.frx":03A1
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   1455
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   1095
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1455
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1095
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptarAnt 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   13
         Top             =   6030
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelAnt 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5190
         TabIndex        =   14
         Top             =   6015
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3285
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2880
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   360
         TabIndex        =   54
         Top             =   5700
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   49
         Left            =   420
         TabIndex        =   333
         Top             =   5160
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Image imgAyuda 
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   6150
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   4290
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Kilos Retirados"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   43
         Left            =   420
         TabIndex        =   278
         Top             =   4620
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   1
         Left            =   360
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   4260
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   6150
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   3690
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   12
         Left            =   390
         TabIndex        =   173
         Top             =   6240
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   10
         Left            =   390
         TabIndex        =   172
         Top             =   6030
         Width           =   3525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1350
         MouseIcon       =   "frmListAnticipos.frx":06AB
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1350
         MouseIcon       =   "frmListAnticipos.frx":07FD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   765
         TabIndex        =   30
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   765
         TabIndex        =   29
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   390
         TabIndex        =   28
         Top             =   1830
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1350
         Picture         =   "frmListAnticipos.frx":094F
         ToolTipText     =   "Buscar fecha"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1350
         Picture         =   "frmListAnticipos.frx":09DA
         ToolTipText     =   "Buscar fecha"
         Top             =   3285
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1350
         MouseIcon       =   "frmListAnticipos.frx":0A65
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1350
         MouseIcon       =   "frmListAnticipos.frx":0BB7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   390
         TabIndex        =   25
         Top             =   900
         Width           =   405
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
         Left            =   420
         TabIndex        =   24
         Top             =   345
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   720
         TabIndex        =   23
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   720
         TabIndex        =   22
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   735
         TabIndex        =   21
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   735
         TabIndex        =   20
         Top             =   2940
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   405
         TabIndex        =   19
         Top             =   2700
         Width           =   450
      End
   End
   Begin VB.Frame FrameGenFactAnticipoVC 
      Height          =   6270
      Left            =   0
      TabIndex        =   196
      Top             =   -30
      Width           =   6675
      Begin VB.CheckBox Check1 
         Caption         =   "Tercero"
         Height          =   195
         Index           =   17
         Left            =   420
         TabIndex        =   301
         Top             =   5670
         Width           =   1545
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3075
         MaxLength       =   30
         TabIndex        =   221
         Top             =   2790
         Width           =   885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1950
         MaxLength       =   30
         TabIndex        =   217
         Top             =   3540
         Width           =   4380
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3075
         MaxLength       =   30
         TabIndex        =   216
         Top             =   3915
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   215
         Top             =   3915
         Width           =   1035
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3075
         MaxLength       =   30
         TabIndex        =   214
         Top             =   3180
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1950
         MaxLength       =   4
         TabIndex        =   213
         Top             =   3180
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   45
         Left            =   1980
         MaxLength       =   15
         TabIndex        =   203
         Top             =   4650
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   51
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   200
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1890
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancelAntVC 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5100
         TabIndex        =   205
         Top             =   5625
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepAntVC 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4020
         TabIndex        =   204
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   50
         Left            =   1950
         MaxLength       =   8
         TabIndex        =   202
         Top             =   2790
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   201
         Top             =   2340
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   49
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   199
         Text            =   "Text5"
         Top             =   2340
         Width           =   3255
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   390
         TabIndex        =   197
         Top             =   5070
         Width           =   1965
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Factura"
            Height          =   195
            Index           =   8
            Left            =   30
            TabIndex        =   198
            Top             =   210
            Width           =   1545
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Zona"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   26
         Left            =   810
         TabIndex        =   220
         Top             =   3930
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "Poblacion"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   810
         TabIndex        =   219
         Top             =   3570
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Partida"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   810
         TabIndex        =   218
         Top             =   3210
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe Factura"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   212
         Top             =   4620
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "venta campo,sin entrada en campo asociada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   14
         Left            =   390
         TabIndex        =   211
         Top             =   1230
         Width           =   4680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Proceso por el que generamos una factura de anticipo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   13
         Left            =   390
         TabIndex        =   210
         Top             =   870
         Width           =   5700
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   49
         Left            =   360
         TabIndex        =   209
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   16
         Left            =   1620
         Picture         =   "frmListAnticipos.frx":0D09
         ToolTipText     =   "Buscar fecha"
         Top             =   1890
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Generación Factura Anticipo Venta Campo"
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
         Left            =   300
         TabIndex        =   208
         Top             =   330
         Width           =   6120
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1620
         MouseIcon       =   "frmListAnticipos.frx":0D94
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar campo"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   26
         Left            =   330
         TabIndex        =   207
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   49
         Left            =   1620
         MouseIcon       =   "frmListAnticipos.frx":0EE6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2340
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   330
         TabIndex        =   206
         Top             =   2310
         Width           =   405
      End
   End
   Begin VB.Frame FrameAnticiposPdtes 
      Height          =   5430
      Left            =   0
      TabIndex        =   305
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   3210
         TabIndex        =   321
         Top             =   3930
         Width           =   2865
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   645
         Left            =   690
         TabIndex        =   320
         Top             =   5460
         Width           =   3045
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   67
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   319
         Text            =   "Text5"
         Top             =   2670
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   66
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   318
         Text            =   "Text5"
         Top             =   2310
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   67
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   311
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   310
         Top             =   2310
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListAnticipos.frx":1038
         Style           =   1  'Graphical
         TabIndex        =   317
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command7 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListAnticipos.frx":1342
         Style           =   1  'Graphical
         TabIndex        =   316
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   65
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   314
         Text            =   "Text5"
         Top             =   1695
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   64
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   312
         Text            =   "Text5"
         Top             =   1365
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   309
         Top             =   1710
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   64
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   308
         Top             =   1350
         Width           =   750
      End
      Begin VB.CommandButton CmdAcepAnticiposPdtes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   307
         Top             =   4590
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelAntPdtes 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5160
         TabIndex        =   306
         Top             =   4575
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   63
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   315
         Top             =   3690
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   62
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   313
         Top             =   3330
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   66
         Left            =   1290
         MouseIcon       =   "frmListAnticipos.frx":164C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2310
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   67
         Left            =   1290
         MouseIcon       =   "frmListAnticipos.frx":179E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2670
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   59
         Left            =   705
         TabIndex        =   331
         Top             =   2745
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   58
         Left            =   705
         TabIndex        =   330
         Top             =   2355
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
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
         Left            =   330
         TabIndex        =   329
         Top             =   2100
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   19
         Left            =   1290
         Picture         =   "frmListAnticipos.frx":18F0
         ToolTipText     =   "Buscar fecha"
         Top             =   3690
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   18
         Left            =   1290
         Picture         =   "frmListAnticipos.frx":197B
         ToolTipText     =   "Buscar fecha"
         Top             =   3315
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   65
         Left            =   1290
         MouseIcon       =   "frmListAnticipos.frx":1A06
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   64
         Left            =   1290
         MouseIcon       =   "frmListAnticipos.frx":1B58
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1350
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
         Index           =   56
         Left            =   330
         TabIndex        =   328
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Anticipos Pendientes Descontar"
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
         Left            =   420
         TabIndex        =   327
         Top             =   345
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   55
         Left            =   660
         TabIndex        =   326
         Top             =   1740
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   54
         Left            =   660
         TabIndex        =   325
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   53
         Left            =   675
         TabIndex        =   324
         Top             =   3705
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   52
         Left            =   675
         TabIndex        =   323
         Top             =   3360
         Width           =   465
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
         Index           =   51
         Left            =   345
         TabIndex        =   322
         Top             =   3120
         Width           =   435
      End
   End
   Begin VB.Frame FrameDesFacturacion 
      Height          =   4740
      Left            =   -270
      TabIndex        =   55
      Top             =   60
      Width           =   6555
      Begin VB.Frame FrameTipoFactura 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   390
         TabIndex        =   99
         Top             =   1410
         Width           =   3615
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            ItemData        =   "frmListAnticipos.frx":1CAA
            Left            =   1380
            List            =   "frmListAnticipos.frx":1CAC
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Tag             =   "Recolección|N|N|0|3|rhisfruta|recolect|||"
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Factura"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   100
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   2475
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   66
         Tag             =   "admon"
         Top             =   1170
         Width           =   1545
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   58
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2685
         Width           =   945
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   57
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   59
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3360
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelDesF 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4860
         TabIndex        =   61
         Top             =   4125
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepDesF 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   60
         Top             =   4125
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   255
         Left            =   420
         TabIndex        =   70
         Top             =   3780
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Este proceso borra facturas correlativas "
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   69
         Top             =   450
         Width           =   5820
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Actualiza contadores"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   68
         Top             =   780
         Width           =   5595
      End
      Begin VB.Label Label6 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Index           =   1
         Left            =   1440
         TabIndex        =   67
         Top             =   1170
         Width           =   2235
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   17
         Left            =   900
         TabIndex        =   65
         Top             =   2685
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   10
         Left            =   900
         TabIndex        =   64
         Top             =   2325
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
         Index           =   9
         Left            =   495
         TabIndex        =   63
         Top             =   2055
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   8
         Left            =   465
         TabIndex        =   62
         Top             =   3045
         Width           =   1815
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1470
         Picture         =   "frmListAnticipos.frx":1CAE
         ToolTipText     =   "Buscar fecha"
         Top             =   3360
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
   Begin VB.Frame FrameRecalculoImporte 
      Height          =   3750
      Left            =   0
      TabIndex        =   223
      Top             =   -30
      Width           =   6675
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   48
         Left            =   2295
         MaxLength       =   30
         TabIndex        =   233
         Top             =   1860
         Width           =   3855
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   48
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   231
         Top             =   1860
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   46
         Left            =   1170
         MaxLength       =   12
         TabIndex        =   232
         Top             =   2640
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   52
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   230
         Top             =   1170
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   52
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   229
         Text            =   "Text5"
         Top             =   1170
         Width           =   3855
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   390
         TabIndex        =   224
         Top             =   5070
         Width           =   1965
      End
      Begin VB.CommandButton CmdAcepRecalImp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4020
         TabIndex        =   234
         Top             =   3030
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5100
         TabIndex        =   235
         Top             =   3015
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   31
         Left            =   300
         TabIndex        =   228
         Top             =   960
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   52
         Left            =   870
         MouseIcon       =   "frmListAnticipos.frx":1D39
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   30
         Left            =   330
         TabIndex        =   227
         Top             =   1650
         Width           =   630
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   48
         Left            =   870
         MouseIcon       =   "frmListAnticipos.frx":1E8B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1860
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Recálculo de Importe según kilos"
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
         Left            =   300
         TabIndex        =   226
         Top             =   330
         Width           =   6120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   16
         Left            =   330
         TabIndex        =   225
         Top             =   2400
         Width           =   525
      End
   End
   Begin VB.Frame FrameResultados 
      Height          =   7320
      Left            =   30
      TabIndex        =   109
      Top             =   -60
      Width           =   6675
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   41
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   187
         Text            =   "Text5"
         Top             =   2055
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   42
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   186
         Text            =   "Text5"
         Top             =   2430
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   114
         Top             =   2055
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   115
         Top             =   2430
         Width           =   830
      End
      Begin VB.Frame FrameFechaCertif 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   825
         Left            =   3150
         TabIndex        =   183
         Top             =   5535
         Width           =   3165
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   1710
            MaxLength       =   13
            TabIndex        =   124
            Tag             =   "Campol|N|S|||clientes|codposta|0000000000000||"
            Text            =   "1234567890123"
            Top             =   495
            Width           =   1335
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   1845
            MaxLength       =   10
            TabIndex        =   123
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   135
            Width           =   1230
         End
         Begin VB.Label Label4 
            Caption         =   "Nro.Justificante"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   41
            Left            =   45
            TabIndex        =   185
            Top             =   495
            Width           =   1395
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   13
            Left            =   1440
            Picture         =   "frmListAnticipos.frx":1FDD
            ToolTipText     =   "Buscar fecha"
            Top             =   135
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Certificado"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   40
            Left            =   45
            TabIndex        =   184
            Top             =   135
            Width           =   1410
         End
      End
      Begin VB.Frame FrameOpc 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1290
         Left            =   840
         TabIndex        =   182
         Top             =   5640
         Width           =   2355
         Begin VB.CheckBox Check1 
            Caption         =   "Agrupado por Epígrafe"
            Height          =   255
            Index           =   20
            Left            =   0
            TabIndex        =   302
            Top             =   1080
            Width           =   1965
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Aportación Fondo Operativo"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   236
            Top             =   750
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Certificado Retenciones"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   121
            Top             =   90
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Salta página por Socio"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   122
            Top             =   405
            Width           =   1965
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   3
            Left            =   1980
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   1050
            Width           =   240
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen"
         Height          =   255
         Index           =   4
         Left            =   855
         TabIndex        =   120
         Top             =   5385
         Width           =   1965
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   140
         Text            =   "Text5"
         Top             =   3420
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   28
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   139
         Text            =   "Text5"
         Top             =   3030
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1725
         MaxLength       =   7
         TabIndex        =   117
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   3405
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1725
         MaxLength       =   7
         TabIndex        =   116
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   3015
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   119
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4665
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   118
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4305
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancelResul 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5295
         TabIndex        =   127
         Top             =   6705
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepResul 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4125
         TabIndex        =   126
         Top             =   6705
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   113
         Top             =   1485
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   112
         Top             =   1110
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   111
         Text            =   "Text5"
         Top             =   1485
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   110
         Text            =   "Text5"
         Top             =   1110
         Width           =   3675
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1110
         Index           =   1
         Left            =   3150
         TabIndex        =   125
         Top             =   4260
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
         NumItems        =   0
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1425
         MouseIcon       =   "frmListAnticipos.frx":2068
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar transportista"
         Top             =   2055
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1425
         MouseIcon       =   "frmListAnticipos.frx":21BA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar transportista"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trasportista"
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
         Index           =   44
         Left            =   420
         TabIndex        =   190
         Top             =   1860
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   43
         Left            =   780
         TabIndex        =   189
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   42
         Left            =   780
         TabIndex        =   188
         Top             =   2100
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   29
         Left            =   1440
         MouseIcon       =   "frmListAnticipos.frx":230C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3420
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1440
         MouseIcon       =   "frmListAnticipos.frx":245E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3030
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Informe de Resultados"
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
         Left            =   330
         TabIndex        =   138
         Top             =   300
         Width           =   6150
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   795
         TabIndex        =   137
         Top             =   3405
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   795
         TabIndex        =   136
         Top             =   3045
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
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
         Left            =   435
         TabIndex        =   135
         Top             =   2790
         Width           =   390
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   23
         Left            =   435
         TabIndex        =   134
         Top             =   4005
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   795
         TabIndex        =   133
         Top             =   4305
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   795
         TabIndex        =   132
         Top             =   4680
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmListAnticipos.frx":25B0
         ToolTipText     =   "Buscar fecha"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   9
         Left            =   1440
         Picture         =   "frmListAnticipos.frx":263B
         ToolTipText     =   "Buscar fecha"
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   795
         TabIndex        =   131
         Top             =   1155
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   19
         Left            =   795
         TabIndex        =   130
         Top             =   1530
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
         Index           =   18
         Left            =   435
         TabIndex        =   129
         Top             =   915
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1440
         MouseIcon       =   "frmListAnticipos.frx":26C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1440
         MouseIcon       =   "frmListAnticipos.frx":2818
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   5790
         Picture         =   "frmListAnticipos.frx":296A
         ToolTipText     =   "Marcar todos"
         Top             =   3990
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   6030
         Picture         =   "frmListAnticipos.frx":91BC
         ToolTipText     =   "Desmarcar todos"
         Top             =   3990
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   7
         Left            =   3150
         TabIndex        =   128
         Top             =   4020
         Width           =   1815
      End
   End
   Begin VB.Frame FrameGrabacionModelos 
      Height          =   7245
      Left            =   -240
      TabIndex        =   141
      Top             =   -30
      Width           =   6675
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   145
         Top             =   2580
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   144
         Top             =   2205
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   44
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   192
         Text            =   "Text5"
         Top             =   2580
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   43
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   191
         Text            =   "Text5"
         Top             =   2205
         Width           =   3675
      End
      Begin VB.Frame FrameContacto 
         Caption         =   "Persona de Contacto"
         ForeColor       =   &H00972E0B&
         Height          =   915
         Left            =   390
         TabIndex        =   162
         Top             =   3780
         Width           =   5865
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   37
            Left            =   4470
            MaxLength       =   9
            TabIndex        =   148
            Tag             =   "Campol|N|S|||clientes|codposta|000000000||"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   36
            Left            =   150
            MaxLength       =   40
            TabIndex        =   147
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   510
            Width           =   4260
         End
         Begin VB.Label Label4 
            Caption         =   "Teléfono"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   4530
            TabIndex        =   164
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label4 
            Caption         =   "Apellidos y Nombre"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   210
            TabIndex        =   163
            Top             =   300
            Width           =   2595
         End
      End
      Begin VB.Frame FrameDomicilio 
         Caption         =   "Domicilio Presentador"
         ForeColor       =   &H00972E0B&
         Height          =   855
         Left            =   390
         TabIndex        =   161
         Top             =   5190
         Width           =   5895
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   40
            Left            =   150
            MaxLength       =   2
            TabIndex        =   151
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   480
            Width           =   450
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   39
            Left            =   4710
            MaxLength       =   5
            TabIndex        =   153
            Tag             =   "Campol|N|S|||clientes|codposta|00000||"
            Top             =   480
            Width           =   780
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   38
            Left            =   780
            MaxLength       =   20
            TabIndex        =   152
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   480
            Width           =   3840
         End
         Begin VB.Label Label4 
            Caption         =   "Siglas"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   39
            Left            =   150
            TabIndex        =   167
            Top             =   270
            Width           =   435
         End
         Begin VB.Label Label4 
            Caption         =   "Número"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   4740
            TabIndex        =   166
            Top             =   270
            Width           =   705
         End
         Begin VB.Label Label4 
            Caption         =   "Nombre"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   780
            TabIndex        =   165
            Top             =   270
            Width           =   2595
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   1710
         MaxLength       =   13
         TabIndex        =   146
         Tag             =   "Campol|N|S|||clientes|codposta|0000000000000||"
         Top             =   3210
         Width           =   1380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   35
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   155
         Text            =   "Text5"
         Top             =   1620
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   34
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   154
         Text            =   "Text5"
         Top             =   1245
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1695
         MaxLength       =   6
         TabIndex        =   143
         Top             =   1620
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1695
         MaxLength       =   6
         TabIndex        =   142
         Top             =   1245
         Width           =   830
      End
      Begin VB.CommandButton CmdAcepModelo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   149
         Top             =   6540
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelModelo 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5310
         TabIndex        =   150
         Top             =   6540
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   2730
         MaxLength       =   13
         TabIndex        =   181
         Top             =   4020
         Width           =   1380
      End
      Begin ComctlLib.StatusBar BarraEst 
         Height          =   285
         Left            =   0
         TabIndex        =   168
         Top             =   7800
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   503
         Style           =   1
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   1
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Datos....."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   48
         Left            =   480
         TabIndex        =   222
         Top             =   4710
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   780
         TabIndex        =   195
         Top             =   2250
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   780
         TabIndex        =   194
         Top             =   2625
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trasportista"
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
         Index           =   45
         Left            =   420
         TabIndex        =   193
         Top             =   2010
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1425
         MouseIcon       =   "frmListAnticipos.frx":9BBE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar transportista"
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1425
         MouseIcon       =   "frmListAnticipos.frx":9D10
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar transportista"
         Top             =   2205
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Nro.Justific."
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   28
         Left            =   420
         TabIndex        =   160
         Top             =   3240
         Width           =   945
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   35
         Left            =   1410
         MouseIcon       =   "frmListAnticipos.frx":9E62
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   34
         Left            =   1410
         MouseIcon       =   "frmListAnticipos.frx":9FB4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1245
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
         Index           =   35
         Left            =   405
         TabIndex        =   159
         Top             =   1035
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   34
         Left            =   765
         TabIndex        =   158
         Top             =   1650
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   33
         Left            =   765
         TabIndex        =   157
         Top             =   1275
         Width           =   465
      End
      Begin VB.Label Label9 
         Caption         =   "Grabación Modelo"
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
         Left            =   420
         TabIndex        =   156
         Top             =   270
         Width           =   5160
      End
   End
   Begin VB.Frame FrameReimpresion 
      Height          =   5220
      Left            =   0
      TabIndex        =   33
      Top             =   -30
      Width           =   6675
      Begin VB.CheckBox Check3 
         Caption         =   "Impresión con Arrobas"
         Height          =   225
         Left            =   540
         TabIndex        =   303
         Top             =   4350
         Width           =   1965
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Text5"
         Top             =   3780
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Text5"
         Top             =   3405
         Width           =   3675
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   39
         Top             =   3780
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   38
         Top             =   3405
         Width           =   830
      End
      Begin VB.CommandButton cmdAceptarReimp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4170
         TabIndex        =   40
         Top             =   4275
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelReimp 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5340
         TabIndex        =   42
         Top             =   4275
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   36
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   37
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2775
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1740
         MaxLength       =   7
         TabIndex        =   34
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1365
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   35
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1755
         Width           =   830
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1110
         Index           =   0
         Left            =   3180
         TabIndex        =   101
         Top             =   1350
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
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   5
         Left            =   3180
         TabIndex        =   102
         Top             =   1110
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   6060
         Picture         =   "frmListAnticipos.frx":A106
         ToolTipText     =   "Desmarcar todos"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   5820
         Picture         =   "frmListAnticipos.frx":AB08
         ToolTipText     =   "Marcar todos"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1470
         MouseIcon       =   "frmListAnticipos.frx":1135A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1470
         MouseIcon       =   "frmListAnticipos.frx":114AC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3405
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
         Index           =   11
         Left            =   510
         TabIndex        =   53
         Top             =   3165
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   870
         TabIndex        =   52
         Top             =   3780
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   855
         TabIndex        =   51
         Top             =   3405
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1485
         Picture         =   "frmListAnticipos.frx":115FE
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1485
         Picture         =   "frmListAnticipos.frx":11689
         ToolTipText     =   "Buscar fecha"
         Top             =   2415
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   825
         TabIndex        =   50
         Top             =   2775
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   825
         TabIndex        =   49
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   465
         TabIndex        =   48
         Top             =   2115
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
         Left            =   495
         TabIndex        =   47
         Top             =   1125
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   46
         Top             =   1395
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   45
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   495
         TabIndex        =   44
         Top             =   315
         Width           =   5160
      End
   End
   Begin VB.Frame FrameGeneraFactura 
      Height          =   5790
      Left            =   0
      TabIndex        =   71
      Top             =   30
      Width           =   6585
      Begin VB.CheckBox Check1 
         Caption         =   "Descontar Fras.Varias"
         Height          =   255
         Index           =   15
         Left            =   4200
         TabIndex        =   299
         Top             =   1410
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Descontar AFO"
         Height          =   255
         Index           =   10
         Left            =   4200
         TabIndex        =   273
         Top             =   1020
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   4110
         TabIndex        =   103
         Top             =   3900
         Width           =   1965
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Factura"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   105
            Top             =   360
            Width           =   1545
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   104
            Top             =   0
            Width           =   1635
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   83
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   82
         Top             =   3810
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   81
         Top             =   3270
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   79
         Top             =   2385
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   "Text5"
         Top             =   2025
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Text5"
         Top             =   2385
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   78
         Top             =   2010
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   80
         Top             =   2880
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "Text5"
         Top             =   2880
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "Text5"
         Top             =   3270
         Width           =   3285
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "frmListAnticipos.frx":11714
         Left            =   1920
         List            =   "frmListAnticipos.frx":11716
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Tag             =   "Recolección|N|N|0|3|rhisfruta|recolect|||"
         Top             =   960
         Width           =   1425
      End
      Begin VB.CommandButton CmdAcepGenFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3930
         TabIndex        =   84
         Top             =   5145
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelGenFac 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5100
         TabIndex        =   85
         Top             =   5145
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   77
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1560
         Width           =   1050
      End
      Begin MSComctlLib.ProgressBar Pb3 
         Height          =   255
         Left            =   420
         TabIndex        =   72
         Top             =   4710
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Entrada"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   9
         Left            =   420
         TabIndex        =   98
         Top             =   3630
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   8
         Left            =   975
         TabIndex        =   97
         Top             =   3870
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   975
         TabIndex        =   96
         Top             =   4215
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   930
         TabIndex        =   95
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   930
         TabIndex        =   94
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   4
         Left            =   420
         TabIndex        =   93
         Top             =   1830
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1590
         MouseIcon       =   "frmListAnticipos.frx":11718
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1590
         MouseIcon       =   "frmListAnticipos.frx":1186A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2010
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1590
         Picture         =   "frmListAnticipos.frx":119BC
         ToolTipText     =   "Buscar fecha"
         Top             =   4215
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1590
         Picture         =   "frmListAnticipos.frx":11A47
         ToolTipText     =   "Buscar fecha"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   92
         Top             =   2700
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   975
         TabIndex        =   91
         Top             =   2955
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   975
         TabIndex        =   90
         Top             =   3330
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1590
         MouseIcon       =   "frmListAnticipos.frx":11AD2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1590
         MouseIcon       =   "frmListAnticipos.frx":11C24
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2910
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   3
         Left            =   420
         TabIndex        =   75
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Generación de Factura Venta Campo"
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
         Left            =   390
         TabIndex        =   74
         Top             =   360
         Width           =   5940
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   1590
         Picture         =   "frmListAnticipos.frx":11D76
         ToolTipText     =   "Buscar fecha"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   6
         Left            =   420
         TabIndex        =   73
         Top             =   1380
         Width           =   1815
      End
   End
   Begin VB.Frame FrameAportaciones 
      Height          =   6930
      Left            =   0
      TabIndex        =   237
      Top             =   30
      Width           =   6615
      Begin VB.CheckBox Check2 
         Caption         =   "Resumen"
         Height          =   255
         Left            =   3420
         TabIndex        =   274
         Top             =   5160
         Width           =   2385
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   253
         Top             =   4380
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   57
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   252
         Top             =   4035
         Width           =   1095
      End
      Begin VB.CommandButton CmdCanApor 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5130
         TabIndex        =   261
         Top             =   6225
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepApor 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   259
         Top             =   6240
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   56
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   249
         Top             =   2415
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   55
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   248
         Top             =   2055
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   55
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   247
         Text            =   "Text5"
         Top             =   2055
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   56
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   246
         Text            =   "Text5"
         Top             =   2415
         Width           =   3375
      End
      Begin VB.CommandButton Command6 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListAnticipos.frx":11E01
         Style           =   1  'Graphical
         TabIndex        =   245
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command5 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListAnticipos.frx":1210B
         Style           =   1  'Graphical
         TabIndex        =   244
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   54
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   251
         Top             =   3390
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   53
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   250
         Top             =   3030
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   53
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   243
         Text            =   "Text5"
         Top             =   3030
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   54
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   242
         Text            =   "Text5"
         Top             =   3390
         Width           =   3375
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   645
         Left            =   360
         TabIndex        =   240
         Top             =   4830
         Width           =   3045
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   47
            Left            =   1230
            MaxLength       =   14
            TabIndex        =   255
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe Aportación"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   17
            Left            =   0
            TabIndex        =   241
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   3210
         TabIndex        =   238
         Top             =   3930
         Width           =   2865
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   5
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   257
            Tag             =   "Recolectado|N|N|0|1|rhisfruta|recolect||N|"
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label Label11 
            Caption         =   "Recolectado"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   239
            Top             =   150
            Width           =   1035
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   300
         TabIndex        =   254
         Top             =   5910
         Visible         =   0   'False
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "previos a la liquidación. Hay que seleccionar las mismas variedades"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Index           =   42
         Left            =   90
         TabIndex        =   272
         Top             =   1260
         Width           =   6405
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Este proceso borra los cálculos anteriores de aportaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   90
         TabIndex        =   271
         Top             =   960
         Width           =   6420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   41
         Left            =   345
         TabIndex        =   270
         Top             =   3840
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   40
         Left            =   675
         TabIndex        =   269
         Top             =   4080
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   39
         Left            =   675
         TabIndex        =   268
         Top             =   4425
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   38
         Left            =   660
         TabIndex        =   267
         Top             =   2100
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   37
         Left            =   660
         TabIndex        =   266
         Top             =   2460
         Width           =   420
      End
      Begin VB.Label Label12 
         Caption         =   "Cálculo de Aportaciones Liquidación"
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
         Left            =   420
         TabIndex        =   265
         Top             =   345
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   36
         Left            =   330
         TabIndex        =   264
         Top             =   1860
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   56
         Left            =   1290
         MouseIcon       =   "frmListAnticipos.frx":12415
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2415
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   55
         Left            =   1290
         MouseIcon       =   "frmListAnticipos.frx":12567
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2085
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   17
         Left            =   1290
         Picture         =   "frmListAnticipos.frx":126B9
         ToolTipText     =   "Buscar fecha"
         Top             =   4035
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   15
         Left            =   1290
         Picture         =   "frmListAnticipos.frx":12744
         ToolTipText     =   "Buscar fecha"
         Top             =   4380
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   35
         Left            =   330
         TabIndex        =   263
         Top             =   2820
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   705
         TabIndex        =   262
         Top             =   3075
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   705
         TabIndex        =   260
         Top             =   3465
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   54
         Left            =   1290
         MouseIcon       =   "frmListAnticipos.frx":127CF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3390
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   53
         Left            =   1290
         MouseIcon       =   "frmListAnticipos.frx":12921
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3060
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Aportaciones"
         Height          =   195
         Index           =   32
         Left            =   330
         TabIndex        =   258
         Top             =   6240
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   29
         Left            =   330
         TabIndex        =   256
         Top             =   6450
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin VB.Frame FrameLiqDirecta 
      Height          =   4200
      Left            =   0
      TabIndex        =   280
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton CmdCanLiqDirecta 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5100
         TabIndex        =   289
         Top             =   3540
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepLiqDirecta 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   288
         Top             =   3540
         Width           =   975
      End
      Begin VB.CommandButton Command9 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListAnticipos.frx":12A73
         Style           =   1  'Graphical
         TabIndex        =   291
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command8 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListAnticipos.frx":12D7D
         Style           =   1  'Graphical
         TabIndex        =   290
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   675
         Left            =   180
         TabIndex        =   284
         Top             =   1650
         Width           =   3795
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   61
            Left            =   1830
            MaxLength       =   10
            TabIndex        =   285
            Top             =   330
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Factura"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   44
            Left            =   210
            TabIndex        =   286
            Top             =   330
            Width           =   1035
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   14
            Left            =   1530
            Picture         =   "frmListAnticipos.frx":13087
            ToolTipText     =   "Buscar fecha"
            Top             =   330
            Width           =   240
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   4080
         TabIndex        =   281
         Top             =   1950
         Width           =   2115
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Factura"
            Height          =   195
            Index           =   19
            Left            =   270
            TabIndex        =   283
            Top             =   390
            Width           =   1995
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   18
            Left            =   270
            TabIndex        =   282
            Top             =   60
            Width           =   1965
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   2010
         MaxLength       =   11
         TabIndex        =   287
         Tag             =   "Kilos Retirados|N|S|||rcampos|canaforo|###,###,###||"
         Top             =   2490
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar Pb4 
         Height          =   255
         Left            =   360
         TabIndex        =   292
         Top             =   3120
         Visible         =   0   'False
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "La factura de liquidación generada, se calcula aplicando el Precio sobre cada una de las calidades de la entrada."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   705
         Index           =   48
         Left            =   390
         TabIndex        =   297
         Top             =   930
         Width           =   5865
      End
      Begin VB.Label Label15 
         Caption         =   "Liquidación Directa"
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
         TabIndex        =   296
         Top             =   345
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   47
         Left            =   390
         TabIndex        =   295
         Top             =   3420
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.Label Label2 
         Caption         =   "Accion"
         Height          =   195
         Index           =   46
         Left            =   390
         TabIndex        =   294
         Top             =   3630
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio Calidad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   45
         Left            =   390
         TabIndex        =   293
         Top             =   2520
         Width           =   1020
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Factura"
      ForeColor       =   &H00972E0B&
      Height          =   255
      Index           =   30
      Left            =   0
      TabIndex        =   180
      Top             =   -30
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Desde"
      Height          =   195
      Index           =   31
      Left            =   375
      TabIndex        =   179
      Top             =   300
      Width           =   465
   End
   Begin VB.Label Label4 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   32
      Left            =   375
      TabIndex        =   178
      Top             =   675
      Width           =   420
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   11
      Left            =   1020
      Picture         =   "frmListAnticipos.frx":13112
      ToolTipText     =   "Buscar fecha"
      Top             =   255
      Width           =   240
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   12
      Left            =   1020
      Picture         =   "frmListAnticipos.frx":1319D
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
      TabIndex        =   177
      Top             =   1065
      Width           =   705
   End
End
Attribute VB_Name = "frmListAnticipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados / Procesos ANTICIPOS ====
    '=============================
    ' 1 .- Informe de Anticipos
    ' 2 .- Prevision de Pagos de Anticipos
    ' 3 .- Facturación de Anticipos
    ' 5 .- Deshacer proceso de Facturación Anticipos
    
    
    '==== Listados / Procesos FACTURAS SOCIOS ====
    '==================================
    ' 4 .- Reimpresion de Facturas
    ' 8 .- Informe de Resultados
    ' 9 .- Informe de Retenciones
    
    ' 10.- Grabacion Modelo 190
    ' 11.- Grabación Modelo 346
    
    ' 20.- Anticipos Pendientes de descontar en liquidacion
    
    
    '==== Listados / Procesos VENTA CAMPO ====
    '=============================
    ' 6 .- Facturación de Venta Campo (Anticipo o Liquidación)
    ' 7 .- Deshacer proceso de Facturación de Venta Campo (Anticipo o Liquidación)
    
    ' 16.- Generacion de Factura de anticipo de Venta Campo sin entradas
    ' 161.- Generacion de Factura de anticipo sin entradas
    ' 17.- Proceso de recalculo de importes vc segun kilos
    
    '==== Listados / Procesos LIQUIDACIONES ====
    '================================
    ' 12 .- Informe de Liquidaciones
    ' 13 .- Prevision de Pagos de Liquidacion
    ' 14 .- Facturación de Liquidacion
    ' 15 .- Deshacer proceso de Facturación Anticipos
    
    
    '==== Calculo Aportaciones previo a Liquidacion (SOLO PICASSENT) ====
    '================================
    ' 18 .- Informe de calculo de aportaciones
    
    
    '==== Liquidacion de entrada de hco (POR PARAMETRO) ====
    '=======================================================
    ' 19 .- Liquidacion de entrada del hco
    
    
Public AnticipoGastos As Boolean ' si true entonces es que se trata de anticipos de gastos de recoleccion
Public LiquidacionIndustria As Boolean ' si true entonces es que se trata de liquidacion de industria
Public AnticipoGenerico As Boolean ' si true entonces es que se trata de anticipos genericos,
    ' todos los kilos independientemente de que esten o no clasificados se anticipan a un mismo precio
    
    
    

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(cliente As String, observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

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
Private WithEvents frmCla As frmComercial 'Ayuda de Clases de comercial
Attribute frmCla.VB_VarHelpID = -1
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
Private WithEvents frmMens5 As frmMensajes 'Mensajes
Attribute frmMens5.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean

Private cadSelect2 As String
Private cadSelect3 As String
Private cadSelect1 As String

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim indice As Integer

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte

Dim Bodega As Boolean
Dim Industria As Boolean

Dim Variedades As String
Dim Albaranes As String

Dim vReturn As Integer

Dim vFechas As String



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check1_Click(Index As Integer)
    Select Case Index
        Case 5
'[Monica]27/01/2016: complementaria catadau
            '++Monica:03/06/2013: distinguimos para Catadau entre entradas
            Check1(16).Enabled = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            Check1(16).visible = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            imgAyuda(2).visible = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            imgAyuda(2).Enabled = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            If Check1(16).Enabled Then
                Check1(16).Top = 3690
                imgAyuda(2).Top = 3690
            End If
            
            FrameTipo.Enabled = (vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            FrameTipo.visible = (vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            FrameTipo.Top = 4530
        
            Check1(25).visible = (Check1(5).Value = 1)
            Check1(25).Enabled = (Check1(5).Value = 1)
            Check1(26).visible = (Check1(5).Value = 1)
            Check1(26).Enabled = (Check1(5).Value = 1)
            If Check1(25).Enabled Then
                Check1(25).Value = 1
                Check1(26).Value = 1
            Else
                Check1(25).Value = 0
                Check1(26).Value = 0
            End If
        
        Case 7
            CertificadoRetencionesVisible
        Case 9
            AportacionesFondoOperativoVisible
        Case 12
            KilosRetiradaVisible
        Case 20
            EpigrafeVisible
    End Select
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 7 Then
        CertificadoRetencionesVisible
    End If
End Sub

Private Sub CmdAcepAnticiposPdtes_Click()
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
Dim b As Boolean
Dim Sql2 As String

Dim cadSelect1 As String
Dim Anticipos As String

    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H SOCIO
    cDesde = Trim(txtcodigo(64).Text)
    cHasta = Trim(txtcodigo(65).Text)
    nDesde = txtNombre(64).Text
    nHasta = txtNombre(65).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    'D/H CLASE
    cDesde = Trim(txtcodigo(66).Text)
    cHasta = Trim(txtcodigo(67).Text)
    nDesde = txtNombre(66).Text
    nHasta = txtNombre(67).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
    End If
    
    Sql2 = ""
    If txtcodigo(66).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(66).Text, "N")
    If txtcodigo(67).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(67).Text, "N")
    
    Set frmMens = New frmMensajes
    
    frmMens.OpcionMensaje = 16
    frmMens.cadWhere = Sql2
    frmMens.Show vbModal
    
    Set frmMens = Nothing
    
    
    'D/H Fecha
    cDesde = Trim(txtcodigo(62).Text)
    cHasta = Trim(txtcodigo(63).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc_variedad.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
        
        
    Tabla = "(rfactsoc INNER JOIN rfactsoc_variedad On rfactsoc.codtipom = rfactsoc_variedad.codtipom and rfactsoc.numfactu = rfactsoc_variedad.numfactu and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu )"
    Tabla = Tabla & " INNER JOIN variedades On rfactsoc_variedad.codvarie = variedades.codvarie "
    
    If Not AnyadirAFormula(cadFormula, "{rfactsoc_variedad.descontado} = 0") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{rfactsoc_variedad.descontado} = 0") Then Exit Sub
    
    Anticipos = CodTipomAnticipos
    
    'En el caso de Montifrut, los anticipos estan marcados como
    If vParamAplic.Cooperativa = 12 Then
        If Not AnyadirAFormula(cadFormula, "{rfactsoc.esanticipogasto} = 1") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{rfactsoc.esanticipogasto} = 1") Then Exit Sub
    Else
        If Not AnyadirAFormula(cadFormula, "{rfactsoc.codtipom} in [" & Anticipos & "]") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{rfactsoc.codtipom} in (" & Anticipos & ")") Then Exit Sub
    End If
    
    If HayRegistros(Tabla, cadSelect) Then
    
        indRPT = 103 ' "rAntPdtesDescontar.rpt"
    
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    
        cadNombreRPT = nomDocu
        cadTitulo = "Informe de Anticipos Pdtes Descontar"
        
        LlamarImprimir
    End If
        

End Sub

Private Sub CmdAcepAntVC_Click()
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
Dim tipoMov As String

Dim vSQL As String

    vSQL = ""
    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        nTabla = "rsocios INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio "
        
        
        Select Case OpcionListado
            Case 16 ' anticipo venta campo sin entrada
        
                If Not ComprobarTiposMovimiento(2, nTabla, cadSelect) Then Exit Sub
                        
                If FacturaAnticipoVentaCampo(txtcodigo(49).Text, txtcodigo(50).Text, txtcodigo(45).Text, txtcodigo(51).Text) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                   
                    
                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE VENTA CAMPO
                    If Me.Check1(8).Value Then
                        cadFormula = ""
                        cadSelect = ""
                        tipoMov = "FAC"
                        cadAux = "({stipom.tipodocu} = 3)"
                        cadTitulo = "Reimpresión Facturas Anticipos V.Campo"
                        
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                        'Nº Factura
                        cadAux = "({rfactsoc.numfactu} IN [" & vParamAplic.UltFactAntVC & "])"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                        'Fecha de Factura
                        FecFac = CDate(txtcodigo(51).Text)
                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = "{rfactsoc.fecfactu}= '" & Format(FecFac, FormatoFecha) & "'"
                        
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                       
                        indRPT = 23 'Impresion de facturas de socios
                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        ConSubInforme = True
                        
                        LlamarImprimir
                        
                        If frmVisReport.EstaImpreso Then
                            ActualizarRegistrosFac "rfactsoc", cadSelect
                        End If
                    End If
                                   
                End If
                
                
            Case 161 ' anticipo normal sin entrada
        
                If Not ComprobarTiposMovimiento(0, nTabla, cadSelect) Then Exit Sub
                        
                If FacturaAnticipoSinEntrada(txtcodigo(49).Text, txtcodigo(50).Text, txtcodigo(45).Text, txtcodigo(51).Text, Check1(17).Value = 1) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                   
                    
                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE VENTA CAMPO
                    If Me.Check1(8).Value Then
                        cadFormula = ""
                        cadSelect = ""
                        
                        '[Monica]07/11/2013: si esta marcado que es un socio tercero cogemos otro contador (Picassent)
                        If Check1(17).Value = 1 Then
                            tipoMov = "FAT"
                        Else
                            tipoMov = "FAA"
                        End If
                        cadAux = "({stipom.tipodocu} = 1)"
                        cadTitulo = "Reimpresión Facturas Anticipos"
                        
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                        'Nº Factura
                        cadAux = "({rfactsoc.numfactu} IN [" & vParamAplic.UltFactAnt & "])"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                        'Fecha de Factura
                        FecFac = CDate(txtcodigo(51).Text)
                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = "{rfactsoc.fecfactu}= '" & Format(FecFac, FormatoFecha) & "'"
                        
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                       
                        indRPT = 23 'Impresion de facturas de socios
                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        ConSubInforme = True
                        
                        LlamarImprimir
                        
                        If frmVisReport.EstaImpreso Then
                            ActualizarRegistrosFac "rfactsoc", cadSelect
                        End If
                    End If
                                   
                End If
        End Select
    
    
    End If
    CmdCancelAntVC_Click

End Sub

Private Sub CmdAcepApor_Click()
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

Dim b As Boolean
Dim Sql2 As String

Dim MaxContador As String


    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(55).Text)
        cHasta = Trim(txtcodigo(56).Text)
        nDesde = txtNombre(55).Text
        nHasta = txtNombre(56).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(53).Text)
        cHasta = Trim(txtcodigo(54).Text)
        nDesde = txtNombre(53).Text
        nHasta = txtNombre(54).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(53).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(53).Text, "N")
        If txtcodigo(54).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(54).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(57).Text)
        cHasta = Trim(txtcodigo(58).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
            
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        'Socio que no sea tercero
        If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        
'        'sólo entradas distintas de VENTA CAMPO y distintas de INDUSTRIA
'        If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3") Then Exit Sub
'        If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3") Then Exit Sub
        
        '++monica: 28/07/2009 dependiendo del tipo de recoleccion (0=coop 1=socio 2=todos)
        Select Case Combo1(2).ListIndex
            Case 0      ' recolectado cooperativa
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 0") Then Exit Sub
            Case 1      ' recolectado socio
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 1") Then Exit Sub
            Case 2      ' ambos
            
        End Select
        
        nTabla = "((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
                      
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWhere = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
            If CargarAportaciones(nTabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                
                cadTitulo = "Informe de Aportación Fondo Operativo"

                indRPT = 75 ' rInformeAFO.rpt

                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub

                cadNombreRPT = Replace(nomDocu, "AFO", "AFOResul")
                
                cadParam = cadParam & "pResumen=" & Check2.Value & "|"
                numParam = numParam + 1
                
                cadFormula = ""
                
                
                LlamarImprimir

                CmdCanApor_Click
            End If
        End If
    End If
    
End Sub

Private Function CargarAportaciones(ctabla As String, cwhere As String)
Dim SQL As String
Dim Sql2 As String
Dim TotalKilos As Long
Dim ImporteSoc As Currency
Dim Importe As Currency
Dim RS As ADODB.Recordset
Dim Precio As Double
Dim TotImpor As Currency
Dim TotalSocios As Long
Dim Reg As Long


    On Error GoTo eCargarAportaciones

    CargarAportaciones = False

    SQL = "delete from raporreparto"
    conn.Execute SQL

    Me.Label2(32).visible = True
    Me.Refresh
    DoEvents

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    SQL = "Select rhisfruta.codsocio, 0,sum(rhisfruta.kilosnet), 0, 0  FROM " & QuitarCaracterACadena(ctabla, "_1")
    SQL = SQL & " where rhisfruta.tipoentr <> 1 "
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " and " & cwhere
    End If
    SQL = SQL & " group by 1, 2 "
    SQL = SQL & " union "
    SQL = SQL & "Select rhisfruta.codsocio, 1,sum(rhisfruta.kilosnet), 0, 0  FROM " & QuitarCaracterACadena(ctabla, "_1")
    SQL = SQL & " where rhisfruta.tipoentr = 1 "
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " and " & cwhere
    End If
    SQL = SQL & " group by 1, 2 "
        

    Sql2 = "insert into raporreparto (codsocio, tipoentr, kilosnet, importe, precio) " & SQL
    conn.Execute Sql2
    
    TotalSocios = DevuelveValor("select count(*) from raporreparto")
    
    
    SQL = "select sum(kilosnet) from raporreparto"
    TotalKilos = DevuelveValor(SQL)
    Importe = CCur(txtcodigo(47).Text)
    
    Precio = Round2(Importe / TotalKilos, 6)
    TotImpor = 0
    Reg = 0
    
    Me.Label2(32).Caption = "Calculando Prorrateo"
    Me.Refresh
    DoEvents
    
    SQL = "select codsocio, tipoentr, kilosnet from raporreparto order by codsocio"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Reg = Reg + 1
        ImporteSoc = Round2(Precio * DBLet(RS!KilosNet, "N"), 2)
        
        If Reg <> TotalSocios Then
            TotImpor = TotImpor + ImporteSoc
        Else
            ImporteSoc = Importe - TotImpor
        End If
        
        SQL = "update raporreparto set importe = " & DBSet(ImporteSoc, "N")
        SQL = SQL & ", precio = " & TransformaComasPuntos(ImporteSinFormato(CStr(Precio)))
        SQL = SQL & " where codsocio= " & DBSet(RS!Codsocio, "N")
        SQL = SQL & " and tipoentr= " & DBSet(RS!TipoEntr, "N")
        conn.Execute SQL
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    
    Me.Label2(32).visible = False
    Me.Refresh
    DoEvents

    CargarAportaciones = True
    Exit Function
    
eCargarAportaciones:
    Me.Label2(32).visible = False
    Me.Refresh

    MuestraError Err.Number, "Cargar Aportaciones", Err.Description
End Function


Private Sub CmdAcepDesF_Click()
Dim Tipo As Byte
    If DatosOk Then
        Pb2.visible = True
        Select Case OpcionListado
            Case 5 ' anticipo
                Tipo = 0
            Case 7
                ' venta campo
                Select Case Combo1(1).ListIndex
                    Case 0 ' anticipo
                        Tipo = 1
                    Case 1 ' liquidacion
                        Tipo = 2
                End Select
            Case 15 ' liquidacion
                Tipo = 3
        End Select
        If DeshacerFacturacion(Tipo, txtcodigo(9).Text, txtcodigo(10).Text, txtcodigo(11).Text, Pb2) Then
            MsgBox "Proceso realizado correctamente", vbExclamation
            cmdCancelDesF_Click
        End If
    End If
End Sub


Private Sub CmdAcepGenFac_Click()
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
Dim tipoMov As String

Dim vSQL As String

    vSQL = ""
    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(16).Text)
        cHasta = Trim(txtcodigo(17).Text)
        nDesde = txtNombre(16).Text
        nHasta = txtNombre(17).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(18).Text)
        cHasta = Trim(txtcodigo(19).Text)
        nDesde = txtNombre(18).Text
        nHasta = txtNombre(19).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        If txtcodigo(18).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(18).Text, "N")
        If txtcodigo(19).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(19).Text, "N")
        
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(22).Text)
        cHasta = Trim(txtcodigo(23).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        'sólo entradas de VENTA CAMPO
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} = 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} = 1") Then Exit Sub
        
        'sólo entradas que tengan importe (rhisfruta.impentrada)
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.impentrada} <> 0") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.impentrada} <> 0") Then Exit Sub
        
        
        nTabla = "(rhisfruta INNER JOIN rsocios_seccion ON rhisfruta.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie "
        nTabla = "(" & nTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
        
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWhere = vSQL
        
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        If HayRegParaInforme(nTabla, cadSelect) Then
            Nregs = TotalFacturas(nTabla, cadSelect)
            If Nregs <> 0 Then
                'combo1(0).listindex = 0 ---> anticipo venta campo
                '                    = 1 ---> liquidación venta campo
                Select Case Combo1(0).ListIndex
                    Case 0 ' anticipo
                        If Not ComprobarTiposMovimiento(2, nTabla, cadSelect) Then Exit Sub
                    Case 1 ' liquidacion venta campo
                        If Not ComprobarTiposMovimiento(3, nTabla, cadSelect) Then Exit Sub
                End Select
                
                Me.Pb3.visible = True
                Me.Pb3.Max = Nregs
                Me.Pb3.Value = 0
                Me.Refresh
                        
                If FacturacionVentaCampo(Combo1(0).ListIndex, nTabla, cadSelect, txtcodigo(14).Text, Me.Pb3, Check1(10).Value, Check1(15).Value = 1) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                   
                    ' si imprimimos resumen
                    If Me.Check1(0).Value Then
                        cadFormula = ""
                        cadParam = cadParam & "pFecFac= """ & txtcodigo(14).Text & """|"
                        numParam = numParam + 1
                        
                        FecFac = CDate(txtcodigo(14).Text)
                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                
                        cadNombreRPT = "rResumFacturas.rpt"
                        
                        Select Case Combo1(0).ListIndex
                            Case 0 ' anticipos
                                cadTitulo = "Resumen Facturas Anticipos Venta Campo"
                                cadParam = cadParam & "pTitulo= ""Resumen Fact.Anticipos V.Campo""|"
                                numParam = numParam + 1
                            Case 1 ' liquidaciones
                                cadTitulo = "Resumen Facturas Liquidación Venta Campo"
                                cadParam = cadParam & "pTitulo= ""Resumen Fact.Liquidación V.Campo""|"
                                numParam = numParam + 1
                        End Select
                        ConSubInforme = True
                        
                        LlamarImprimir
                    End If
                    
                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE VENTA CAMPO
                    If Me.Check1(1).Value Then
                        cadFormula = ""
                        cadSelect = ""
                        'Tipo de Factura: Anticipo
                        Select Case Combo1(0).ListIndex
                            Case 0 ' anticipos
                                tipoMov = "FAC"
                                cadAux = "({stipom.tipodocu} = 3)"
                                cadTitulo = "Reimpresión Facturas Anticipos V.Campo"
                            Case 1
                                tipoMov = "FLC"
                                cadAux = "({stipom.tipodocu} = 4)"
                                cadTitulo = "Reimpresión Facturas Liquidación V.Campo"
                        End Select
                        
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                        'Nº Factura
                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(Combo1(0).ListIndex) & "])"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                        'Fecha de Factura
                        FecFac = CDate(txtcodigo(14).Text)
                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = "{rfactsoc.fecfactu}= '" & Format(FecFac, FormatoFecha) & "'"
                        
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                       
                        indRPT = 23 'Impresion de facturas de socios
                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        ConSubInforme = True
                        
                        If indRPT = 23 And vParamAplic.Cooperativa = 2 Then
                            Dim PrecioApor As Double
                            PrecioApor = DevuelveValor("select min(precio) from raporreparto")
                            
                            cadParam = cadParam & "pPrecioApor=""" & Replace(Format(PrecioApor, "#0.000000"), ",", ".") & """|"
                            numParam = numParam + 1
                        End If
                        
                        LlamarImprimir
                        
                        If frmVisReport.EstaImpreso Then
                            ActualizarRegistrosFac "rfactsoc", cadSelect
                        End If
                    End If
                                   
                End If
            Else
                MsgBox "No hay entradas a facturar.", vbExclamation
            End If
            
            Me.Pb3.visible = False
            CmdCancelGenFac_Click
        End If
    End If
End Sub

Private Sub CmdAcepLiqDirecta_Click()
Dim vTabla As String
Dim vWhere As String
Dim FecFac As Date
Dim cadAux As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim nTabla As String

    '[Monica]11/03/2015: observaciones de factura
    ObsFactura = txtcodigo(68)



    InicializarVbles
    
    If Not DatosOk Then Exit Sub

    'comprobamos que los tipos de iva existen en la contabilidad de horto
    If Not ComprobarTiposIVA("rhisfruta", "rhisfruta.numalbar = " & NumCod) Then Exit Sub
    
    
    If FacturacionLiquidacionDirecta(NumCod, txtcodigo(61).Text, txtcodigo(60).Text) Then
        MsgBox "Proceso realizado correctamente", vbExclamation
        
        'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
        If Me.Check1(18).Value Then
            cadFormula = ""
            cadParam = cadParam & "pFecFac= """ & txtcodigo(61).Text & """|"
            numParam = numParam + 1
            cadParam = cadParam & "pTitulo= ""Resumen Facturación de Liquidaciones""|"
            numParam = numParam + 1
            
            FecFac = CDate(txtcodigo(61).Text)
            cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
            If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
            ConSubInforme = True
            
            LlamarImprimir
        End If
        'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS/LIQUIDACION
        If Me.Check1(3).Value Then
            cadFormula = ""
            cadSelect = ""
            cadAux = "({stipom.tipodocu} = 2)"
            If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
            'Nº Factura
            cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(3) & "])"
            If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
            cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
            If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

            'Fecha de Factura
            FecFac = CDate(txtcodigo(61).Text)
            cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
            If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
            cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
            If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

            indRPT = 23 'Impresion de facturas de socios
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            'Nombre fichero .rpt a Imprimir
            cadNombreRPT = nomDocu
            'Nombre fichero .rpt a Imprimir
            cadTitulo = "Reimpresión de Facturas Liquidaciones"
            ConSubInforme = True

            If indRPT = 23 And vParamAplic.Cooperativa = 2 Then
                Dim PrecioApor As Double
                PrecioApor = DevuelveValor("select min(precio) from raporreparto")
                
                cadParam = cadParam & "pPrecioApor=""" & Replace(Format(PrecioApor, "#0.000000"), ",", ".") & """|"
                numParam = numParam + 1
            End If

            LlamarImprimir

            If frmVisReport.EstaImpreso Then
                ActualizarRegistrosFac "rfactsoc", cadSelect
            End If
        End If
        'SALIR DE LA FACTURACION DE ANTICIPOS / LIQUIDACIONES
        cmdCancelAnt_Click
        
    End If
    
End Sub


Private Sub CmdAcepModelo_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim i As Byte
Dim nTabla As String
Dim nTabla2 As String
Dim nTabla3 As String

Dim vWhere As String
Dim b As Boolean
Dim Tipo As Byte
Dim FecFin As String
Dim FecIni As String
Dim Codigo1 As String
Dim codigo2 As String

Dim vCampAnt As CCampAnt


    InicializarVbles
    
    If Not DatosOk Then Exit Sub

    '++monica:[30/11/2009] montamos las fechas de inicio y fin de año natural de la fecha de inicio
    FecFin = Format(CStr(Format(Year(vParam.FecIniCam), "0000")) & "-" & "12" & "-" & "31")
    If Not EsFechaOK(FecFin) Then
        MsgBox "Fecha inicio de campaña incorrecta. Revise.", vbExclamation
        Exit Sub
    End If
    
    FecIni = CStr(DateAdd("d", 1, DateAdd("yyyy", -1, FecFin)))
    '++

    'D/H Socios
    cDesde = Trim(txtcodigo(34).Text)
    cHasta = Trim(txtcodigo(35).Text)
    nDesde = txtNombre(34).Text
    nHasta = txtNombre(35).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        Codigo1 = "{rcafter.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
        
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect1, Codigo & ">= " & cDesde) Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect1, Codigo & "<= " & cHasta) Then Exit Sub
        
        '[Monica]20/01/2015: añadida la tabla de terceros
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect3, Codigo1 & ">= " & cDesde) Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect3, Codigo1 & "<= " & cHasta) Then Exit Sub
        
    End If
    
    
    'D/H Transportistas
    cDesde = Trim(txtcodigo(43).Text)
    cHasta = Trim(txtcodigo(44).Text)
    nDesde = txtNombre(43).Text
    nHasta = txtNombre(44).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfacttra.codtrans}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTransport= """) Then Exit Sub
    
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect2, Codigo & ">= '" & cDesde & "'") Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect2, Codigo & "<= '" & cHasta & "'") Then Exit Sub
    
    End If
    
'--monica[30/11/2009]: ya no pedimos desde hasta fecha, pq es el año natural de la fecha inicio campaña
'    'D/H Fecha factura
'    cDesde = Trim(txtcodigo(32).Text)
'    cHasta = Trim(txtcodigo(33).Text)
    cDesde = FecIni
    cHasta = FecFin
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        Codigo1 = "{rfacttra.fecfactu}"
        codigo2 = "{rcafter.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
        
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect1, Codigo & ">= '" & Format(cDesde, FormatoFecha) & "'") Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect1, Codigo & "<= '" & Format(cHasta, FormatoFecha) & "'") Then Exit Sub
    
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect2, Codigo1 & ">= '" & Format(cDesde, FormatoFecha) & "'") Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect2, Codigo1 & "<= '" & Format(cHasta, FormatoFecha) & "'") Then Exit Sub
    
        '[Monica]20/01/2015: añadida la tabla de terceros
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect3, codigo2 & ">= '" & Format(cDesde, FormatoFecha) & "'") Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect3, codigo2 & "<= '" & Format(cHasta, FormatoFecha) & "'") Then Exit Sub
    End If
   
    nTabla = vEmpresa.BDAriagro & ".rfactsoc INNER JOIN usuarios.stipom stipom ON rfactsoc.codtipom = stipom.codtipom "
    
    
    txtcodigo(30).Text = Format(Year(FecIni), "0000") ' inicio del año natural
    
    
    Select Case OpcionListado
        Case 10 'modelo 190
            If Not AnyadirAFormula(cadFormula, "{rfactsoc.impreten} <> 0") Then Exit Sub
            If Not AnyadirAFormula(cadSelect1, "{rfactsoc.impreten} <> 0") Then Exit Sub
        
            If Not AnyadirAFormula(cadSelect2, "{rfacttra.impreten} <> 0") Then Exit Sub
            
            '[Monica]20/01/2015: Añadimos la tabla de facturas de terceros
            If Not AnyadirAFormula(cadSelect3, "{rcafter.trefacpr} <> 0") Then Exit Sub
            
        
            If Not AnyadirAFormula(cadFormula, "{stipom.tipodocu} in [1,2,3,4,5,6,7,8,9,10,11]") Then Exit Sub
            If Not AnyadirAFormula(cadSelect1, "{stipom.tipodocu} in (1,2,3,4,5,6,7,8,9,10,11)") Then Exit Sub
        
        Case 11 'modelo 346
            ' seleccionamos tipodocu: 5 = subvencion
            '                         6 = siniestro
            If Not AnyadirAFormula(cadFormula, "{stipom.tipodocu} in [5,6]") Then Exit Sub
            If Not AnyadirAFormula(cadSelect1, "{stipom.tipodocu} in (5,6)") Then Exit Sub
            If Not AnyadirAFormula(cadSelect2, "{stipom.tipodocu} in (5,6)") Then Exit Sub
    
            If Not AnyadirAFormula(cadFormula, "{rfactsoc_variedad.imporvar} <> 0") Then Exit Sub
            If Not AnyadirAFormula(cadSelect2, "{rfacttra_albaran.importe} <> 0") Then Exit Sub
            
            If Not AnyadirAFormula(cadSelect1, "{rfactsoc_variedad.imporvar} <> 0") Then Exit Sub
            
            nTabla = "(" & nTabla & ") INNER JOIN " & vEmpresa.BDAriagro & ".rfactsoc_variedad ON rfactsoc.codtipom = rfactsoc_variedad.codtipom "
            nTabla = nTabla & " and rfactsoc.numfactu = rfactsoc_variedad.numfactu "
            nTabla = nTabla & " and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
            
    End Select
    
    nTabla2 = Replace(Replace(nTabla, "rfactsoc_variedad", "rfacttra_albaran"), "rfactsoc", "rfacttra")
    
    nTabla3 = vEmpresa.BDAriagro & ".rcafter INNER JOIN " & vEmpresa.BDAriagro & ".rlifter ON rcafter.codsocio = rlifter.codsocio and rcafter.numfactu = rlifter.numfactu and rcafter.fecfactu = rlifter.fecfactu "
    
    Label4(48).visible = True
    DoEvents
    
    If CargarFacturas(nTabla, cadSelect1, nTabla2, cadSelect2, nTabla3, cadSelect3) Then
        
        If HayRegParaInforme("tmprfactsoc", "tmprfactsoc.codusu=" & vUsu.Codigo) Then 'nTabla, cadSelect) Then
'            b = GeneraFicheroModelo(OpcionListado - 10, nTabla, cadSelect)
            Label4(48).Caption = "Generando fichero..."
            DoEvents
            b = GeneraFicheroModelo(OpcionListado - 10, "tmprfactsoc", "tmprfactsoc.codusu=" & vUsu.Codigo)
            If b Then
                If CopiarFichero Then
                    MsgBox "Proceso realizado correctamente", vbExclamation
                    CmdCancelModelo_Click
                End If
            End If
        End If
        
   End If
   Label4(48).visible = False
   DoEvents

End Sub

Private Sub CmdAcepRecalImp_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim nTabla As String
Dim nTabla2 As String
Dim vSQL As String
Dim Codigo1 As String
Dim SQL As String

    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub

    ' Socio el socio es obligatorio introducirlo
    If Not AnyadirAFormula(cadSelect, "rhisfruta.codsocio = " & Trim(txtcodigo(52).Text)) Then Exit Sub
    
    ' Variedad
    If Trim(txtcodigo(48).Text) <> "" Then
        If Not AnyadirAFormula(cadSelect, "rhisfruta.codvarie = " & Trim(txtcodigo(48).Text)) Then Exit Sub
    End If
    
    If Not AnyadirAFormula(cadSelect, "rhisfruta.tipoentr = 1") Then Exit Sub
    
    SQL = "select count(*) from rhisfruta where " & cadSelect
    If TotalRegistros(SQL) = 0 Then
        MsgBox "No hay entradas de venta campo de este socio. Revise.", vbExclamation
    Else
        ' cargamos el listview para que se seleccionen que campos hemos de modificar
        Set frmMens2 = New frmMensajes
        frmMens2.cadWhere = cadSelect
        frmMens2.OpcionMensaje = 26 '6
        frmMens2.Show vbModal
        Set frmMens2 = Nothing
        
        If Albaranes = "" Then
            MsgBox "No se han seleccionado albaranes para hacer el reparto. Revise.", vbExclamation
        Else
            If RecalculoImportes(Albaranes) Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                cmdCancel_Click
            End If
        End If
    End If

End Sub

Private Sub CmdAcepResul_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim nTabla As String
Dim nTabla2 As String
Dim vSQL As String
Dim Codigo1 As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    'Tipo de movimiento:
    Tipos = ""
    For i = 1 To ListView1(1).ListItems.Count
        If ListView1(1).ListItems(i).Checked Then
            Tipos = Tipos & DBSet(ListView1(1).ListItems(i).Key, "T") & ","
        End If
    Next i
    
    If Tipos = "" Then
        MsgBox "Debe seleccionar al menos un tipo de factura.", vbExclamation
        Exit Sub
    Else
        ' quitamos la ultima coma
        Tipos = "{rfactsoc.codtipom} in (" & Mid(Tipos, 1, Len(Tipos) - 1) & ")"
        If Not AnyadirAFormula(cadSelect1, Tipos) Then Exit Sub
        Tipos = Replace(Replace(Tipos, "(", "["), ")", "]")
        If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
    End If
    
    'D/H Socios
    cDesde = Trim(txtcodigo(24).Text)
    cHasta = Trim(txtcodigo(25).Text)
    nDesde = txtNombre(24).Text
    nHasta = txtNombre(25).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect1, Codigo & ">= " & cDesde) Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect1, Codigo & "<= " & cHasta) Then Exit Sub
        
    End If
    
    
    'D/H Transportistas
    cDesde = Trim(txtcodigo(41).Text)
    cHasta = Trim(txtcodigo(42).Text)
    nDesde = txtNombre(41).Text
    nHasta = txtNombre(42).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfacttra.codtrans}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTransport= """) Then Exit Sub
    
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect2, Codigo & ">= '" & cDesde & "'") Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect2, Codigo & "<= '" & cHasta & "'") Then Exit Sub
    
    End If
    
    
    
    'D/H Clase
    cDesde = Trim(txtcodigo(28).Text)
    cHasta = Trim(txtcodigo(29).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
        
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect1, Codigo & ">= " & cDesde) Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect1, Codigo & "<= " & cHasta) Then Exit Sub
    
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect2, Codigo & ">= " & cDesde) Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect2, Codigo & "<= " & cHasta) Then Exit Sub
    
    End If
    
    vSQL = ""
    If txtcodigo(28).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(28).Text, "N")
    If txtcodigo(29).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(29).Text, "N")
    
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(26).Text)
    cHasta = Trim(txtcodigo(27).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        Codigo1 = "{rfacttra.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
        
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect1, Codigo & ">= '" & Format(cDesde, FormatoFecha) & "'") Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect1, Codigo & "<= '" & Format(cHasta, FormatoFecha) & "'") Then Exit Sub
    
        If cDesde <> "" Then If Not AnyadirAFormula(cadSelect2, Codigo1 & ">= '" & Format(cDesde, FormatoFecha) & "'") Then Exit Sub
        If cHasta <> "" Then If Not AnyadirAFormula(cadSelect2, Codigo1 & "<= '" & Format(cHasta, FormatoFecha) & "'") Then Exit Sub
    End If
        
    nTabla = "(" & vEmpresa.BDAriagro & ".rfactsoc INNER JOIN " & vEmpresa.BDAriagro & ".rfactsoc_variedad ON rfactsoc.codtipom = rfactsoc_variedad.codtipom "
    nTabla = nTabla & " and rfactsoc.numfactu = rfactsoc_variedad.numfactu and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu) "
    nTabla = nTabla & " INNER JOIN " & vEmpresa.BDAriagro & ".variedades ON rfactsoc_variedad.codvarie = variedades.codvarie "
    
    Set frmMens = New frmMensajes
    
    frmMens.OpcionMensaje = 16
    frmMens.cadWhere = vSQL
    frmMens.Show vbModal
    
    Set frmMens = Nothing
    
    nTabla2 = Replace(Replace(nTabla, "rfactsoc_variedad", "rfacttra_albaran"), "rfactsoc", "rfacttra")
    nTabla2 = Replace(nTabla2, "rfacttra.fecfactu = rfacttra_albaran.fecfactu", "rfacttra.fecfactu = rfacttra_albaran.fecfactu and rfacttra.codtrans = rfacttra_albaran.codtrans")
    
    '[Monica]09/12/2013: Comprobamos si es por epigrafe que no hayan facturas con variedades de distinto grupo
    If Check1(20).Value = 1 Then
        If HayFacturasConLineasDeDistintoGrupo(nTabla, cadSelect1) Then
            MsgBox "Hay Facturas con variedades de distinto grupo. El informe no será correcto.", vbExclamation
        End If
    End If
    
    
    If CargarFacturas(nTabla, cadSelect1, nTabla2, cadSelect2) Then

        If HayRegistros("tmprfactsoc", "tmprfactsoc.codusu=" & vUsu.Codigo) Then
            cadParam = cadParam & "pResumen=" & Me.Check1(4).Value & "|"
            numParam = numParam + 1
            'Nombre fichero .rpt a Imprimir
            Select Case OpcionListado
                Case 8
                    'Nombre fichero .rpt a Imprimir
                    cadNombreRPT = "rInfResultados.rpt"
                    cadTitulo = "Informe de Resultados"
                    
'                    If vParamAplic.Cooperativa = 12 Then
'                        cadNombreRPT = "MonInfResultados.rpt"
'                    End If
                    
                Case 9
                    'Nombre fichero .rpt a Imprimir
                    cadNombreRPT = "rInfRetenciones.rpt"
                    cadTitulo = "Informe de Retenciones"
                    
                    indRPT = 76 ' Informe de Retenciones

                    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub

                    cadNombreRPT = nomDocu
                    
                    If Check1(7).Value = 1 Then
                        'cadNombreRPT = "rInfRetenciones.rpt"
                        cadTitulo = "Certificado de Retenciones"

                        indRPT = 39 ' Certificado de Retenciones

                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub

                        cadNombreRPT = nomDocu

                        cadParam = cadParam & "pDesdeFecha=""" & txtcodigo(26).Text & """|"
                        numParam = numParam + 1
                        cadParam = cadParam & "pHastaFecha=""" & txtcodigo(27).Text & """|"
                        numParam = numParam + 1
                        cadParam = cadParam & "pFecha=""" & txtcodigo(32).Text & """|"
                        numParam = numParam + 1
                        cadParam = cadParam & "pJustificante=" & txtcodigo(33).Text & "|"
                        numParam = numParam + 1
                    Else
                        If Check1(9).Value = 1 Then ' informe de aportaciones
                            cadTitulo = "Informe de Aportación Fondo Operativo"
    
                            indRPT = 75 ' rInformeAFO.rpt
    
                            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    
                            cadNombreRPT = nomDocu
    
                            cadParam = cadParam & "pDesdeFecha=""" & txtcodigo(26).Text & """|"
                            numParam = numParam + 1
                            cadParam = cadParam & "pHastaFecha=""" & txtcodigo(27).Text & """|"
                            numParam = numParam + 1
                            cadParam = cadParam & "pFecha=""" & txtcodigo(32).Text & """|"
                            numParam = numParam + 1
                            
                            '[Monica]12/01/2012 : el precio lo paso a capon
                            Dim Kilos As Long
                            Dim Importe As Currency
                            Dim Precio As Double
                            Dim Sql5 As String
                            
                            Sql5 = "select sum(if(impapor is null,0,impapor)) from tmprfactsoc where codusu = " & vUsu.Codigo
                            Importe = DevuelveValor(Sql5)
                            Sql5 = "select sum(if(kilosnet is null,0,kilosnet)) from tmprfactsoc_variedad where codusu = " & vUsu.Codigo
                            Kilos = DevuelveValor(Sql5)
                            Precio = Round2(Importe / Kilos, 6)
                            cadParam = cadParam & "pPrecio=" & TransformaComasPuntos(ImporteSinFormato(CStr(Precio))) & "|"
                            numParam = numParam + 1
                            
                        Else
                            '[Monica]10/12/2013: informe de retenciones por epígrafe
                            If Check1(20).Value = 1 Then
                                'cadNombreRPT = "rInfRetenciones.rpt"
                                cadNombreRPT = Replace(cadNombreRPT, "Retenciones", "RetencionesEpi")
                                cadTitulo = "Informe de Retenciones por Socio/Epígrafe"
            
                                cadParam = cadParam & "pDesdeFecha=""" & txtcodigo(26).Text & """|"
                                numParam = numParam + 1
                                cadParam = cadParam & "pHastaFecha=""" & txtcodigo(27).Text & """|"
                                numParam = numParam + 1
                                cadParam = cadParam & "pFecha=""" & txtcodigo(32).Text & """|"
                                numParam = numParam + 1
                                
                                ConSubInforme = True
                            Else
                                cadParam = cadParam & "pSaltaPag=" & Check1(6).Value & "|"
                                numParam = numParam + 1
                            End If
                        End If
                    End If
                        
            End Select
            cadFormula = "{tmprfactsoc.codusu}=" & vUsu.Codigo
            ConSubInforme = True
            
            If OpcionListado = 9 And (Mid(cadNombreRPT, 1, 3) = "Cat" Or Mid(cadNombreRPT, 1, 3) = "Moi") Then
                ConSubInforme = True
            End If
            
            LlamarImprimir
        End If
    End If
End Sub

Private Sub cmdAceptarAnt_Click()
    '[Monica]11/03/2015: observaciones de factura
    ObsFactura = txtcodigo(68)
    
    Select Case vParamAplic.Cooperativa
        Case 0 'COOPERATIVA CATADAU
               ProcesoCatadau

        Case 1 'COOPERATIVA VALSUR
               ProcesoValsur
               
        Case 2 ' COOPERATIVA DE PICASSENT
               ProcesoPicassent
               
        Case 3 'COOPERATIVA MOIXENT
               'ProcesoMoixent
               If OpcionListado = 12 Or OpcionListado = 13 Or OpcionListado = 14 Or OpcionListado = 15 Then
                    ProcesoValsur
               End If
        Case 4 'COOPERATIVA DE ALZIRA
            ' en la coopoerativa de Alzira el proceso de liquidacion es el mismo que el de Valsur
            ' pero los calculos de importes que se hacen cuando se carga la temporal son distintos
            ' en cuanto a gastos
            
            '   ProcesoValsur
            '[Monica]05/03/2014: ahora se liquida y anticipa por tramos
                ProcesoCatadau
                
                
        Case 5 ' COOPERATIVA DE CASTELDUC
            Select Case OpcionListado
                Case 1, 2, 3
                    '[Monica]12/09/2012:antes los anticipos los hacia como en Valsur, un solo anticipo por campaña
                    ProcesoCatadau
                Case 12, 13, 14
                    ProcesoCatadau
            End Select
              ' ProcesoValsur
              
        Case 7 ' COOPERATIVA DE QUATRETONDA
            ' Partimos del proceso de catadau, pero ellos pueden liquidar tanto horto como almazara por el mismo punto
               ProcesoQuatretonda
               
        '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
        Case 9 ' NATURAL DE MONTAÑA
               ProcesoCatadau

        Case 14 ' COOPERATIVA BOLBAITE
               ProcesoCatadau

    End Select

End Sub


Private Sub ProcesoCatadau()
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
Dim b As Boolean
Dim Sql2 As String

Dim cadSelect1 As String
Dim Tabla1 As String


    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(12).Text)
        cHasta = Trim(txtcodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(20).Text)
        cHasta = Trim(txtcodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(20).Text, "N")
        If txtcodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(6).Text)
        cHasta = Trim(txtcodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
            
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        
        '[Monica]05/03/2014: entra alzira a facturarse por tramos
        If vParamAplic.Cooperativa = 4 Then ' Alzira
        '[Monica]24/06/2011: si es tercero de modulos en Alzira se liquida con los precios del socio
        '
        '           El orden es, primero se liquidan los socios no terceros con un precio y luego los socios terceros de modulos.
        '           Los socios terceros entidad se tratan como tales en la recepcion de socios terceros
        '
            'Socio que no sea tercero
            If Check1(11).Value = 0 Then
                If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
            Else
                ' socio tercero de modulos
                If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} = 1 and {rsocios.tipoirpf} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} = 1 and {rsocios.tipoirpf} = 0") Then Exit Sub
            End If
        Else
            'Socio que no sea tercero
            If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        End If
                
        
        '[Monica]03/06/2013: distinguimos entre entradas normales y entradas de p.integrado (solo para catadau)
        If vParamAplic.Cooperativa = 0 And (OpcionListado = 12 Or OpcionListado = 13 Or OpcionListado = 14) Then
        '[Monica]27/01/2016: cambiamos lo de la seleccion de las entradas
'                If Check1(16).Value = 1 Then ' solo entradas normales
'                    If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} = 0") Then Exit Sub
'                    If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} = 0") Then Exit Sub
'                Else ' solo producto integrado
'                    If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} = 2") Then Exit Sub
'                    If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} = 2") Then Exit Sub
'                End If
            If Check1(23).Value = 0 Then ' normales
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 0") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 0") Then Exit Sub
            End If
            If Check1(24).Value = 0 Then ' producto integrado
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 2") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 2") Then Exit Sub
            End If
            If Check1(25).Value = 0 Then ' venta campo
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1") Then Exit Sub
            End If
        
            If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 3 and {rhisfruta.tipoentr} <> 4 and {rhisfruta.tipoentr} <> 6") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 3 and {rhisfruta.tipoentr} <> 4 and {rhisfruta.tipoentr} <> 6") Then Exit Sub
        
        
        Else
            '[Monica]27/03/2013: nuevo tipo de entradas que tampoco hemos de liquidar (complementarias=siniestro) solo catadau
            '[Monica]30/11/2011: antes no estaba ni industria ni retirada
            'sólo entradas distintas de VENTA CAMPO y distintas de INDUSTRIA y distintas de RETIRADA
            If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3 and {rhisfruta.tipoentr} <> 4 and {rhisfruta.tipoentr} <> 6") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3 and {rhisfruta.tipoentr} <> 4 and {rhisfruta.tipoentr} <> 6") Then Exit Sub
        End If
        
        
        '++monica: 28/07/2009 dependiendo del tipo de recoleccion (0=coop 1=socio 2=todos)
        Select Case Combo1(2).ListIndex
            Case 0      ' recolectado cooperativa
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 0") Then Exit Sub
            Case 1      ' recolectado socio
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 1") Then Exit Sub
            Case 2      ' ambos
            
        End Select
        
        nTabla = "(((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN rhisfruta_clasif ON rhisfruta.numalbar = rhisfruta_clasif.numalbar) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodegaÇ
        
        cadSelect1 = cadSelect
        Tabla1 = nTabla
        
        
        Select Case OpcionListado
            Case 1 ' Listado de anticipos
                'Nombre fichero .rpt a Imprimir
                indRPT = 24 ' informe de anticipos
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatAnticipos.rpt"
                cadTitulo = "Informe de Anticipos"
            Case 2 ' Prevision de pago de anticipos
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    cadNombreRPT = "rPrevPagosAnt.rpt"
                Else ' agrupado por variedad
                    cadNombreRPT = "rPrevPagosAnt1.rpt"
                End If
                cadTitulo = "Previsión de Pago de Anticipos"
            
            Case 3 ' Facturación de Anticipos
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Anticipos"
            
            Case 12 ' Listado de Liquidacion
                'Nombre fichero .rpt a Imprimir
                indRPT = 26 ' informe de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatLiquidacion.rpt"
                cadTitulo = "Informe de Liquidación"
                
            Case 13 ' Prevision de pago de liquidacion
'[Monica]:09/09/2009 Parametrizamos el informe de prevision
'                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
'                    cadNombreRPT = "rPrevPagosLiq.rpt"
'                Else ' agrupado por variedad
'                    cadNombreRPT = "rPrevPagosLiq1.rpt"
'                End If

                'Nombre fichero .rpt a Imprimir
                indRPT = 33 ' informe de prevision de pago de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu 'rPrevPagosLiq.rpt
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    ' no hacemos nada dejamos el nombre de fichero como estaba
                    
                Else ' agrupado por variedad
                    cadNombreRPT = Replace(cadNombreRPT, "PrevPagosLiq.rpt", "PrevPagosLiq1.rpt")
                End If
                
                cadTitulo = "Previsión de Pago de Liquidación"
                
                '[Monica]27/01/2016: si es Catadau y es complementaria sale un subreport con los diferentes kilos agrupados
                '                    por variedad
                If vParamAplic.Cooperativa = 0 Then
                    cadParam = cadParam & "pComple=" & Check1(5).Value & "|"
                    numParam = numParam + 1
                End If
                
                
                '
            
            Case 14 ' Facturación de Liquidacion
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación"
                
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWhere = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        
        '[Monica]11/12/2013: en el caso de natural de montaña, si es liquidacion sacamos las distintas fechas de anticipos que
        '                    esten pendiente de descontar
        vFechas = ""
        If vParamAplic.Cooperativa = 9 Then
            If OpcionListado = 12 Or OpcionListado = 13 Or OpcionListado = 14 Then
                vFechas = CargarFechasAnticiposPendientes(nTabla, cadSelect)
                
                If vFechas <> "" Then
                    Set frmMens5 = New frmMensajes
                    
                    frmMens5.OpcionMensaje = 56
                    frmMens5.cadWhere = vFechas
                    frmMens5.Show vbModal
                    
                    Set frmMens5 = Nothing
                    
                    If vFechas = "-1" Then Exit Sub
                    
                End If
                
            End If
        End If
        
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
        
            Select Case OpcionListado
                Case 1, 2, 3
                    TipoPrec = 0 ' ANTICIPOS
                Case 12, 13, 14
                    TipoPrec = 1 ' LIQUIDACIONES
            End Select
            
            '[Monica]05/03/2014: añado todo lo de alzira aqui
            'comprobamos que los tipos de iva existen en la contabilidad de horto
            If Not ComprobarTiposIVA(nTabla, cadSelect) Then Exit Sub
            
            
            '[Monica]27/04/2011: de momento solo alzira comprobamos si los albaranes seccionado ya estan liquidados
            If vParamAplic.Cooperativa = 4 Then
                Dim CadenaAlbaranes As String
                
                CadenaAlbaranes = ""
                If Not AlbaranesFacturados(nTabla, cadSelect, CadenaAlbaranes) Then Exit Sub
                
                If Not AnyadirAFormula(cadSelect1, CadenaAlbaranes) Then Exit Sub
                
                ' volvemos a comprobar si hay albaranes pendientes de liquidar
                If Not HayRegParaInforme(Tabla1, cadSelect1) Then Exit Sub
            End If
            
            If HayPreciosVariedadesCatadau(TipoPrec, nTabla, cadSelect, Combo1(2).ListIndex) Then
                'D/H fecha
                cDesde = Trim(txtcodigo(6).Text)
                cHasta = Trim(txtcodigo(7).Text)
                cadDesde = CDate(cDesde)
                cadhasta = CDate(cHasta)
                cadAux = "{rprecios.fechaini}= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechaini}=" & DBSet(txtcodigo(6).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                cadAux = "{rprecios.fechafin}= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechafin}=" & DBSet(txtcodigo(7).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                
                If Check1(5).Value = 0 Then
                    ' si se trata de anticipos--> seleccionamos los precios de anticipos
                    ' sino los de liquidaciones
                    If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = " & TipoPrec) Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = " & TipoPrec) Then Exit Sub
                Else
                    If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = 3") Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = 3") Then Exit Sub
                End If
                
                Select Case OpcionListado
                    Case 1, 12
                        If CargarTemporalCatadau(Tabla1, cadSelect1, TipoPrec) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpliquidacion.codusu} = " & vUsu.Codigo) Then Exit Sub
                            
                            cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
                            numParam = numParam + 1
                            
                            ConSubInforme = True
                            
                            LlamarImprimir
                        End If
                        
                    
                    Case 2  '2 - listado de prevision de pagos de anticipos
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        If CargarTemporalAnticiposCatadau(Tabla1, cadSelect1) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                            ConSubInforme = True
                            
                            LlamarImprimir
                        End If
                        
                    Case 13 '13- listado de prevision de pagos de liquidaciones
                        If vParamAplic.Cooperativa = 4 Then
                            If CargarTemporalLiquidacionAlziraNew(Tabla1, cadSelect1) Then
                                cadFormula = ""
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                
                                ConSubInforme = True
                                
                                LlamarImprimir
                            End If
                        
                        Else
                            'catadau
                            If CargarTemporalLiquidacionCatadau(Tabla1, cadSelect1) Then
                                cadFormula = ""
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                
                                ConSubInforme = True
                                
                                LlamarImprimir
                            End If
                        End If
                    Case 3, 14 '3 .- factura de anticipos
                               '14.- factura de liquidaciones
                               
                        If CargarTemporalCatadau(Tabla1, cadSelect1, TipoPrec) Then
                            Nregs = TotalFacturasNew("tmpliquidacion", "codusu = " & vUsu.Codigo, "tmpliquidacion.codsocio")
                                
                            If Nregs <> 0 Then
'                                    If Not ComprobarTiposMovimiento(TipoPrec, nTabla, cadSelect) Then
'                                        Exit Sub
'                                    End If
                                
                                Me.Pb1.visible = True
                                Me.Pb1.Max = Nregs
                                Me.Pb1.Value = 0
                                Me.Refresh
                                b = False
                                If TipoPrec = 0 Then
                                    b = FacturacionAnticiposCatadau(Tabla1, cadSelect1, txtcodigo(15).Text, Me.Pb1)
                                Else
                                    '[Monica]07/02/2012: pasamos a la funcion si es o no liquidacion complementaria
                                    b = FacturacionLiquidacionesCatadau(Tabla1, cadSelect1, txtcodigo(15).Text, Me.Pb1, (Check1(5).Value = 1), txtcodigo(6).Text, txtcodigo(7).Text, vFechas, (Check1(21).Value = 1))
                                End If
                                If b Then
                                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                                   
                                    'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                                    If Me.Check1(2).Value Then
                                        cadFormula = ""
                                        cadParam = cadParam & "pFecFac= """ & txtcodigo(15).Text & """|"
                                        numParam = numParam + 1
                                        If TipoPrec = 0 Then
                                            cadParam = cadParam & "pTitulo= ""Resumen Facturación de Anticipos""|"
                                        Else
                                            cadParam = cadParam & "pTitulo= ""Resumen Facturación de Liquidaciones""|"
                                        End If
                                        numParam = numParam + 1
                                        
                                        FecFac = CDate(txtcodigo(15).Text)
                                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                        ConSubInforme = True
                                        
                                        LlamarImprimir
                                    End If
                                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS/LIQUIDACION
                                    If Me.Check1(3).Value Then
                                        cadFormula = ""
                                        cadSelect = ""
                                        If TipoPrec = 0 Then 'Tipo de Factura: Anticipo
                                            cadAux = "({stipom.tipodocu} = 1)"
                                        Else  'Tipo de Factura: Liquidación
                                            cadAux = "({stipom.tipodocu} = 2)"
                                        End If
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                        'Nº Factura
                                        If TipoPrec = 0 Then
                                            cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(2) & "])"
                                        Else
                                            cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(3) & "])"
                                        End If
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                         
                                        'Fecha de Factura
                                        FecFac = CDate(txtcodigo(15).Text)
                                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                                       
                                        indRPT = 23 'Impresion de facturas de socios
                                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                                        'Nombre fichero .rpt a Imprimir
                                        cadNombreRPT = nomDocu
                                        'Nombre fichero .rpt a Imprimir
                                        If TipoPrec = 0 Then
                                            cadTitulo = "Reimpresión de Facturas Anticipos"
                                        Else
                                            cadTitulo = "Reimpresión de Facturas Liquidaciones"
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
                        End If
                End Select
'            '++monica:27/07/2009
'            Else
'                MsgBox "No hay precios para las calidades en este rango. Revise.", vbExclamation
            End If
        End If
    End If

End Sub


Private Function CargarFechasAnticiposPendientes(vTabla As String, vSelect As String) As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cadena As String


    On Error GoTo eCargarFechasAnticiposPendientes
    
    vTabla = QuitarCaracterACadena(vTabla, "{")
    vTabla = QuitarCaracterACadena(vTabla, "}")
    If vSelect <> "" Then
        vSelect = QuitarCaracterACadena(vSelect, "{")
        vSelect = QuitarCaracterACadena(vSelect, "}")
        vSelect = QuitarCaracterACadena(vSelect, "_1")
    End If

    SQL = "select distinct fff.fecfactu from (rfactsoc_variedad fff inner join usuarios.stipom aaa on fff.codtipom = aaa.codtipom and aaa.tipodocu = 1) inner join rfactsoc ccc on ccc.codtipom = fff.codtipom and ccc.numfactu = fff.numfactu and ccc.fecfactu = fff.fecfactu   "
    SQL = SQL & " where fff.descontado = 0 and ccc.esanticipogasto = 0 and (fff.codvarie, ccc.codsocio, fff.codcampo)  in (select distinct rhisfruta.codvarie, rhisfruta.codsocio, rhisfruta.codcampo from " & vTabla & " where " & vSelect & ") "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cadena = ""
    
    While Not RS.EOF
        Cadena = Cadena & DBSet(RS.Fields(0).Value, "F") & ","
    
        RS.MoveNext
    Wend
    
    'quitamos la ultima coma
    If Len(Cadena) > 0 Then
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
    End If
    
    Set RS = Nothing
    
    CargarFechasAnticiposPendientes = Cadena
    Exit Function
    
eCargarFechasAnticiposPendientes:
    MuestraError Err.Number, "Cargar Fechas de Anticipos Pendientes", Err.Description
End Function


Private Sub ProcesoValsur()
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
Dim b As Boolean
Dim Sql2 As String


Dim MaxContador As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(12).Text)
        cHasta = Trim(txtcodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(20).Text)
        cHasta = Trim(txtcodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(20).Text, "N")
        If txtcodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(6).Text)
        cHasta = Trim(txtcodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
            
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        
        If vParamAplic.Cooperativa = 4 Then ' Alzira
        '[Monica]24/06/2011: si es tercero de modulos en Alzira se liquida con los precios del socio
        '
        '           El orden es, primero se liquidan los socios no terceros con un precio y luego los socios terceros de modulos.
        '            Los socios terceros entidad se tratan como tales en la recepcion de socios terceros
        '
            'Socio que no sea tercero
            If Check1(11).Value = 0 Then
                If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
            Else
                ' socio tercero de modulos
                If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} = 1 and {rsocios.tipoirpf} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} = 1 and {rsocios.tipoirpf} = 0") Then Exit Sub
            End If
        Else
            'Socio que no sea tercero
            If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        End If
        
        'sólo entradas distintas de VENTA CAMPO y distintas de INDUSTRIA y distintas de RETIRADA
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3 and {rhisfruta.tipoentr} <> 4") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3 and {rhisfruta.tipoentr} <> 4") Then Exit Sub
        
        '++monica: 28/07/2009 dependiendo del tipo de recoleccion (0=coop 1=socio 2=todos)
        Select Case Combo1(2).ListIndex
            Case 0      ' recolectado cooperativa
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 0") Then Exit Sub
            Case 1      ' recolectado socio
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 1") Then Exit Sub
            Case 2      ' ambos
            
        End Select
        
        nTabla = "(((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN rhisfruta_clasif ON rhisfruta.numalbar = rhisfruta_clasif.numalbar) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        Select Case OpcionListado
            Case 1 ' Listado de anticipos
                'Nombre fichero .rpt a Imprimir
                indRPT = 24 ' informe de anticipos
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatAnticipos.rpt"
                cadTitulo = "Informe de Anticipos"
            
            Case 2 ' Prevision de pago de anticipos
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    cadNombreRPT = "rPrevPagosAnt.rpt"
                Else
                    If Combo1(3).ListIndex = 1 Then ' agrupado por variedad
                        cadNombreRPT = "rPrevPagosAnt1.rpt"
                    Else ' por calidad
                        cadNombreRPT = "rPrevPagosAnt2.rpt"
                    End If
                End If
                cadTitulo = "Previsión de Pago de Anticipos"
            
            Case 3 ' Facturación de Anticipos
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Anticipos"
            
            Case 12 ' Listado de Liquidacion
                'Nombre fichero .rpt a Imprimir
                indRPT = 26 ' informe de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatLiquidacion.rpt"
                cadTitulo = "Informe de Liquidación"
                
            Case 13 ' Prevision de pago de liquidacion
                'Nombre fichero .rpt a Imprimir
                indRPT = 33 ' informe de prevision de pago de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"ValPrevPagosLiq.rpt"
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    ' no hacemos nada dejamos el nombre de fichero como estaba
                    
                Else
                    If Combo1(3).ListIndex = 1 Then ' agrupado por variedad
                        cadNombreRPT = Replace(cadNombreRPT, "PrevPagosLiq.rpt", "PrevPagosLiq1.rpt")
                    Else ' por calidad
                        cadNombreRPT = Replace(cadNombreRPT, "PrevPagosLiq.rpt", "PrevPagosLiq2.rpt")
                    End If
                End If
                
                cadTitulo = "Previsión de Pago de Liquidación"
                
            Case 14 ' Facturación de Liquidacion
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación"
                
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWhere = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
            Select Case OpcionListado
                Case 1, 2, 3
                    TipoPrec = 0 ' ANTICIPOS
                Case 12, 13, 14
                    TipoPrec = 1 ' LIQUIDACIONES
            End Select
            
            'comprobamos que los tipos de iva existen en la contabilidad de horto
            If Not ComprobarTiposIVA(nTabla, cadSelect) Then Exit Sub
            
            
            '[Monica]27/04/2011: de momento solo alzira comprobamos si los albaranes seccionado ya estan liquidados
            If vParamAplic.Cooperativa = 4 Then
                If Not AlbaranesFacturados(nTabla, cadSelect) Then Exit Sub
                ' volvemos a comprobar si hay albaranes pendientes de liquidar
                If Not HayRegParaInforme(nTabla, cadSelect) Then Exit Sub
            End If
            
            If HayPreciosVariedadesValsur(TipoPrec, nTabla, cadSelect, Combo1(2).ListIndex) Then
                'D/H fecha
                cDesde = Trim(txtcodigo(6).Text)
                cHasta = Trim(txtcodigo(7).Text)
                cadDesde = CDate(cDesde)
                cadhasta = CDate(cHasta)
                cadAux = "{rprecios.fechaini}= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechaini}=" & DBSet(txtcodigo(6).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                cadAux = "{rprecios.fechafin}= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechafin}=" & DBSet(txtcodigo(7).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                If Check1(5).Value = 0 Then
                    ' si se trata de anticipos--> seleccionamos los precios de anticipos
                    ' sino los de liquidaciones
                    If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = " & TipoPrec) Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = " & TipoPrec) Then Exit Sub
                Else
                    If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = 3") Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = 3") Then Exit Sub
                End If
                
                '02/09/2010
                cadAux = "{rprecios.contador} = (select max(p.contador) from rprecios p where p.codvarie = rhisfruta.codvarie and "
                cadAux = cadAux & " p.tipofact = " & TipoPrec & " and p.fechaini = " & DBSet(txtcodigo(6).Text, "F")
                cadAux = cadAux & " and p.fechafin = " & DBSet(txtcodigo(7).Text, "F") & ")"
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                Select Case OpcionListado
                    Case 1  '1 - informe de anticipos
                        'pasamos como parametro la fecha de anticipo
                        cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
                        numParam = numParam + 1
                        ConSubInforme = False
                        
                        InsertarTemporal (Variedades)
                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub

                        
                        LlamarImprimir
                    
                    Case 12 '12- informe de liquidaciones
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        Select Case vParamAplic.Cooperativa
                            Case 1, 3, 5  ' valsur / mogente
                                If CargarTemporalLiquidacionValsur(nTabla, cadSelect) Then
'                                    cadFormula = ""
                                    
                                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                                                        
                                    ConSubInforme = True
                                    'pasamos como parametro la fecha de anticipo
                                    cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
                                    numParam = numParam + 1
                                    cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
                                    numParam = numParam + 1
                                    LlamarImprimir
                                End If
                            
                            Case 2 ' Picassent
                                If CargarTemporalLiquidacionPicassent(nTabla, cadSelect) Then
'                                    cadFormula = ""
                                    If Not AnyadirAFormula(cadFormula, "{tmpliquidacion.codusu} = " & vUsu.Codigo) Then Exit Sub
                                                                        
                                    ConSubInforme = True
                                    'pasamos como parametro la fecha de anticipo
                                    cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
                                    numParam = numParam + 1
                                    cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
                                    numParam = numParam + 1
                                    LlamarImprimir
                                End If
                            
                            Case 4 ' Alzira
                                If CargarTemporalLiquidacionAlzira(nTabla, cadSelect, TipoPrec) Then
        '                            cadFormula = ""
        '                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                    ConSubInforme = True
                                    'pasamos como parametro la fecha de anticipo
                                    cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
                                    numParam = numParam + 1
                                    cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
                                    numParam = numParam + 1
                                    LlamarImprimir
                                End If
                        End Select
                    
                    Case 2  '2 - listado de prevision de pagos de anticipos
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        '[Monica]20/01/2012: Nuevo proceso para Alzira que hasta el momento no tenia anticipos
                        Select Case vParamAplic.Cooperativa
                            Case 2 ' Picassent
                                If Combo1(3).ListIndex = 2 Then
                                    If CargarTemporalAnticiposCalidadPicassent(nTabla, cadSelect) Then
                                        cadFormula = ""
                                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                        ConSubInforme = False
                                        
                                        LlamarImprimir
                                    End If
                                Else
                                    If CargarTemporalAnticiposPicassent(nTabla, cadSelect) Then
                                        cadFormula = ""
                                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                        ConSubInforme = True
                                        
                                        cadParam = cadParam & "pConBonifica=1|"
                                        numParam = numParam + 1
                                        LlamarImprimir
                                    End If
                                End If
                                                        
                            Case 4 ' Alzira
                                If CargarTemporalAnticiposAlzira(nTabla, cadSelect) Then
                                    cadFormula = ""
                                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                    ConSubInforme = False
                                    
                                    LlamarImprimir
                                End If
                            
                            Case Else
                                If CargarTemporalAnticiposValsur(nTabla, cadSelect) Then
                                    cadFormula = ""
                                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                    ConSubInforme = False
                                    
                                    LlamarImprimir
                                End If
                        End Select
                        
                    Case 13 '13- listado de prevision de pagos de liquidaciones
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        
                        Select Case vParamAplic.Cooperativa
                            Case 1, 3, 5   ' valsur / mogente
                                If CargarTemporalLiquidacionValsur(nTabla, cadSelect) Then
                                    cadFormula = ""
                                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                    ConSubInforme = False
                                    
                                    LlamarImprimir
                                End If
                            Case 2 ' Picassent
                                '[Monica]22/03/2012: indicamos en el informe si hacemos o no el descuento de comision segun el check1(13)
                                If Check1(13).Value = 1 Then
                                    cadParam = cadParam & "pTipo=""Cálculo con comisión""|"
                                Else
                                    cadParam = cadParam & "pTipo=""Cálculo sin comisión""|"
                                End If
                                numParam = numParam + 1
                                
                                If Combo1(3).ListIndex = 2 Then
                                    ' es igual que la cargatempporal de anticipos pero aqui coge los precios de liquidacion
                                    If CargarTemporalLiquidacionesCalidadPicassent(nTabla, cadSelect) Then
                                        cadFormula = ""
                                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                        ConSubInforme = False
                                        
                                        LlamarImprimir
                                    End If
                                Else
                                    If CargarTemporalLiquidacionPicassent(nTabla, cadSelect) Then
                                        cadFormula = ""
                                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                        ConSubInforme = False
                                        
                                        LlamarImprimir
                                    End If
                                End If
                                
                            Case 4 ' Alzira
                                If CargarTemporalLiquidacionAlzira(nTabla, cadSelect, TipoPrec) Then
                                    cadFormula = ""
                                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                    ConSubInforme = False
                                    
                                    LlamarImprimir
                                End If
                            
                        End Select
                    Case 3, 14 '3 .- factura de anticipos
                               '14.- factura de liquidaciones
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        Nregs = TotalFacturas(nTabla, cadSelect)
                        If Nregs <> 0 Then
                            If Not ComprobarTiposMovimiento(TipoPrec, nTabla, cadSelect) Then
                                Exit Sub
                            End If
                            
                            Me.Pb1.visible = True
                            Me.Pb1.Max = Nregs
                            Me.Pb1.Value = 0
                            Me.Refresh
                            b = False
                            Select Case vParamAplic.Cooperativa
                                Case 1, 3, 5  ' valsur / mogente
                                    If TipoPrec = 0 Then
                                        b = FacturacionAnticiposValsur(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1)
                                    Else
                                        b = FacturacionLiquidacionesValsur(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1, Check1(5).Value)
                                    End If
                                Case 4 ' alzira
                                    If TipoPrec = 0 Then
                                        '[Monica]20/01/2012: alzira no ha hecho hasta el momento anticipos
                                        'b = FacturacionAnticiposValsur(nTabla, cadSelect, txtcodigo(15).Text, Me.pb1)
                                        b = FacturacionAnticiposAlzira(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1)
                                    Else
                                        b = FacturacionLiquidacionesAlzira(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1, TipoPrec)
                                    End If
                                Case 2 ' Picassent
                                    If TipoPrec = 0 Then
                                        b = FacturacionAnticiposPicassent(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1, Check1(14).Value = 1)
                                    Else
                                        b = FacturacionLiquidacionesPicassent(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1, TipoPrec, Check1(14).Value = 1)
                                    End If
                                
                                
                            End Select
                            If b Then
                                MsgBox "Proceso realizado correctamente.", vbExclamation
                                               
                                'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                                If Me.Check1(2).Value Then
                                    cadFormula = ""
                                    cadParam = cadParam & "pFecFac= """ & txtcodigo(15).Text & """|"
                                    numParam = numParam + 1
                                    If TipoPrec = 0 Then
                                        cadParam = cadParam & "pTitulo= ""Resumen Facturación de Anticipos""|"
                                    Else
                                        cadParam = cadParam & "pTitulo= ""Resumen Facturación de Liquidaciones""|"
                                    End If
                                    numParam = numParam + 1
                                    
                                    FecFac = CDate(txtcodigo(15).Text)
                                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                    ConSubInforme = True
                                    
                                    LlamarImprimir
                                End If
                                'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS/LIQUIDACION
                                If Me.Check1(3).Value Then
                                    cadFormula = ""
                                    cadSelect = ""
                                    If TipoPrec = 0 Then 'Tipo de Factura: Anticipo
                                        cadAux = "({stipom.tipodocu} = 1)"
                                    Else  'Tipo de Factura: Liquidación
                                        cadAux = "({stipom.tipodocu} = 2)"
                                    End If
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                    'Nº Factura
                                    If TipoPrec = 0 Then
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(2) & "])"
                                    Else
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(3) & "])"
                                    End If
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                                    'Fecha de Factura
                                    FecFac = CDate(txtcodigo(15).Text)
                                    cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                    cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                    If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                                    indRPT = 23 'Impresion de facturas de socios
                                    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                                    'Nombre fichero .rpt a Imprimir
                                    cadNombreRPT = nomDocu
                                    'Nombre fichero .rpt a Imprimir
                                    If TipoPrec = 0 Then
                                        cadTitulo = "Reimpresión de Facturas Anticipos"
                                    Else
                                        cadTitulo = "Reimpresión de Facturas Liquidaciones"
                                    End If
                                    ConSubInforme = True

                                    If indRPT = 23 And vParamAplic.Cooperativa = 2 Then
                                        Dim PrecioApor As Double
                                        PrecioApor = DevuelveValor("select min(precio) from raporreparto")
                                        
                                        cadParam = cadParam & "pPrecioApor=""" & Replace(Format(PrecioApor, "#0.000000"), ",", ".") & """|"
                                        numParam = numParam + 1
                                    End If

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
                MsgBox "No hay precios para las calidades en este rango. Revise.", vbExclamation
            End If
        End If
    End If
End Sub

Private Sub ProcesoQuatretonda()
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
Dim b As Boolean
Dim Sql2 As String

Dim cadSelect1 As String
Dim Tabla1 As String

Dim SqlIva As String
Dim PorcIva As Currency
Dim vPorcIva As String
Dim vSocio As cSocio
Dim vSeccion As CSeccion

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(12).Text)
        cHasta = Trim(txtcodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(20).Text)
        cHasta = Trim(txtcodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(20).Text, "N")
        If txtcodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(6).Text)
        cHasta = Trim(txtcodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
            
'[Monica]25/06/2012: quitamos de aqui la seccion, la ponemos mas abajo
'        'SECCION
'        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
'        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        'Socio que no sea tercero
        If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        
        '[Monica]30/11/2011: en quatretonda en el informe de liquidacion se cogen todos los kilos incluidos los de retirada
'        If OpcionListado = 12 Then
            'sólo entradas distintas de VENTA CAMPO y distintas de INDUSTRIA
            If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3 ") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3 ") Then Exit Sub
'        Else
'            'sólo entradas distintas de VENTA CAMPO y distintas de INDUSTRIA y distintas de RETIRADA
'            If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3 and {rhisfruta.tipoentr} <> 4") Then Exit Sub
'            If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3 and {rhisfruta.tipoentr} <> 4") Then Exit Sub
'        End If
        
        '++monica: 28/07/2009 dependiendo del tipo de recoleccion (0=coop 1=socio 2=todos)
        Select Case Combo1(2).ListIndex
            Case 0      ' recolectado cooperativa
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 0") Then Exit Sub
            Case 1      ' recolectado socio
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 1") Then Exit Sub
            Case 2      ' ambos
            
        End Select
        
        nTabla = "(((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN rhisfruta_clasif ON rhisfruta.numalbar = rhisfruta_clasif.numalbar) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
        '[Monica]30/11/2011: en quatretonda se pueden liquidar lo de aceituna
        nTabla = nTabla & " and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodegaÇ
        
        cadSelect1 = cadSelect
        Tabla1 = nTabla
        
        
        Select Case OpcionListado
            Case 1 ' Listado de anticipos
                'Nombre fichero .rpt a Imprimir
                indRPT = 24 ' informe de anticipos
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatAnticipos.rpt"
                cadTitulo = "Informe de Anticipos"
            Case 2 ' Prevision de pago de anticipos
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    cadNombreRPT = "rPrevPagosAnt.rpt"
                Else ' agrupado por variedad
                    cadNombreRPT = "rPrevPagosAnt1.rpt"
                End If
                cadTitulo = "Previsión de Pago de Anticipos"
            
            Case 3 ' Facturación de Anticipos
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Anticipos"
            
            Case 12 ' Listado de Liquidacion
                'Nombre fichero .rpt a Imprimir
                indRPT = 26 ' informe de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatLiquidacion.rpt"
                cadTitulo = "Informe de Liquidación"
                
            Case 13 ' Prevision de pago de liquidacion
'[Monica]:09/09/2009 Parametrizamos el informe de prevision
'                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
'                    cadNombreRPT = "rPrevPagosLiq.rpt"
'                Else ' agrupado por variedad
'                    cadNombreRPT = "rPrevPagosLiq1.rpt"
'                End If

                'Nombre fichero .rpt a Imprimir
                indRPT = 33 ' informe de prevision de pago de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu 'rPrevPagosLiq.rpt
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    ' no hacemos nada dejamos el nombre de fichero como estaba
                    
                Else ' agrupado por variedad
                    cadNombreRPT = Replace(cadNombreRPT, "PrevPagosLiq.rpt", "PrevPagosLiq1.rpt")
                End If
                
                cadTitulo = "Previsión de Pago de Liquidación"
            
            Case 14 ' Facturación de Liquidacion
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación"
                
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWhere = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        '[Monica]30/11/2011: hemos de controlar que no se incluyan a la vez facturas de recoleccion con facturas de almazara
        ' para que en la integracion contable integremos a la contabilidad de fruta o a la de bodega
        ' aqui entraran todas las facturas como FAL aunque algunas tengan variedades de almazara
        If HayVariedadesAlmazaraconHorto(nTabla, cadSelect) Then
            MsgBox "Las variedades seleccionadas deben ser todas de Horto o todas de Almazara. Revise.", vbExclamation
            Exit Sub
        End If
        
'[Monica]25/06/2012: metemos la seccion que corresponda segun sea de horto o de almazara
        Dim Seccion As String
        'SECCION
        If HayVariedadesAlmazara(nTabla, cadSelect) Then
            If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.SeccionAlmaz) Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.SeccionAlmaz) Then Exit Sub
            If Not AnyadirAFormula(cadSelect1, "{rsocios_seccion.codsecci} = " & vParamAplic.SeccionAlmaz) Then Exit Sub
            
            Seccion = vParamAplic.SeccionAlmaz
        Else
            If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
            If Not AnyadirAFormula(cadSelect1, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
            
            Seccion = vParamAplic.Seccionhorto
        End If
        
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
        
            Select Case OpcionListado
                Case 1, 2, 3
                    TipoPrec = 0 ' ANTICIPOS
                Case 12, 13, 14
                    TipoPrec = 1 ' LIQUIDACIONES
            End Select
            
            If HayPreciosVariedadesCatadau(TipoPrec, nTabla, cadSelect, Combo1(2).ListIndex) Then
            
                'D/H fecha
                cDesde = Trim(txtcodigo(6).Text)
                cHasta = Trim(txtcodigo(7).Text)
                cadDesde = CDate(cDesde)
                cadhasta = CDate(cHasta)
                cadAux = "{rprecios.fechaini}= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechaini}=" & DBSet(txtcodigo(6).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                cadAux = "{rprecios.fechafin}= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechafin}=" & DBSet(txtcodigo(7).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                
                If Check1(5).Value = 0 Then
                    ' si se trata de anticipos--> seleccionamos los precios de anticipos
                    ' sino los de liquidaciones
                    If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = " & TipoPrec) Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = " & TipoPrec) Then Exit Sub
                Else
                    If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = 3") Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = 3") Then Exit Sub
                End If
                
                Select Case OpcionListado
                    Case 1, 12
                        If CargarTemporalQuatretonda(Tabla1, cadSelect1, TipoPrec) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpliquidacion.codusu} = " & vUsu.Codigo) Then Exit Sub
                            
                            cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
                            numParam = numParam + 1
                            
                            '[Monica]14/02/2012: pasamos el iva de la seccion de horto para el calculo del importe con iva
                            Set vSeccion = New CSeccion
                            '[Monica]25/06/2012: seccion
                            'If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                            If vSeccion.LeerDatos(Seccion) Then
                                If vSeccion.AbrirConta Then
                                    SqlIva = "select min(codsocio) from tmpliquidacion where codusu = " & vUsu.Codigo
                                    
                                    Set vSocio = New cSocio
                                    '[Monica]25/06/2012: seccion
                                    'If vSocio.LeerDatosSeccion(CStr(DevuelveValor(SqlIva)), vParamAplic.Seccionhorto) Then
                                    If vSocio.LeerDatosSeccion(CStr(DevuelveValor(SqlIva)), Seccion) Then
                                         vPorcIva = ""
                                         vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                                    End If
                                    Set vSocio = Nothing
                                End If
                            End If
                            Set vSeccion = Nothing
                            
'                            PorcIva = 0
'                            If vPorcIva <> "" Then PorcIva = CCur(vPorcIva)
                            cadParam = cadParam & "pPorciva=" & TransformaComasPuntos(ImporteSinFormato(vPorcIva)) & "|"
                            ' fin de iva
                            
                            
                            ConSubInforme = True
                            
                            LlamarImprimir
                        End If
                    
                    Case 2  '2 - listado de prevision de pagos de anticipos
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        If CargarTemporalAnticiposCatadau(Tabla1, cadSelect1) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                            ConSubInforme = False
                            
                            LlamarImprimir
                        End If
                        
                    Case 13 '13- listado de prevision de pagos de liquidaciones
                        'catadau
                        If CargarTemporalLiquidacionQuatretonda(Tabla1, cadSelect1, Seccion) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                            
                            ConSubInforme = True
                            
                            LlamarImprimir
                        End If
                        
                    Case 3, 14 '3 .- factura de anticipos
                               '14.- factura de liquidaciones
                               
                        If CargarTemporalQuatretonda(Tabla1, cadSelect1, TipoPrec) Then
                            Nregs = TotalFacturasNew("tmpliquidacion", "codusu = " & vUsu.Codigo, "tmpliquidacion.codsocio")
                                
                            If Nregs <> 0 Then
                                
                                Me.Pb1.visible = True
                                Me.Pb1.Max = Nregs
                                Me.Pb1.Value = 0
                                Me.Refresh
                                b = False
                                If TipoPrec = 0 Then
                                    b = FacturacionAnticiposCatadau(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1)
                                Else
                                   '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
                                    b = FacturacionLiquidacionesQuatretonda(Tabla1, cadSelect1, txtcodigo(15).Text, Me.Pb1, (TipoPrec = 3), Seccion)
                                End If
                                If b Then
                                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                                   
                                    'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                                    If Me.Check1(2).Value Then
                                        cadFormula = ""
                                        cadParam = cadParam & "pFecFac= """ & txtcodigo(15).Text & """|"
                                        numParam = numParam + 1
                                        If TipoPrec = 0 Then
                                            cadParam = cadParam & "pTitulo= ""Resumen Facturación de Anticipos""|"
                                        Else
                                            cadParam = cadParam & "pTitulo= ""Resumen Facturación de Liquidaciones""|"
                                        End If
                                        numParam = numParam + 1
                                        
                                        FecFac = CDate(txtcodigo(15).Text)
                                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                        ConSubInforme = False
                                        
                                        LlamarImprimir
                                    End If
                                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS/LIQUIDACION
                                    If Me.Check1(3).Value Then
                                        cadFormula = ""
                                        cadSelect = ""
                                        If TipoPrec = 0 Then 'Tipo de Factura: Anticipo
                                            cadAux = "({stipom.tipodocu} = 1)"
                                        Else  'Tipo de Factura: Liquidación
                                            cadAux = "({stipom.tipodocu} = 2)"
                                        End If
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                        'Nº Factura
                                        If TipoPrec = 0 Then
                                            cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(2) & "])"
                                        Else
                                            cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(3) & "])"
                                        End If
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                         
                                        'Fecha de Factura
                                        FecFac = CDate(txtcodigo(15).Text)
                                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                                       
                                        indRPT = 23 'Impresion de facturas de socios
                                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                                        'Nombre fichero .rpt a Imprimir
                                        cadNombreRPT = nomDocu
                                        'Nombre fichero .rpt a Imprimir
                                        If TipoPrec = 0 Then
                                            cadTitulo = "Reimpresión de Facturas Anticipos"
                                        Else
                                            cadTitulo = "Reimpresión de Facturas Liquidaciones"
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
                        End If
                End Select
'            '++monica:27/07/2009
'            Else
'                MsgBox "No hay precios para las calidades en este rango. Revise.", vbExclamation
            End If
        End If
    End If

End Sub



Private Sub cmdAceptarAntGastos_Click()
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
Dim b As Boolean
Dim Sql2 As String

Dim cadSelect1 As String

    '[Monica]11/03/2015: observaciones de factura
    ObsFactura = txtcodigo(68)


    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(12).Text)
        cHasta = Trim(txtcodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(20).Text)
        cHasta = Trim(txtcodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(20).Text, "N")
        If txtcodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(6).Text)
        cHasta = Trim(txtcodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
            
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        'Socio que no sea tercero
        If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        
        
        'sólo entradas distintas de VENTA CAMPO
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1") Then Exit Sub
        
        
        'sólo entradas recolectadas por socio
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 1") Then Exit Sub
        
        nTabla = "((((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN rhisfruta_clasif ON rhisfruta.numalbar = rhisfruta_clasif.numalbar) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        Select Case OpcionListado
            Case 2 ' Prevision de pago de anticipos gastos recoleccion
                cadNombreRPT = "rPrevPagosAntGastos.rpt"
                cadTitulo = "Previsión Pago de Anticipos Gastos"
            
            Case 3 ' Facturación de Anticipos de Gastos de recoleccion
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Anticipos Gastos"
            
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWhere = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
        
            Select Case OpcionListado
                Case 2  '2 - listado de prevision de pagos de anticipos
                    cadSelect1 = " rhisfruta.tipoentr <> 1 and rhisfruta.recolect = 1 "
                    If txtcodigo(6).Text <> "" Then cadSelect1 = cadSelect1 & " and rhisfruta.fecalbar >=" & DBSet(txtcodigo(6).Text, "F")
                    If txtcodigo(7).Text <> "" Then cadSelect1 = cadSelect1 & " and rhisfruta.fecalbar <=" & DBSet(txtcodigo(7).Text, "F")
                    
                    If CargarTemporalAnticiposGastos(nTabla, cadSelect, cadSelect1) Then
                        cadFormula = ""
                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                        ConSubInforme = False
                        
                        LlamarImprimir
                    End If
                    
                Case 3  '3 .- factura de anticipos de gastos
                    TipoPrec = 0 ' son anticipos
                    
                    Nregs = TotalFacturas(nTabla, cadSelect)
                    If Nregs <> 0 Then
                        If Not ComprobarTiposMovimiento(TipoPrec, nTabla, cadSelect) Then
                            Exit Sub
                        End If
                        
                        Me.Pb1.visible = True
                        Me.Pb1.Max = Nregs
                        Me.Pb1.Value = 0
                        Me.Refresh
                        
                        cadSelect1 = " rhisfruta.tipoentr <> 1 and rhisfruta.recolect = 1 "
                        If txtcodigo(6).Text <> "" Then cadSelect1 = cadSelect1 & " and rhisfruta.fecalbar >=" & DBSet(txtcodigo(6).Text, "F")
                        If txtcodigo(7).Text <> "" Then cadSelect1 = cadSelect1 & " and rhisfruta.fecalbar <=" & DBSet(txtcodigo(7).Text, "F")
                        
                        
                        b = FacturacionAnticiposGastos(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1, cadSelect1)
                        
                        If b Then
                            MsgBox "Proceso realizado correctamente.", vbExclamation
                                           
                            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS GASTOS
                            If Me.Check1(2).Value Then
                                cadFormula = ""
                                cadParam = cadParam & "pFecFac= """ & txtcodigo(15).Text & """|"
                                numParam = numParam + 1
                                If TipoPrec = 0 Then
                                    cadParam = cadParam & "pTitulo= ""Resumen Facturación de Anticipos Gastos""|"
                                End If
                                numParam = numParam + 1
                                
                                FecFac = CDate(txtcodigo(15).Text)
                                cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                ConSubInforme = False
                                
                                LlamarImprimir
                            End If
                            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS GASTOS
                            If Me.Check1(3).Value Then
                                cadFormula = ""
                                cadSelect = ""
                                If TipoPrec = 0 Then 'Tipo de Factura: Anticipo
                                    cadAux = "({stipom.tipodocu} = 1)"
                                End If
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                'Nº Factura
                                If TipoPrec = 0 Then
                                    cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(2) & "])"
                                Else
                                    cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(3) & "])"
                                End If
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                 
                                'Fecha de Factura
                                FecFac = CDate(txtcodigo(15).Text)
                                cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                               
                                indRPT = 23 'Impresion de facturas de socios
                                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                                'Nombre fichero .rpt a Imprimir
                                cadNombreRPT = nomDocu
                                'Nombre fichero .rpt a Imprimir
                                If TipoPrec = 0 Then
                                    cadTitulo = "Reimpresión de Facturas Anticipos"
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
                        MsgBox "No hay Gastos Recolección a facturar.", vbExclamation
                    End If
            End Select
        End If
    End If
End Sub


Private Sub cmdAceptarAntGene_Click()
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
Dim b As Boolean
Dim Sql2 As String
Dim Sql3 As String

Dim cadSelect1 As String

    '[Monica]11/03/2015: observaciones de factura
    ObsFactura = txtcodigo(68)


    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        If Check1(12).Value = 1 And ComprobarCero(txtcodigo(59).Text) = 0 Then
            MsgBox "Debe introducir obligatoriamente la cantidad de kilos retirados.", vbExclamation
            PonerFoco txtcodigo(59)
            Exit Sub
        End If
        
        
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(12).Text)
        cHasta = Trim(txtcodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rclasifica.codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        Sql3 = ""
        If txtcodigo(12).Text <> "" Then Sql3 = Sql3 & " and rsocios.codsocio >=" & DBSet(txtcodigo(12).Text, "N")
        If txtcodigo(13).Text <> "" Then Sql3 = Sql3 & " and rsocios.codsocio <=" & DBSet(txtcodigo(13).Text, "N")
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(20).Text)
        cHasta = Trim(txtcodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(20).Text, "N")
        If txtcodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(6).Text)
        cHasta = Trim(txtcodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rclasifica.fechaent}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
            
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        'Socio que no sea tercero
        If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        
        
        'sólo entradas distintas de VENTA CAMPO
        If Not AnyadirAFormula(cadSelect, "{rclasifica.tipoentr} <> 1") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rclasifica.tipoentr} <> 1") Then Exit Sub
        
       
        nTabla = "((((rclasifica INNER JOIN rsocios ON rclasifica.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN variedades ON rclasifica.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        Select Case OpcionListado
            Case 2 ' Prevision de pago de anticipos generico
                cadNombreRPT = "rPrevPagosAntGene.rpt"
                If Check1(12).Value = 0 Then
                    cadTitulo = "Previsión Pago Anticipos Genéricos"
                Else
                    cadTitulo = "Previsión Pago Anticipos Retirada"
                    '[Monica]23/12/2014:VR
                    If Check1(22).Value = 1 Then cadTitulo = cadTitulo & " VR"
                End If
                cadParam = cadParam & "pTitulo=""" & cadTitulo & """|"
                numParam = numParam + 1
            
            Case 3 ' Facturación de Anticipos de Genericos de recoleccion
                cadNombreRPT = "rResumFacturas.rpt"
                
                If Check1(12).Value = 0 Then
                    cadTitulo = "Resumen de Facturas de Anticipos Genéricos"
                Else
                    cadTitulo = "Resumen de Facturas de Anticipos Retirada"
                    '[Monica]23/12/2014:VR
                    If Check1(22).Value = 1 Then cadTitulo = cadTitulo & " VR"
                    
                End If
                cadParam = cadParam & "pTitulo=""" & cadTitulo & """|"
                numParam = numParam + 1
            
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWhere = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'solo los anticipos de retirada seleccionamos que socios queremos facturar
        If Check1(12).Value = 1 Then
            Set frmMens3 = New frmMensajes
            
            frmMens3.OpcionMensaje = 9
            frmMens3.cadWhere = " rsocios_seccion.codsecci = " & vParamAplic.Seccionhorto & " and rsocios.tipoprod <> 1" & Sql3
            frmMens3.Show vbModal
            
            Set frmMens3 = Nothing
        End If
        
        ' insertamos en la tabla intermedia de liquidacion lo que vamos a facturar
        b = InsertarTablaIntermedia(nTabla, cadSelect, False)
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If b And HayRegParaInforme("tmpliquidacion", "codusu = " & vUsu.Codigo) Then
        
            Select Case OpcionListado
                Case 2  '2 - listado de prevision de pagos de anticipos
                    '[Monica]18/10/2011: si check1(12).value = 1 indicamos que el anticipo es de retirada
                    '                    si check1(12).value = 0 indicamos que el anticipo es generico
                    If CargarTemporalAnticiposGenericos(nTabla, cadSelect, False, Check1(12).Value = 1) Then
                        
                        cadFormula = ""
                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                        ConSubInforme = True
                        
                        LlamarImprimir
                    End If
                    
                Case 3  '3 .- factura de anticipos de gastos
                    TipoPrec = 0 ' son anticipos
                    
                    Nregs = TotalFacturasNew("tmpliquidacion", "codusu = " & vUsu.Codigo, "codsocio")
                    If Nregs <> 0 Then
                        If Not ComprobarTiposMovimiento(TipoPrec, "tmpliquidacion inner join rsocios on tmpliquidacion.codsocio = rsocios.codsocio ", "codusu = " & vUsu.Codigo, Check1(22).Value = 1) Then
                            Exit Sub
                        End If
                        
                        Me.Pb1.visible = True
                        Me.Pb1.Max = Nregs
                        Me.Pb1.Value = 0
                        Me.Refresh
                        
                        b = FacturacionAnticiposGenerico("tmpliquidacion", "codigo = " & vUsu.Codigo, txtcodigo(15).Text, Me.Pb1, txtcodigo(6).Text, txtcodigo(7).Text, Check1(12).Value = 1, Check1(22).Value = 1)
                        
                        If b Then
                            MsgBox "Proceso realizado correctamente.", vbExclamation
                                           
                            'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS GASTOS
                            If Me.Check1(2).Value Then
                                cadFormula = ""
                                cadParam = cadParam & "pFecFac= """ & txtcodigo(15).Text & """|"
                                numParam = numParam + 1
                                If TipoPrec = 0 Then
                                    cadParam = cadParam & "pTitulo= ""Resumen Facturación de Anticipos Genérico""|"
                                End If
                                numParam = numParam + 1
                                
                                FecFac = CDate(txtcodigo(15).Text)
                                cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                ConSubInforme = False
                                
                                LlamarImprimir
                            End If
                            'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS GENERICO
                            If Me.Check1(3).Value Then
                                cadFormula = ""
                                cadSelect = ""
                                If TipoPrec = 0 Then 'Tipo de Factura: Anticipo
                                    cadAux = "({stipom.tipodocu} = 1)"
                                End If
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                'Nº Factura
                                If TipoPrec = 0 Then
                                    cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(2) & "])"
                                Else
                                    cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(3) & "])"
                                End If
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                 
                                'Fecha de Factura
                                FecFac = CDate(txtcodigo(15).Text)
                                cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                               
                                indRPT = 23 'Impresion de facturas de socios
                                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                                'Nombre fichero .rpt a Imprimir
                                cadNombreRPT = nomDocu
                                'Nombre fichero .rpt a Imprimir
                                If TipoPrec = 0 Then
                                    cadTitulo = "Reimpresión de Facturas Anticipos"
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
                        MsgBox "No hay Gastos Recolección a facturar.", vbExclamation
                    End If
            End Select
        End If
    End If

End Sub

Private Sub cmdAceptarLiqIndustria_Click()
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
Dim Tabla1 As String


Dim Nregs As Long
Dim FecFac As Date
Dim TipoPrec As Byte ' 0 anticipos
                     ' 1 liquidaciones
Dim b As Boolean
Dim Sql2 As String

Dim cadSelect1 As String

Dim CadenaAlbaranes As String
    
    '[Monica]11/03/2015: observaciones de factura
    ObsFactura = txtcodigo(68)

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(12).Text)
        cHasta = Trim(txtcodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(20).Text)
        cHasta = Trim(txtcodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(20).Text, "N")
        If txtcodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(6).Text)
        cHasta = Trim(txtcodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
            
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
'[Monica] 28/12/2009 : quito la condicion de que el socio no sea tercero (solo para liquidacion de industria)
'        'Socio que no sea tercero
'        If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
'        If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        
        'sólo entradas de Industria directa
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} = 3") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} = 3") Then Exit Sub
        
        
        nTabla = "(((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN rhisfruta_clasif ON rhisfruta.numalbar = rhisfruta_clasif.numalbar) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        cadSelect1 = cadSelect
        Tabla1 = nTabla
        
        
        Select Case OpcionListado
            '[Monica]23/05/2013: añadimos el informe de Liquidacion en ppio solo para Catadau
            Case 12 ' Informe de liquidacion
                'Nombre fichero .rpt a Imprimir
                indRPT = 26 ' informe de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatLiquidacion.rpt"
                cadTitulo = "Informe de Liquidación"
            
            
            Case 13 ' Prevision de pago de liquidacion de industria
                'Nombre fichero .rpt a Imprimir
                indRPT = 37 ' informe de prevision de pago de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu 'AlzPrevPagosLiqInd.rpt
                
                cadTitulo = "Previsión de Pago de Liquidación Industria"
            
            Case 14 ' Facturación de Liquidacion de Industria
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación Industria"
                
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWhere = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
        
            TipoPrec = 1 ' LIQUIDACIONES
            
            '[Monica]27/12/2012: de momento solo alzira comprobamos si los albaranes seccionado ya estan liquidados
            CadenaAlbaranes = ""
            If vParamAplic.Cooperativa = 4 Then
                If Not AlbaranesFacturados(nTabla, cadSelect, CadenaAlbaranes) Then Exit Sub
                ' volvemos a comprobar si hay albaranes pendientes de liquidar
                If Not HayRegParaInforme(nTabla, cadSelect) Then Exit Sub
                cadSelect1 = cadSelect
            End If
            
            If HayPreciosVariedadesIndustria(nTabla, cadSelect) Then
                'D/H fecha
                cDesde = Trim(txtcodigo(6).Text)
                cHasta = Trim(txtcodigo(7).Text)
                cadDesde = CDate(cDesde)
                cadhasta = CDate(cHasta)
                cadAux = "{rprecios.fechaini}= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechaini}=" & DBSet(txtcodigo(6).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                cadAux = "{rprecios.fechafin}= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechafin}=" & DBSet(txtcodigo(7).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                
                If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = 2") Then Exit Sub
                If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = 2") Then Exit Sub
                
                Select Case OpcionListado
                    Case 12 '[Monica]23/05/2013: Informe de liquidacion de industria
                        If CargarTemporalIndustria(Tabla1, cadSelect1) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpliquidacion.codusu} = " & vUsu.Codigo) Then Exit Sub
                            
                            cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
                            numParam = numParam + 1
                            
                            ConSubInforme = True
                            
                            LlamarImprimir
                        End If
                    
                    Case 13 '13- listado de prevision de pagos de liquidaciones industria
                        'catadau
                        If CargarTemporalLiquidacionIndustria(Tabla1, cadSelect1) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                            
'                            cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
'                            numParam = numParam + 1
                            
                            ConSubInforme = True
                            
                            LlamarImprimir
                        End If
                        
                    Case 14 '14.- factura de liquidaciones de industria (una factura por campo)
                        If CargarTemporalIndustria(Tabla1, cadSelect1) Then
                            Nregs = TotalFacturasNew("tmpliquidacion", "codusu = " & vUsu.Codigo, "tmpliquidacion.codsocio,tmpliquidacion.codcampo")
                                
                            If Nregs <> 0 Then
'                                    If Not ComprobarTiposMovimiento(TipoPrec, nTabla, cadSelect) Then
'                                        Exit Sub
'                                    End If
                                
                                Me.Pb1.visible = True
                                Me.Pb1.Max = Nregs
                                Me.Pb1.Value = 0
                                Me.Refresh
                                b = FacturacionLiquidacionIndustria(nTabla, cadSelect, txtcodigo(15).Text, Me.Pb1, CadenaAlbaranes)
                                If b Then
                                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                                   
                                    'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                                    If Me.Check1(2).Value Then
                                        cadFormula = ""
                                        cadParam = cadParam & "pFecFac= """ & txtcodigo(15).Text & """|"
                                        numParam = numParam + 1
                                        cadParam = cadParam & "pTitulo= ""Resumen Facturación Liquidación Industria""|"
                                        numParam = numParam + 1
                                        
                                        FecFac = CDate(txtcodigo(15).Text)
                                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                        ConSubInforme = False
                                        
                                        LlamarImprimir
                                    End If
                                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE LIQUIDACION INDUSTRIA
                                    If Me.Check1(3).Value Then
                                        cadFormula = ""
                                        cadSelect = ""
                                        cadAux = "({stipom.tipodocu} = 2)"
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                        'Nº Factura
                                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(3) & "])"
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                         
                                        'Fecha de Factura
                                        FecFac = CDate(txtcodigo(15).Text)
                                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                                       
                                        indRPT = 38 'Impresion de facturas de socios
                                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                                        'Nombre fichero .rpt a Imprimir
                                        cadNombreRPT = nomDocu
                                        'Nombre fichero .rpt a Imprimir
                                        cadTitulo = "Reimpresión de Facturas Liquidaciones Industria"
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
                        End If
                End Select
'            '++monica:27/07/2009
'            Else
'                MsgBox "No hay precios para las calidades en este rango. Revise.", vbExclamation
            End If
        End If
    End If

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


    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'Tipo de movimiento:
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
        If TipoFacturaOk Then
            ' quitamos la ultima coma
            Tipos = "{rfactsoc.codtipom} in (" & Mid(Tipos, 1, Len(Tipos) - 1) & ")"
            If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
            Tipos = Replace(Replace(Tipos, "(", "["), ")", "]")
            If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    'D/H Cliente
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Nro de Factura
    cDesde = Trim(txtcodigo(4).Text)
    cHasta = Trim(txtcodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    If HayRegistros(Tabla, cadSelect) Then
        If Industria Then
            indRPT = 38 'Impresion de Factura Socio de Industria
            ConSubInforme = False
            cadTitulo = "Reimpresión de Facturas Socios Industria"
        ElseIf Bodega Then
            indRPT = 42 'Impresion de Factura Socio de Bodega
            ConSubInforme = True
            cadTitulo = "Reimpresión de Facturas Socios Bodega"
        Else
            indRPT = 23 'Impresion de Factura Socio
            ConSubInforme = True
            cadTitulo = "Reimpresión de Facturas Socios"
        End If
        
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
          
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = nomDocu
          
        If indRPT = 23 And vParamAplic.Cooperativa = 2 Then
            Dim PrecioApor As Double
            PrecioApor = DevuelveValor("select min(precio) from raporreparto")
            
            cadParam = cadParam & "pPrecioApor=""" & Replace(Format(PrecioApor, "#0.000000"), ",", ".") & """|"
            numParam = numParam + 1
        End If
        
        '[Monica]28/01/2014: impresion con arrobas para Montifrut
        If vParamAplic.Cooperativa = 12 Then
            If Check3.Value = 1 Then
                cadParam = cadParam & "pConArrobas=1|"
            Else
                cadParam = cadParam & "pConArrobas=0|"
            End If
            numParam = numParam + 1
        End If
        
        
        LlamarImprimir
        
        If frmVisReport.EstaImpreso Then
            ActualizarRegistros "rfactsoc", cadSelect
        End If
    End If


End Sub

Private Sub CmdCanApor_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdCancelAntPdtes_Click()
    Unload Me
End Sub

Private Sub CmdCancelAntVC_Click()
    Unload Me
End Sub

Private Sub cmdCancelAnt_Click()
    Unload Me
End Sub

Private Sub CmdCancelGenFac_Click()
    Unload Me
End Sub

Private Sub CmdCancelModelo_Click()
    Unload Me
End Sub

Private Sub cmdCancelReimp_Click()
    Unload Me
End Sub

Private Sub cmdCancelDesF_Click()
    Unload Me
End Sub

Private Sub CmdCancelResul_Click()
    Unload Me
End Sub


Private Sub CmdCanLiqDirecta_Click()
    Unload Me
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Index = 1 Then
        Select Case Combo1(Index).ListIndex
            Case 0 ' anticipo venta campo
                ' si solo hay un tipo de movimiento de anticipo venta campo
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(3) = 1 Then
                    txtcodigo(9).Text = vParamAplic.PrimFactAntVC
                    txtcodigo(10).Text = vParamAplic.UltFactAntVC
                End If
            Case 1 ' liquidacion venta campo
                ' si solo hay un tipo de movimiento de liquidacion venta campo
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(4) = 1 Then
                    txtcodigo(9).Text = vParamAplic.PrimFactLiqVC
                    txtcodigo(10).Text = vParamAplic.UltFactLiqVC
                End If
        End Select
    End If
    
    If Index = 0 Then
        If vParamAplic.Cooperativa = 2 Then
            Check1(10).visible = True
            Check1(10).Enabled = Combo1(Index).ListIndex = 1
            If Combo1(Index).ListIndex = 0 Then Check1(10).Value = 0
        End If
    End If
    
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 1, 2, 3 ' 1-Inf.Anticipos
                         ' 2-Listado de Previsión de pago
                         ' 3-Facturas de Anticipos
                PonerFoco txtcodigo(12)
                
            Case 4    ' reimpresion de facturas de SOCIOS
                PonerFoco txtcodigo(4)
                
            Case 5    ' deshacer proceso de facturacion de anticipos
                PonerFoco txtcodigo(8)
                Me.Pb2.visible = False
                ' si solo hay un tipo de movimiento de anticipo
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(1) = 1 Then
                    txtcodigo(9).Text = vParamAplic.PrimFactAnt
                    txtcodigo(10).Text = vParamAplic.UltFactAnt
                End If
                
            Case 6    ' generacion de factura de venta campo (anticipo o liquidacion)
                Combo1(0).ListIndex = 0 ' por defecto anticipo
                Pb3.visible = False
                txtcodigo(14).Text = Format(Now, "dd/mm/yyyy")
                Check1(0).Value = 1
                Check1(1).Value = 1
                PonerFocoCmb Combo1(0)
                
            Case 16 ' generacion de factura de anticipo venta campo sin entradas asociadas
                txtcodigo(51).Text = Format(Now, "dd/mm/yyyy")
                Check1(8).Value = 1
                
            Case 7    ' deshacer proceso de facturacion de venta campo
                Me.Pb2.visible = False
                Combo1(1).ListIndex = 0 ' por defecto anticipo
'                txtCodigo(9).Text = vParamAplic.PrimFactAntVC
'                txtCodigo(10).Text = vParamAplic.UltFactAntVC
                PonerFoco txtcodigo(8)
            
            Case 8, 9   ' 8 - informe de resultados
                        ' 9 - informe de retenciones
                PonerFoco txtcodigo(24)
                
            Case 10, 11  ' 10 - grabacion modelo 190
                         ' 11 - grabacion modelo 346
                PonerFoco txtcodigo(34)
                
                ' [Monica] 14/01/2010 No hay cabecera 190a
                Me.FrameDomicilio.visible = False '(OpcionListado = 10)
                Me.FrameDomicilio.Enabled = False '(OpcionListado = 10)
                Me.BarraEst.Enabled = False ' (OpcionListado = 10)
                Me.BarraEst.visible = False '(OpcionListado = 10)
                
'                txtcodigo(30).Text = Format(Year(Now), "0000")
                txtcodigo(36).Text = vParam.PerContacto
                txtcodigo(37).Text = vParam.Telefono
            
            Case 12, 13, 14 ' 12-Inf.Liquidacion
                            ' 13-Listado de Previsión de pago
                            ' 14-Facturas de Liquidacion
                PonerFoco txtcodigo(12)
            
            Case 15    ' deshacer proceso de facturacion de liquidacion
                PonerFoco txtcodigo(8)
                Me.Pb2.visible = False
                ' si solo hay un tipo de movimiento de liquidacion
                ' mostramos cual fue la ultima facturacion
                If NroTotalMovimientos(2) = 1 Then
                    txtcodigo(9).Text = vParamAplic.PrimFactLiq
                    txtcodigo(10).Text = vParamAplic.UltFactLiq
                End If
                
            Case 17 ' recalculo de importes de venta campo
                PonerFoco txtcodigo(52)
                
            Case 19 ' liquidacion de entrada directa
                PonerFoco txtcodigo(61)
                txtcodigo(61).Text = Format(Now, "dd/mm/yyyy")
                
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    Set List = New Collection
    For h = 24 To 27
        List.Add h
    Next h
    For h = 1 To 10
        List.Add h
    Next h
    List.Add 12
    List.Add 13
    List.Add 14
    List.Add 15
    List.Add 18
    List.Add 19
    
    For h = 0 To 5
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 12 To 13
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 16 To 21
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 24 To 25
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 28 To 29
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 34 To 35
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 48 To 49
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 52 To 56
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 64 To 67
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    
    
    For h = 0 To imgAyuda.Count - 1
        imgAyuda(h).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next h
    
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    FrameAnticipos.visible = False
    FrameReimpresion.visible = False
    FrameDesFacturacion.visible = False
    FrameGeneraFactura.visible = False
    FrameResultados.visible = False
    FrameGrabacionModelos.visible = False
    FrameGenFactAnticipoVC.visible = False
    FrameRecalculoImporte.visible = False
    Me.FrameAportaciones.visible = False
    FrameLiqDirecta.visible = False
    Me.FrameAnticiposPdtes.visible = False
    
    '[Monica]11/04/2013: check de Descontar facturas varias (por defecto inhibido)
    Check1(14).Enabled = False
    Check1(14).visible = False
    Check1(14).Value = 0
    
    '[Monica]30/05/2013: check de Descontar facturas varias en venta campo(por defecto inhibido)
    Check1(15).Enabled = False
    Check1(15).visible = False
    Check1(15).Value = 0
    
    
    '[Monica]11/03/2015: Observaciones de la factura de liquidacion
    ObsFactura = ""
    
    
    
    
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 1, 12   '1- Informe de Anticipos
                 '12- Informe de Liquidacion
        FrameAnticiposVisible True, h, w
        Tabla = "rhisfruta"
        Me.FrameFechaAnt.visible = True
        Me.FrameFechaAnt.Enabled = True
        Me.FrameAgrupado.visible = False
        Me.FrameAgrupado.Enabled = False
        
        If OpcionListado = 1 Then
            Me.Label3.Caption = "Informe de Anticipos"
            Me.Label2(25).Caption = "Fecha Anticipo"
            '++Monica:03/12/2009
            Check1(5).Enabled = False
            Check1(5).visible = False
            Check1(5).Value = 0
        Else
            Me.Label3.Caption = "Informe de Liquidación"
            Me.Label2(25).Caption = "Fecha Liquidación"
            '++Monica:03/12/2009
            Check1(5).Enabled = True
            Check1(5).visible = True
            Check1(5).Value = 0
            
            
'[Monica]27/01/2016: cambiamos esto liquidaciones y complementaria
            '++Monica:03/06/2013: distinguimos para Catadau entre entradas
            Check1(16).Enabled = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            Check1(16).visible = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            imgAyuda(2).visible = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            imgAyuda(2).Enabled = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            If Check1(16).Enabled Then
                Check1(16).Top = 3690
                imgAyuda(2).Top = 3690
            End If
            
            FrameTipo.Enabled = (vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            FrameTipo.visible = (vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
            FrameTipo.Top = 4530
            
            Check1(25).visible = (Check1(5).Value = 1)
            Check1(25).Enabled = (Check1(5).Value = 1)
            Check1(26).visible = (Check1(5).Value = 1)
            Check1(26).Enabled = (Check1(5).Value = 1)
            If Check1(25).Enabled Then
                Check1(25).Value = 1
                Check1(26).Value = 1
            Else
                Check1(25).Value = 0
                Check1(26).Value = 0
            End If
            

            '[Monica]10/03/2014: no permitimos facturas negativas colo para alzira
            If vParamAplic.Cooperativa = 4 Then
                Check1(21).visible = True
                Check1(21).Enabled = True
            End If
        End If
        Me.Pb1.visible = False
        Me.Label2(10).Caption = ""
        Me.Label2(12).Caption = ""
        
        Me.FrameOpciones.visible = False
        Me.FrameOpciones.Enabled = False
        
        CargaCombo
        Combo1(2).ListIndex = 2
    Case 2, 13   '2 - Listado de prevision de pagos de anticipos
                 '13- Listado de prevision de pagos de liquidacion
        FrameAnticiposVisible True, h, w
        Tabla = "rhisfruta"
        Me.FrameFechaAnt.visible = False
        Me.FrameFechaAnt.Enabled = False
        Me.FrameAgrupado.visible = True And Not AnticipoGastos And Not AnticipoGenerico
        Me.FrameAgrupado.Enabled = True And Not AnticipoGastos And Not AnticipoGenerico
        
        '[Monica]11/04/2013: activamos el check de descontar facturas varias
        Check1(14).visible = (vParamAplic.HayFacVarias And Not AnticipoGastos And Not AnticipoGenerico)
        Check1(14).Enabled = (vParamAplic.HayFacVarias And Not AnticipoGastos And Not AnticipoGenerico)
        
        If OpcionListado = 2 Then
            Me.Label3.Caption = "Previsión de Pagos Anticipos"
            If AnticipoGastos Then
                Me.Label3.Caption = "Previsión Pagos Anticipos Gastos"
            End If
            If AnticipoGenerico Then
                Me.Label3.Caption = "Previsión Anticipos Genérico/Retirada"
                Check1(12).visible = True
                Check1(12).Enabled = True
                '[Monica]23/12/2014:VR
                Check1(22).visible = True
                Check1(22).Enabled = False
            
                
                imgAyuda(1).Enabled = True
                imgAyuda(1).visible = True
                Label2(43).visible = True
                txtcodigo(59).visible = True
            End If
            '++Monica:03/12/2009
            Check1(5).Enabled = False
            Check1(5).visible = False
            Check1(5).Value = 0
        Else
            Me.Label3.Caption = "Previsión de Pagos Liquidación"
            If LiquidacionIndustria Then
                Me.Label3.Caption = "Previsión de Pagos Liquidación Industria"
                Check1(5).Enabled = False
                Check1(5).visible = False
            Else
                '++Monica:03/12/2009
                Check1(5).Enabled = True
                Check1(5).visible = True
                Check1(5).Value = 0
            
                
                
'[Monica]27/01/2016: cambiamos esto liquidaciones y complementaria
                '++Monica:03/06/2013: distinguimos para Catadau entre entradas
                Check1(16).Enabled = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria) And Check1(5).Value = 0
                Check1(16).visible = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria) And Check1(5).Value = 0
                imgAyuda(2).visible = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria) And Check1(5).Value = 0
                imgAyuda(2).Enabled = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria) And Check1(5).Value = 0
                If Check1(16).Enabled Then
                    Check1(16).Top = 3690
                    imgAyuda(2).Top = 3690
                End If
                
                FrameTipo.Enabled = (vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
                FrameTipo.visible = (vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
                FrameTipo.Top = 4530
                
                Check1(25).visible = (Check1(5).Value = 1)
                Check1(25).Enabled = (Check1(5).Value = 1)
                Check1(26).visible = (Check1(5).Value = 1)
                Check1(26).Enabled = (Check1(5).Value = 1)
                If Check1(25).Enabled Then
                    Check1(25).Value = 1
                    Check1(26).Value = 1
                Else
                    Check1(25).Value = 0
                    Check1(26).Value = 0
                End If
            
            
                '[Monica]10/03/2014: no permitimos facturas negativas colo para alzira
                If vParamAplic.Cooperativa = 4 Then
                    Check1(21).visible = True
                    Check1(21).Enabled = True
                End If
            
            
            End If
        End If
        
        Me.Pb1.visible = False
        Me.Label2(10).Caption = ""
        Me.Label2(12).Caption = ""
        Me.FrameOpciones.visible = False
        Me.FrameOpciones.Enabled = False
        
        CargaCombo
        Combo1(2).ListIndex = 2
        Combo1(3).ListIndex = 0
        
        '[Monica]22/03/2012: Añadimos si calculamos sobre el precio comision o no (sólo para Picassent)
        Check1(13).Enabled = (vParamAplic.Cooperativa = 2) And (OpcionListado = 13)
        Check1(13).visible = (vParamAplic.Cooperativa = 2) And (OpcionListado = 13)
        If Check1(13).Enabled Then Check1(13).Left = 3390
        
        If LiquidacionIndustria Then 'ocultamos agrupado y recolectado
            FrameRecolectado.visible = False
            FrameRecolectado.Enabled = False
            FrameAgrupado.visible = False
            FrameAgrupado.Enabled = False
        End If
        
    Case 3, 14   '3 - Factura de Anticipos
                 '14- Factura de Liquidacion
        FrameAnticiposVisible True, h, w
        Tabla = "rhisfruta"
        Me.FrameFechaAnt.visible = True
        Me.FrameFechaAnt.Enabled = True
        Me.FrameAgrupado.visible = False
        Me.FrameAgrupado.Enabled = False
        Me.Caption = "Facturación"
        
        '[Monica]11/03/2015: observaciones de la factura visibles
        If OpcionListado = 14 Then
            Me.Label2(49).visible = (vParamAplic.Cooperativa = 4)
            Me.txtcodigo(68).visible = (vParamAplic.Cooperativa = 4)
            Me.txtcodigo(68).Enabled = (vParamAplic.Cooperativa = 4)
        End If
        
        '[Monica]11/04/2013: activamos el check de descontar facturas varias
        Check1(14).visible = (vParamAplic.HayFacVarias And Not AnticipoGastos And Not AnticipoGenerico)
        Check1(14).Enabled = (vParamAplic.HayFacVarias And Not AnticipoGastos And Not AnticipoGenerico)
        
        If OpcionListado = 3 Then
            Me.Label3.Caption = "Factura de Anticipos"
            Me.Label2(25).Caption = "Fecha Anticipo"
            If AnticipoGastos Then
                Me.Label3.Caption = "Factura de Anticipos Gastos"
            End If
            If AnticipoGenerico Then
                Me.Label3.Caption = "Factura Anticipos Genérico/Retirada"
                Check1(12).visible = True
                Check1(12).Enabled = True
                '[Monica]23/12/2014:VR
                Check1(22).visible = True
                Check1(22).Enabled = False
                
                imgAyuda(1).Enabled = True
                imgAyuda(1).visible = True
                Label2(43).visible = True
                txtcodigo(59).visible = True
            End If
            
            '++Monica:03/12/2009
            Check1(5).Enabled = False
            Check1(5).visible = False
            Check1(5).Value = 0
        Else
            Me.Label3.Caption = "Factura de Liquidación"
            Me.Label2(25).Caption = "Fecha Liquidación"
            If LiquidacionIndustria Then
                Me.Label3.Caption = "Factura de Liquidación Industria"
                Check1(5).Enabled = False
                Check1(5).visible = False
                Check1(5).Value = 0
            Else
                '++Monica:03/12/2009
                Check1(5).Enabled = True
                Check1(5).visible = True
                Check1(5).Value = 0
                
'[Monica]27/01/2016: cambiamos esto liquidaciones y complementaria
                '++Monica:03/06/2013: distinguimos para Catadau entre entradas
                Check1(16).Enabled = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
                Check1(16).visible = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
                imgAyuda(2).visible = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
                imgAyuda(2).Enabled = False '(vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
                If Check1(16).Enabled Then
                    Check1(16).Top = 3690
                    imgAyuda(2).Top = 3690
                End If
                
                FrameTipo.Enabled = (vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
                FrameTipo.visible = (vParamAplic.Cooperativa = 0 And Not LiquidacionIndustria)
                FrameTipo.Top = 4530
                
                Check1(25).visible = (Check1(5).Value = 1)
                Check1(25).Enabled = (Check1(5).Value = 1)
                Check1(26).visible = (Check1(5).Value = 1)
                Check1(26).Enabled = (Check1(5).Value = 1)
                If Check1(25).Enabled Then
                    Check1(25).Value = 1
                    Check1(26).Value = 1
                Else
                    Check1(25).Value = 0
                    Check1(26).Value = 0
                End If
            
                '[Monica]10/03/2014: no permitimos facturas negativas colo para alzira
                If vParamAplic.Cooperativa = 4 Then
                    Check1(21).visible = True
                    Check1(21).Enabled = True
                End If
            
            End If
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
        Me.Check1(5).Value = 0
        
        CargaCombo
        Combo1(2).ListIndex = 2
        
    Case 4   ' Reimpresion de facturas de SOCIOS
        FrameReimpresionVisible True, h, w
        Tabla = "rfactsoc"
        Me.FrameFechaAnt.visible = True
        Me.FrameFechaAnt.Enabled = True
        Me.Label3.Caption = "Factura de Socios"
        CargarListView (0)
        
        '[Monica]28/01/2014: Impresion con Arrobas Montifrut
        Check3.visible = (vParamAplic.Cooperativa = 12)
        Check3.Enabled = (vParamAplic.Cooperativa = 12)
        
        
    Case 5   ' Deshacer Proceso de facturación de Anticipos
        ActivarCLAVE
        FrameTipoFactura.visible = False
        FrameDesFacturacionVisible True, h, w
        Tabla = "rfactsoc"
        Me.Caption = "Deshacer Proceso Facturación de Anticipos"
        
    Case 6   ' Generacion de factura de venta campo (anticipo o liquidacion)
        
        '[Monica]30/05/2013: activamos el check de descontar facturas varias en venta campo
        Check1(15).visible = (vParamAplic.HayFacVarias)
        Check1(15).Enabled = (vParamAplic.HayFacVarias)
    
        FrameGeneraFacturaVisible True, h, w
        CargaCombo
        Tabla = "rhisfruta"
        Me.Caption = "Facturación"
    
        If vParamAplic.Cooperativa = 2 Then
            Check1(10).visible = True
            Check1(10).Enabled = (Combo1(0).ListIndex = 1)
            Check1(10).Value = 0
        End If
    
    Case 16   ' Generacion de factura de anticipo venta campo sin entradas asociadas
        FrameGenFactAnticipoVCVisible True, h, w
    
    Case 161 '  Generacion de factura de anticipo sin entradas
        FrameGenFactAnticipoSinEntVisible True, h, w
        
    
    Case 17   ' Recalculo de importes de venta campo
        FrameRecalculoImporteVisible True, h, w
    
    Case 7   ' Deshacer Proceso de facturación de venta campo
        ActivarCLAVE
        FrameTipoFactura.visible = True
        CargaCombo
        FrameDesFacturacionVisible True, h, w
        Tabla = "rfactsoc"
        Me.Caption = "Deshacer Proceso Facturación Venta Campo"
                
    Case 8, 9   '8= Informe de Resultados de facturas de SOCIOS
                '9= Informe de Retenciones de facturas de SOCIOS
        If OpcionListado = 8 Then
            Label8.Caption = "Listado de Resultados"
            txtcodigo(26).Text = Format(vParam.FecIniCam, "dd/mm/yyyy")
            txtcodigo(27).Text = Format(vParam.FecFinCam, "dd/mm/yyyy")
        Else
            Label8.Caption = "Listado de Retenciones/Aportaciones"
            txtcodigo(26).Text = Format(DateAdd("yyyy", -1, vParam.FecIniCam), "dd/mm/yyyy")
            txtcodigo(27).Text = Format(vParam.FecFinCam, "dd/mm/yyyy")
        End If
        
        FrameOpc.Enabled = (OpcionListado = 9)
        FrameOpc.visible = (OpcionListado = 9)
        
        txtcodigo(32).Text = Format(Now, "dd/mm/yyyy")
        
        FrameFechaCertif.visible = False
        FrameFechaCertif.Enabled = False
        
        FrameResultadosVisible True, h, w
        Tabla = "rfactsoc"
        CargarListView (1)
        
    Case 10, 11 '10 = grabacion modelo 190
                '11 = grabacion modelo 346
        If OpcionListado = 10 Then
            Label9.Caption = "Grabación Modelo 190"
        Else
            Label9.Caption = "Grabación Modelo 346"
        End If
        FrameGrabacionModelosVisible True, h, w
        Tabla = "rfactsoc"
    
    Case 15   ' Deshacer Proceso de facturación de Liquidacion
        ActivarCLAVE
        FrameTipoFactura.visible = False
        FrameDesFacturacionVisible True, h, w
        Tabla = "rfactsoc"
        Me.Caption = "Deshacer Proceso Facturación de Liquidación"
        
    Case 18   ' Calculo de Aportaciones (SOLO PICASSENT)
        CargaCombo
        FrameAportacionVisible True, h, w
        Tabla = "rhisfruta"
        
    Case 19   ' Liquidacion Directa para Alzira
        FrameLiqDirectaVisible True, h, w
        Tabla = "rhisfruta"
    
    
    Case 20   ' Anticipos pendientes de descontar
        FrameAnticiposPdteDescontarVisible True, h, w
        Tabla = "rhisfruta"
    
    
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case OpcionListado
        Case 3
            DesBloqueoManual ("FACANT")
        Case 14
            DesBloqueoManual ("FACLIQ")
    End Select
    
'    LiqComplementariaUnica = False
End Sub




Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de calidades
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
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
    If Not AnyadirAFormula(cadSelect1, SQL) Then Exit Sub
    If Not AnyadirAFormula(cadSelect2, SQL) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

    Variedades = SQL


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
    PonerDatosCampo txtcodigo(50)
End Sub


Private Sub frmMens2_DatoSeleccionado(CadenaSeleccion As String)
    Albaranes = CadenaSeleccion
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
    If Not AnyadirAFormula(cadSelect, SQL) Then Exit Sub


End Sub

Private Sub frmMens4_DatoSeleccionado(CadenaSeleccion As String)

    vReturn = 2
    If CadenaSeleccion <> "" Then vReturn = CInt(CadenaSeleccion)

End Sub

Private Sub frmMens5_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        vFechas = CadenaSeleccion
    Else
        vFechas = ""
    End If
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

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si está marcado se liquidan los socios que sean terceros de módulos" & vbCrLf & _
                      "con los precios del socio. En caso contrario, sólo se liquidan los" & vbCrLf & _
                      "socios que no sean terceros." & vbCrLf & vbCrLf & _
                      "Los socios terceros entidad se tratan como tales en la recepcion " & _
                      "de facturas de socios terceros" & vbCrLf & vbCrLf
        Case 1
           ' "____________________________________________________________"
            vCadena = "Si está marcado se hará un anticipo de Retirada y se marcará como " & vbCrLf & _
                      "tal en el mantenimiento de las Facturas de Socio. Utiliza el precio" & vbCrLf & _
                      "de anticipo de Retirada. " & vbCrLf & vbCrLf & _
                      "En caso contrario, se generará un anticipo Genérico que utiliza el" & vbCrLf & _
                      "precio de anticipo Genérico del mantenimiento de precios." & vbCrLf & vbCrLf & _
                      "Ambos calculan sobre el total de kilos sin tener en cuenta calidades" & vbCrLf & vbCrLf & _
                      "Sólo se descontarán en la Factura de Liquidación los anticipos " & vbCrLf & _
                      "Genéricos los de retirada aparecen descontados en el informe de " & vbCrLf & _
                      "Liquidación, por tanto en el resultado de la Factura de Liquidación." & vbCrLf & vbCrLf
                      
        Case 2
           ' "____________________________________________________________"
            vCadena = "Si está marcado se liquidan las entradas que sean Normales (Piedra). " & vbCrLf & _
                      "En caso contrario, sólo se liquidan las entradas de Producto" & vbCrLf & _
                      "Integrado." & vbCrLf & vbCrLf
                      
                      
        Case 3
           ' "____________________________________________________________"
            vCadena = "Este informe sólo saldrá correctamente, si las variedades de las facturas" & vbCrLf & _
                      "seleccionadas son del mismo grupo de retención." & vbCrLf & vbCrLf
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 18, 19, 20, 21, 28, 29, 53, 54, 66, 67 'Clases
            AbrirFrmClase (Index)
        
        
        Case 0, 1, 12, 13, 16, 17, 24, 25, 49, 52, 55, 56, 64, 65 'SOCIOS
            AbrirFrmSocios (Index)
        
        
        Case 2, 3, 4, 5 ' TRANSPORTISTAS
            AbrirFrmTransportistas (Index)
        
        Case 48 ' VARIEDAD
            AbrirFrmVariedad (Index)
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

    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    Select Case Index
        Case 0
            indice = 6
        Case 1
            indice = 7
        Case 2
            indice = 15
        Case 3, 4
            indice = Index - 1
        Case 5
            indice = 14
        Case 6
            indice = 11
        Case 7, 8
            indice = Index + 15
        Case 9, 10
            indice = Index + 17
        Case 11, 12
            indice = Index + 21
        Case 13
            indice = 32
        Case 14
            indice = 61
        Case 16
            indice = 51
        Case 17
            indice = 57
        Case 15
            indice = 58
        Case 18
            indice = 62
        Case 19
            indice = 63
    End Select

    imgFec(0).Tag = indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(indice).Text <> "" Then frmC.NovaData = txtcodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(CByte(imgFec(0).Tag)) '<===
    ' ********************************************

End Sub







Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
    If OpcionListado = 10 Then
        If Index = 40 Then
            BarraEst.SimpleText = " CL = Calle    AV = Avenida."
        Else
            BarraEst.SimpleText = ""
        End If
        BarraEst.visible = (BarraEst.SimpleText <> "")
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
            Case 2: KEYBusqueda KeyAscii, 41 'transportista desde
            Case 3: KEYBusqueda KeyAscii, 42 'transportista hasta
            Case 4: KEYBusqueda KeyAscii, 43 'transportista desde
            Case 5: KEYBusqueda KeyAscii, 44 'transportista hasta
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
            Case 48: KEYBusqueda KeyAscii, 48 'variedad
            Case 49: KEYBusqueda KeyAscii, 49 'socio
            Case 52: KEYBusqueda KeyAscii, 52 'socio
            
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
            Case 51: KEYFecha KeyAscii, 16 'fecha
            Case 61: KEYFecha KeyAscii, 14 'fecha de liquidacion directa
        
            Case 64: KEYBusqueda KeyAscii, 64 'socio desde
            Case 65: KEYBusqueda KeyAscii, 65 'socio hasta
            Case 66: KEYBusqueda KeyAscii, 66 'clase desde
            Case 67: KEYBusqueda KeyAscii, 67 'clase hasta
            Case 62: KEYFecha KeyAscii, 18 'fecha
            Case 63: KEYFecha KeyAscii, 19 'fecha
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
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim b As Boolean

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
    
        Case 0, 1, 12, 13, 16, 17, 24, 25, 34, 35, 49, 52, 55, 56, 64, 65 'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            
            If Index = 49 Then PonerCamposSocio
            
        Case 41, 42, 43, 44 ' TRANSPORTISTAS
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rtransporte", "nomtrans", "codtrans", "T")
'            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            
        Case 4, 5 ' NROS DE FACTURA
            PonerFormatoEntero txtcodigo(Index)
            
        Case 2, 3, 6, 7, 11, 15, 26, 27, 32, 51, 57, 5, 62, 63 'FECHAS
            b = True
            If txtcodigo(Index).Text <> "" Then b = PonerFormatoFecha(txtcodigo(Index))
            
            '[Monica]11/03/2015: si es factura vamos a las observaciones
            If b And Index = 15 And (OpcionListado = 3 Or OpcionListado = 14) Then ' And vParamAplic.Cooperativa = 4 Then
                PonerFoco txtcodigo(68)
                Exit Sub
            End If
            
            
            If b And Index = 7 And (Me.OpcionListado = 1 Or Me.OpcionListado = 3 Or Me.OpcionListado = 12 Or Me.OpcionListado = 14) Then PonerFoco txtcodigo(15)
            If b And Index = 15 And (Me.OpcionListado = 1 Or Me.OpcionListado = 3 Or Me.OpcionListado = 12 Or Me.OpcionListado = 14) Then
                If AnticipoGastos Then
                    cmdAceptarAntGastos.SetFocus
                Else
                    If AnticipoGenerico Then
                        cmdAceptarAntGene.SetFocus
                    Else
                        If LiquidacionIndustria Then
                            Me.cmdAceptarLiqIndustria.SetFocus
                        Else
                            cmdAceptarAnt.SetFocus
                        End If
                    End If
                End If
            End If
            
            
        Case 68 ' observaciones de la factura
            If (OpcionListado = 3 Or OpcionListado = 14) Then
                If AnticipoGastos Then
                    cmdAceptarAntGastos.SetFocus
                Else
                    If AnticipoGenerico Then
                        cmdAceptarAntGene.SetFocus
                    Else
                        If LiquidacionIndustria Then
                            Me.cmdAceptarLiqIndustria.SetFocus
                        Else
                            cmdAceptarAnt.SetFocus
                        End If
                    End If
                End If
            End If
            
        Case 14, 22, 23, 61 ' FECHA DE GENERACION DE FACTURA
            '[Monica]28/08/2013: no miramos si la fecha esta dentro de campaña
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index), True
            
        Case 8 ' password de deshacer facturacion
            If txtcodigo(Index).Text = "" Then Exit Sub
            If Trim(txtcodigo(Index).Text) <> Trim(txtcodigo(Index).Tag) Then
                MsgBox "    ACCESO DENEGADO    ", vbExclamation
                txtcodigo(Index).Text = ""
                PonerFoco txtcodigo(Index)
            Else
                DesactivarCLAVE
                Select Case OpcionListado
                    Case 5, 15 '5 = anticipos
                               '15= liquidaciones
                        PonerFoco txtcodigo(9)
                    Case 7 ' venta campo
                        PonerFocoCmb Combo1(1)
                End Select
            End If
        
        Case 9, 10 ' numero de facturas
            If txtcodigo(Index).Text <> "" Then PonerFormatoEntero txtcodigo(Index)
        
        Case 30, 31, 37, 39 ' datos de modelo190 y modelo346
            If txtcodigo(Index).Text <> "" Then PonerFormatoEntero txtcodigo(Index)
            
        Case 33 ' nro de justificante en el certificado de retenciones
            If PonerFormatoEntero(txtcodigo(Index)) Then
                CmdAcepResul.SetFocus
            End If
            
        Case 18, 19, 20, 21, 28, 29, 53, 54, 66, 67 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
            
        Case 45, 46 ' importe
            PonerFormatoDecimal txtcodigo(Index), 3
        
        Case 50 ' campo
            PonerDatosCampo txtcodigo(Index).Text
        
        Case 48 ' variedad
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
    
        Case 47 ' Importe de aportacion a repartir
            PonerFormatoDecimal txtcodigo(Index), 3
                
        Case 59 ' Kilos de Retirada
            PonerFormatoEntero txtcodigo(59)
    
        Case 60 ' precio calidad en liquidacion directa
            PonerFormatoDecimal txtcodigo(60), 7
    
    
    End Select
End Sub


Private Sub FrameAnticiposVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
Dim b As Boolean

'Frame para el listado de socios por seccion
    Me.FrameAnticipos.visible = visible
    If visible = True Then
    
        Me.FrameAnticipos.Top = -90
        Me.FrameAnticipos.Left = 0
        Me.FrameAnticipos.Height = 6630 '5970 '5640
        Me.FrameAnticipos.Width = 6615
        w = Me.FrameAnticipos.Width
        h = Me.FrameAnticipos.Height
        
        b = (OpcionListado = 1 Or OpcionListado = 2 Or OpcionListado = 3 Or _
             OpcionListado = 12 Or OpcionListado = 13 Or OpcionListado = 14) And _
             Not AnticipoGastos And Not LiquidacionIndustria And Not AnticipoGenerico
             
        
        FrameRecolectado.Enabled = b
        FrameRecolectado.visible = b
    
        '[Monica]24/06/2011: si el socio es Alzira puede seleccionar si liquidar socios terceros de modulos o no terceros
                                                            '[Monica11/10/2013: Picassent pasa a tener terceros
        Check1(11).Enabled = (vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 2) And b
        Check1(11).visible = (vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 2) And b
        imgAyuda(0).Enabled = (vParamAplic.Cooperativa = 4) And b
        imgAyuda(0).visible = (vParamAplic.Cooperativa = 4) And b
        
        
        '[Monica]11/10/2013: colocamos el check de terceros mas a la izquierda
        If vParamAplic.Cooperativa = 2 And (OpcionListado = 1 Or OpcionListado = 2 Or OpcionListado = 3) Then
            Check1(11).Left = 3390
        End If
   
        If AnticipoGastos Then
            ' desactivo los botones de anticipos normales
            Me.cmdAceptarAnt.visible = False
            Me.cmdAceptarAnt.Enabled = False
            ' activo los botones de anticipos de gastos
            Me.cmdAceptarAntGastos.visible = True
            Me.cmdAceptarAntGastos.Enabled = True
            ' desactivo los botones de liquidacion industria
            Me.cmdAceptarLiqIndustria.visible = False
            Me.cmdAceptarLiqIndustria.Enabled = False
            ' desactivo los botones de anticipos generico
            Me.cmdAceptarAntGene.visible = False
            Me.cmdAceptarAntGene.Enabled = False
            
            ' los situo
            Me.cmdAceptarAntGastos.Left = 4110
            Me.cmdAceptarAntGastos.Caption = "&Aceptar"
        End If
        If AnticipoGenerico Then
            ' desactivo los botones de anticipos normales
            Me.cmdAceptarAnt.visible = False
            Me.cmdAceptarAnt.Enabled = False
            ' desactivo los botones de anticipos de gastos
            Me.cmdAceptarAntGastos.visible = False
            Me.cmdAceptarAntGastos.Enabled = False
            ' desactivo los botones de liquidacion industria
            Me.cmdAceptarLiqIndustria.visible = False
            Me.cmdAceptarLiqIndustria.Enabled = False
            ' activo los botones de anticipos generico
            Me.cmdAceptarAntGene.visible = True
            Me.cmdAceptarAntGene.Enabled = True
            
            ' los situo
            Me.cmdAceptarAntGene.Left = 4110
            Me.cmdAceptarAntGene.Caption = "&Aceptar"
        End If
        If LiquidacionIndustria Then
            ' desactivo los botones de anticipos normales
            Me.cmdAceptarAnt.visible = False
            Me.cmdAceptarAnt.Enabled = False
            ' desactivo los botones de anticipos de gastos
            Me.cmdAceptarAntGastos.visible = False
            Me.cmdAceptarAntGastos.Enabled = False
            ' activo los botones de liquidacion industria
            Me.cmdAceptarLiqIndustria.visible = True
            Me.cmdAceptarLiqIndustria.Enabled = True
            ' desactivo los botones de anticipos generico
            Me.cmdAceptarAntGene.visible = False
            Me.cmdAceptarAntGene.Enabled = False
            
            ' los situo
            Me.cmdAceptarLiqIndustria.Left = 4110
            Me.cmdAceptarLiqIndustria.Caption = "&Aceptar"
        End If
    
    End If
End Sub


Private Sub FrameReimpresionVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameReimpresion.visible = visible
    If visible = True Then
        Me.FrameReimpresion.Top = -90
        Me.FrameReimpresion.Left = 0
        Me.FrameReimpresion.Height = 5640
        Me.FrameReimpresion.Width = 6675
        w = Me.FrameReimpresion.Width
        h = Me.FrameReimpresion.Height
    End If
End Sub


Private Sub FrameResultadosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameResultados.visible = visible
    If visible = True Then
        Me.FrameResultados.Top = -90
        Me.FrameResultados.Left = 0
        Me.FrameResultados.Height = 7320
        Me.FrameResultados.Width = 6675
        w = Me.FrameResultados.Width
        h = Me.FrameResultados.Height
    End If
End Sub

Private Sub FrameGrabacionModelosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameGrabacionModelos.visible = visible
    If visible = True Then
        Me.FrameGrabacionModelos.Top = -90
        Me.FrameGrabacionModelos.Left = 0
        Select Case OpcionListado
            Case 10, 11
                Me.FrameGrabacionModelos.Height = 5490
                Me.CmdAcepModelo.Top = 4740
                Me.CmdCancelModelo.Top = 4740
        End Select
        Me.FrameGrabacionModelos.Width = 6675
        w = Me.FrameGrabacionModelos.Width
        h = Me.FrameGrabacionModelos.Height
    End If
End Sub


Private Sub FrameAportacionVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameAportaciones.visible = visible
    If visible = True Then
        Me.FrameAportaciones.Top = -90
        Me.FrameAportaciones.Left = 0
        Me.FrameAportaciones.Height = 6930
        Me.FrameAportaciones.Width = 6615
        w = Me.FrameAportaciones.Width
        h = Me.FrameAportaciones.Height
    End If
End Sub




Private Sub FrameDesFacturacionVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameDesFacturacion.visible = visible
    If visible = True Then
        Me.FrameDesFacturacion.Top = -90
        Me.FrameDesFacturacion.Left = 0
        Me.FrameDesFacturacion.Height = 4740
        Me.FrameDesFacturacion.Width = 6615
        w = Me.FrameDesFacturacion.Width
        h = Me.FrameDesFacturacion.Height
    End If
End Sub

Private Sub FrameGeneraFacturaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameGeneraFactura.visible = visible
    If visible = True Then
        Me.FrameGeneraFactura.Top = -90
        Me.FrameGeneraFactura.Left = 0
        Me.FrameGeneraFactura.Height = 5790
        Me.FrameGeneraFactura.Width = 6615
        w = Me.FrameGeneraFactura.Width
        h = Me.FrameGeneraFactura.Height
    End If
End Sub

'[Monica]06/09/2013: generacion de facturas de anticipo sin entradas
Private Sub FrameGenFactAnticipoSinEntVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameGenFactAnticipoVC.visible = visible
    If visible = True Then
    
        Label10.Caption = "Generación Factura Anticipo"
        Label2(14).Caption = "sin entrada en campo asociada"
    
        Me.FrameGenFactAnticipoVC.Top = -90
        Me.FrameGenFactAnticipoVC.Left = 0
        Me.FrameGenFactAnticipoVC.Height = 6270
        Me.FrameGenFactAnticipoVC.Width = 6675
        w = Me.FrameGenFactAnticipoVC.Width
        h = Me.FrameGenFactAnticipoVC.Height
        
        '[Monica]06/11/2013: Modificacion para Picassent, debemos poder crear un anticipo a cuenta de terceros
        Check1(17).visible = (vParamAplic.Cooperativa = 2)
        Check1(17).Enabled = (vParamAplic.Cooperativa = 2)
        
    End If
End Sub




Private Sub FrameGenFactAnticipoVCVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameGenFactAnticipoVC.visible = visible
    If visible = True Then
        Me.FrameGenFactAnticipoVC.Top = -90
        Me.FrameGenFactAnticipoVC.Left = 0
        Me.FrameGenFactAnticipoVC.Height = 6270
        Me.FrameGenFactAnticipoVC.Width = 6675
        w = Me.FrameGenFactAnticipoVC.Width
        h = Me.FrameGenFactAnticipoVC.Height
    End If
End Sub


Private Sub FrameRecalculoImporteVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameRecalculoImporte.visible = visible
    If visible = True Then
        Me.FrameRecalculoImporte.Top = -90
        Me.FrameRecalculoImporte.Left = 0
        Me.FrameRecalculoImporte.Height = 3750
        Me.FrameRecalculoImporte.Width = 6675
        w = Me.FrameRecalculoImporte.Width
        h = Me.FrameRecalculoImporte.Height
    End If
End Sub

Private Sub FrameLiqDirectaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameLiqDirecta.visible = visible
    If visible = True Then
        Me.FrameLiqDirecta.Top = -90
        Me.FrameLiqDirecta.Left = 0
        Me.FrameLiqDirecta.Height = 4200
        Me.FrameLiqDirecta.Width = 6615
        w = Me.FrameLiqDirecta.Width
        h = Me.FrameLiqDirecta.Height
    End If
End Sub

Private Sub FrameAnticiposPdteDescontarVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameAnticiposPdtes.visible = visible
    If visible = True Then
        Me.FrameAnticiposPdtes.Top = -90
        Me.FrameAnticiposPdtes.Left = 0
        Me.FrameAnticiposPdtes.Height = 5430
        Me.FrameAnticiposPdtes.Width = 6615
        w = Me.FrameAnticiposPdtes.Width
        h = Me.FrameAnticiposPdtes.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadSelect1 = ""
    cadSelect2 = ""
    cadSelect3 = ""
    cadParam = ""
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
        .Opcion = OpcionListado
        .ConSubInforme = ConSubInforme
        .Show vbModal
    End With
End Sub


Private Sub AbrirFrmCalidad(indice As Integer)
    indCodigo = indice
    Set frmCal = New frmManCalidades
    frmCal.DatosADevolverBusqueda = "2|3|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmCal.Show vbModal
    Set frmCal = Nothing
End Sub

Private Sub AbrirFrmSeccion(indice As Integer)
    indCodigo = indice
    Set frmSec = New frmManSeccion
    frmSec.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmSec.Show vbModal
    Set frmSec = Nothing
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


Private Sub AbrirFrmTransportistas(indice As Integer)
    indCodigo = indice + 39
    Set frmTra = New frmManTranspor
    frmTra.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmTra.Show vbModal
    Set frmTra = Nothing
End Sub


Private Sub AbrirFrmSituacion(indice As Integer)
    indCodigo = indice
    Set frmSit = New frmManSituCamp
    frmSit.DatosADevolverBusqueda = "0|1|"
    frmSit.Show vbModal
    Set frmSit = Nothing
End Sub


Private Sub AbrirFrmClase(indice As Integer)
    indCodigo = indice
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtcodigo(indice).Text
    
    Set frmCla = Nothing
End Sub



Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
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


Private Function DatosOk() As Boolean
Dim b As Boolean
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

    b = True
    Select Case OpcionListado
        Case 1, 3
            '1 - Informe de Anticipos
            '3 - Factura de Anticipos
            If b Then
                If txtcodigo(6).Text = "" Or txtcodigo(7) = "" Then
                    MsgBox "Debe introducir obligatoriamente el rango de fechas.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(6)
                End If
            End If
            If b Then
                If txtcodigo(15).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de Anticipo.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(15)
                End If
            End If
            
       Case 2 'Prevision de pagos
            If b Then
                If txtcodigo(6).Text = "" Or txtcodigo(7) = "" Then
                    MsgBox "Para realizar la Previsión de Pago de Anticipos debe introducir obligatoriamente el rango de fechas.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(6)
                End If
            End If
       
       Case 5 'Deshacer proceso de facturacion de anticipos
            If txtcodigo(9).Text = "" Or txtcodigo(10).Text = "" Then
                MsgBox "Debe introducir la primera y última factura de la Facturación de Anticipos", vbExclamation
                b = False
                PonerFoco txtcodigo(9)
'            Else
'                ' si la factura hasta no coincide con el contador de stipom no seguir
'                Set vCont = New CTiposMov
'                If vCont.leer("FAA") Then
'                    If vCont.Contador <> CLng(txtCodigo(10).Text) Then
'                        MsgBox "La Factura hasta no es el último número de Factura de Anticipos. Revise.", vbExclamation
'                        b = False
'                    End If
'                End If
'                Set vCont = Nothing
            End If
            
            If b Then
                If txtcodigo(11).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Anticipo.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(11)
                End If
            End If
    
        Case 6 ' factura de ventas campo (anticipo o liquidacion)
            If b Then
                If txtcodigo(14).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de Factura.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(14)
                End If
            End If
        
        Case 16, 161 ' 16 factura de anticipo de venta campo sin entradas asociadas
                     ' 161 factura de anticipo normal
            If b Then
                If txtcodigo(51).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de Factura.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(51)
                End If
            End If
            
            If b Then
                If txtcodigo(49).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente un socio para la Factura.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(49)
                Else
                    ' el socio ha de ser obligatoriamente de la seccion de horto
                    SQL = "select count(*) from rsocios_seccion where codsocio = " & DBSet(txtcodigo(49).Text, "N")
                    SQL = SQL & " and codsecci = " & DBSet(vParamAplic.Seccionhorto, "N")
                    
                    If TotalRegistros(SQL) = 0 Then
                        MsgBox "El socio ha de ser obligatoriamente de la sección de Horto.", vbExclamation
                        b = False
                    End If
                
                End If
            End If
            
            If b Then
                If txtcodigo(45).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente el importe de la Factura.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(45)
                End If
            End If
            
            If b Then
                If txtcodigo(50).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente el campo del socio.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(50)
                Else
                    SQL = "select count(*) from rcampos where codcampo = " & DBSet(txtcodigo(50).Text, "N")
                    SQL = SQL & " and codsocio = " & DBSet(txtcodigo(49).Text, "N")
                    If TotalRegistros(SQL) = 0 Then
                        MsgBox "El código del campo no existe o no es del socio.", vbExclamation
                        b = False
                        PonerFoco txtcodigo(50)
                    End If
                End If
            End If
            
        Case 17 ' recalculo de importe de venta campo
            If txtcodigo(52).Text = "" Then
                MsgBox "Debe introducir obligatoriamente el campo del socio.", vbExclamation
                b = False
                PonerFoco txtcodigo(52)
            End If
            
            If b Then
                If txtcodigo(46).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente un importe.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(46)
                End If
            End If
        
        Case 7 'deshacer facturacion de venta campo ( anticipo o liquidacion )
            If txtcodigo(9).Text = "" Or txtcodigo(10).Text = "" Then
                MsgBox "Debe introducir la primera y última factura de la Facturación", vbExclamation
                b = False
                PonerFoco txtcodigo(9)
'            Else
'                ' si la factura hasta no coincide con el contador de stipom no seguir
'                Select Case Combo1(1).ListIndex
'                    Case 0
'                        TipoMov = "FAC"
'                    Case 1
'                        TipoMov = "FLC"
'                End Select
'
'                Set vCont = New CTiposMov
'                If vCont.leer(TipoMov) Then
'                    If vCont.Contador <> CLng(txtCodigo(10).Text) Then
'                        MsgBox "La Factura hasta no es el último número de Factura. Revise.", vbExclamation
'                        b = False
'                    End If
'                End If
'                Set vCont = Nothing
            End If
            
            If b Then
                If txtcodigo(11).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Factura.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(11)
                End If
            End If
            ' comprobamos que si son anticipos no esten liquidados
            If b And tipoMov = "FAC" Then
                If AnticiposLiquidados(tipoMov, txtcodigo(9).Text, txtcodigo(10).Text, txtcodigo(11).Text) Then
                    MsgBox "Hay Facturas de Anticipos que han sido liquidadas. Revise.", vbExclamation
                    b = False
                    PonerFocoBtn cmdCancelDesF
                End If
            End If
            
       Case 13 'Prevision de pagos de liquidacion de industria
            If b And LiquidacionIndustria Then
                If txtcodigo(6).Text = "" Or txtcodigo(7) = "" Then
                    MsgBox "Para realizar la Previsión de Pago de Industria debe introducir obligatoriamente el rango de fechas.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(6)
                End If
            End If
            
       Case 9 ' informe de retenciones
            ' en el caso de certificado de retenciones obligamos a que nos introduzcan el rango
            ' de fechas que sale en el certificado
            If Check1(7).Value = 1 Then
                If txtcodigo(26).Text = "" Or txtcodigo(27).Text = "" Then
                    MsgBox "Debe introducir un valor en los campos de Fechas.", vbExclamation
                    b = False
                    If txtcodigo(26).Text = "" Then
                        PonerFoco txtcodigo(26)
                    Else
                        PonerFoco txtcodigo(27)
                    End If
                Else
                    If txtcodigo(32).Text = "" Then
                        MsgBox "Debe meter obligatoriamente la Fecha del Certificado.", vbExclamation
                        b = False
                        PonerFoco txtcodigo(32)
                    End If
                End If
            End If
            
         Case 14 ' factura de liquidaciones
            If b Then
                If txtcodigo(15).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de Liquidación.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(15)
                End If
            End If
         
        Case 19 ' factura de liquidacion directa
            If b Then
                If txtcodigo(61).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente la Fecha de Factura.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(61)
                End If
            End If
            
            If b Then
                If txtcodigo(60).Text = "" Then
                    MsgBox "Debe introducir obligatoriamente el precio calidad.", vbExclamation
                    b = False
                    PonerFoco txtcodigo(60)
                End If
            End If
         
    End Select
    DatosOk = b

End Function



Private Function ConcatenarCampos(ctabla As String, cwhere As String) As String
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String

    ConcatenarCampos = ""

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    SQL = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(ctabla, "_1")
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    
    
    SQL = "select distinct rcampos.codcampo  from " & ctabla & " where " & cwhere
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not RS.EOF
        Sql1 = Sql1 & DBLet(RS.Fields(0).Value, "N") & ","
        RS.MoveNext
    Wend
    Set RS = Nothing
    'quitamos el ultimo or
    ConcatenarCampos = Mid(Sql1, 1, Len(Sql1) - 1)
    
End Function

Private Function CargarTemporalAnticiposValsur(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
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
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposValsur = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid,"
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilos "
     SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilos "
    
    
    SQL = SQL & " FROM  " & ctabla

    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto, baseimpo, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!Codsocio
        VarieAnt = RS!codvarie
        NVarieAnt = RS!nomvarie
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(RS!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not RS.EOF
        '++monica:28/07/2009 añadida la segunda condicion
        If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            
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
            
            VarieAnt = RS!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
        End If
        
        If RS!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        Recolect = DBLet(RS!Recolect, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(RS!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(RS!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(RS!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
        
        
            Select Case Recolect
                Case 0
                    baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * PreCoop, 2)
                Case 1
                    baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * PreSocio, 2)
            End Select
            
        End If
        Set Rs9 = Nothing
        'hasta aqui
        
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
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
    
    CargarTemporalAnticiposValsur = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function

'[Monica]20/01/2012: nueva funcion de carga de anticipos de alzira que antes no tenia
Private Function CargarTemporalAnticiposAlzira(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
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
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposAlzira = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid,"
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilos "
     SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilos "
    
    
    SQL = SQL & " FROM  " & ctabla

    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto, baseimpo, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!Codsocio
        VarieAnt = RS!codvarie
        NVarieAnt = RS!nomvarie
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(RS!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                vPorcIva = ""
                '[Monica]29/04/2011: INTERNAS
                If vSocio.EsFactADVInt Then
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                Else
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not RS.EOF
        '++monica:28/07/2009 añadida la segunda condicion
        If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            
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
            
            VarieAnt = RS!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
        End If
        
        If RS!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    '[Monica]29/04/2011: INTERNAS
                    If vSocio.EsFactADVInt Then
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        Recolect = DBLet(RS!Recolect, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(RS!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(RS!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(RS!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
        
        
            Select Case Recolect
'                Case 0
'                    baseimpo = baseimpo + Round2(DBLet(Rs!Kilos, "N") * PreCoop, 2)
'                Case 1
'                    baseimpo = baseimpo + Round2(DBLet(Rs!Kilos, "N") * PreSocio, 2)
                Case 0
                    baseimpo = baseimpo + (DBLet(RS!Kilos, "N") * PreCoop)
                Case 1
                    baseimpo = baseimpo + (DBLet(RS!Kilos, "N") * PreSocio)
            End Select
            
        End If
        Set Rs9 = Nothing
        'hasta aqui
        
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
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
    
    CargarTemporalAnticiposAlzira = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalAnticiposPicassent(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SqlVar As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bonifica As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

Dim PorcBoni As Currency
Dim PorcComi As Currency

Dim ImporteFVar As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposPicassent = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    '[Monica]15/04/2013: introducimos las facturas varias
    Sql2 = "delete from tmpsuperficies where codusu = " & vUsu.Codigo
    conn.Execute Sql2


    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid, rhisfruta.fecalbar, "
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilos "
    SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilos "
    
    
    SQL = SQL & " FROM  " & ctabla

    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, Kneto, baseimpo, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac, bonificacion
    Sql2 = Sql2 & " porcen2, importeb1, importeb2, importeb3) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!Codsocio
        VarieAnt = RS!codvarie
        NVarieAnt = RS!nomvarie
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(RS!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not RS.EOF
        '++monica:28/07/2009 añadida la segunda condicion
        If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            
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
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "," & DBSet(Bonifica, "N") & "),"
            
            VarieAnt = RS!codvarie
            
            baseimpo = 0
            Bonifica = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
        End If
        
        If RS!Codsocio <> SocioAnt Then
            '[Monica]15/04/2013: descontamos las facturas varias
            If Check1(14).Value Then                                                                                                 'anticipos            no descontados   q no sean de ventacampo
                ImporteFVar = DevuelveValor("select sum(totalfac) from fvarcabfact where codsocio = " & DBSet(SocioAnt, "N") & " and enliquidacion = 2 and intliqui = 0 and envtacampo = 0 ")
                                                    'usuario, codsocio, importe facturas varias
                SqlVar = "insert into tmpsuperficies (codusu, codvarie, superficie1) values (" & vUsu.Codigo & ","
                SqlVar = SqlVar & DBSet(SocioAnt, "N") & ","
                SqlVar = SqlVar & DBSet(ImporteFVar, "N") & ")"
                conn.Execute SqlVar
            End If
        
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        Recolect = DBLet(RS!Recolect, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(RS!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(RS!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(RS!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
        
            PorcBoni = 0
            Select Case Recolect
                Case 0
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreCoop > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                    
                        '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                        PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                        If CCur(PorcComi) <> 0 Then
                            PreCoop = PreCoop - Round2(PreCoop * PorcComi / 100, 4)
                        End If
                    
                    End If
                
                    Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * PreCoop, 2)
                    baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * PreCoop * (1 + (PorcBoni / 100)), 2)
                Case 1
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreSocio > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                        
                        '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                        PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                        If CCur(PorcComi) <> 0 Then
                            PreSocio = PreSocio - Round2(PreSocio * PorcComi / 100, 4)
                        End If
                    End If
                
                    Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * PreSocio, 2)
                    baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * PreSocio * (1 + (PorcBoni / 100)), 2)
            End Select
            
        End If
        Set Rs9 = Nothing
        'hasta aqui
        
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        '[Monica]15/04/2013: descontamos las facturas varias
        If Check1(14).Value = 1 Then                                                                                            'anticipos             no descontados   que no sean de ventacampo
            ImporteFVar = DevuelveValor("select sum(totalfac) from fvarcabfact where codsocio = " & DBSet(SocioAnt, "N") & " and enliquidacion = 2 and intliqui = 0 and envtacampo = 0")
                                                'usuario, codsocio, importe facturas varias
            SqlVar = "insert into tmpsuperficies (codusu, codvarie, superficie1) values (" & vUsu.Codigo & ","
            SqlVar = SqlVar & DBSet(SocioAnt, "N") & ","
            SqlVar = SqlVar & DBSet(ImporteFVar, "N") & ")"
            conn.Execute SqlVar
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
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "," & DBSet(Bonifica, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalAnticiposPicassent = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalAnticiposCalidadPicassent(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CalidAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bonifica As Currency
Dim Importe As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

Dim PorcBoni As Currency
Dim PrecioAnt As Currency
Dim PorcComi As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposCalidadPicassent = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    '[Monica]15/04/2013: introducimos las facturas varias
    Sql2 = "delete from tmpsuperficies where codusu = " & vUsu.Codigo
    conn.Execute Sql2


    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio,  rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid, rhisfruta.fecalbar, "
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilos "
     SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilos "
    
    SQL = SQL & " FROM  " & ctabla
    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "


    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu,  codvarie, nomvarie, calidad, Kneto,  Precio, importe, bonificacion,
    Sql2 = "insert into tmpinformes (codusu,  importe1, nombre1, campo1, importe2, precio1, importe3, importe4, "
                   'importetotal
    Sql2 = Sql2 & " importe5) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        VarieAnt = RS!codvarie
        NVarieAnt = RS!nomvarie
        CalidAnt = RS!codcalid
    End If
    
    While Not RS.EOF
        '++monica:28/07/2009 añadida la segunda condicion
        If VarieAnt <> RS!codvarie Or CalidAnt <> RS!codcalid Then
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(CalidAnt, "N") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(PrecioAnt, "N") & ","
            Sql1 = Sql1 & DBSet(Importe, "N") & "," & DBSet(Bonifica - Importe, "N") & ","
            Sql1 = Sql1 & DBSet(Bonifica, "N") & "),"
            
            VarieAnt = RS!codvarie
            CalidAnt = RS!codcalid
            
            baseimpo = 0
            Bonifica = 0
            Importe = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        Recolect = DBLet(RS!Recolect, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(RS!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(RS!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(RS!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
            PorcBoni = 0
            PorcComi = 0
            Select Case Recolect
                Case 0
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreCoop > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                        
                        '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                        PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                        If CCur(PorcComi) <> 0 Then
                            PreCoop = PreCoop - Round2(PreCoop * PorcComi / 100, 4)
                        End If
                    End If
                    PrecioAnt = PreCoop
                    Importe = Importe + Round2(DBLet(RS!Kilos, "N") * PreCoop, 2)
                    Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * PreCoop * (1 + (PorcBoni / 100)), 2)
                Case 1
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreSocio > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                        
                        '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                        PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                        If CCur(PorcComi) <> 0 Then
                            PreSocio = PreSocio - Round2(PreSocio * PorcComi / 100, 4)
                        End If
                    End If
                    PrecioAnt = PreSocio
                    Importe = Importe + Round2(DBLet(RS!Kilos, "N") * PreSocio, 2)
                    Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * PreSocio * (1 + (PorcBoni / 100)), 2)
            End Select
            
        End If
        Set Rs9 = Nothing
        'hasta aqui
        
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        Sql1 = Sql1 & "(" & vUsu.Codigo & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(CalidAnt, "N") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(PrecioAnt, "N") & ","
        Sql1 = Sql1 & DBSet(Importe, "N") & "," & DBSet(Bonifica - Importe, "N") & ","
        Sql1 = Sql1 & DBSet(Bonifica, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalAnticiposCalidadPicassent = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargando temporal", Err.Description
End Function





Private Function CargarTemporalAnticiposCatadau(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
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
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposCatadau = False
    
    '[Monica]15/04/2013: introducimos las facturas varias
    Sql2 = "delete from tmpsuperficies where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    If CargarTemporalCatadau(ctabla, cwhere, 0) Then
        '[Monica]24/04/2013: pq en la anterio funcion se graba
        Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
        conn.Execute Sql2
    
        SQL = "SELECT tmpliquidacion.codsocio, tmpliquidacion.codvarie, variedades.nomvarie, "
        SQL = SQL & "sum(tmpliquidacion.kilosnet) as kilos, sum(tmpliquidacion.importe) as importe  "
        SQL = SQL & " FROM  tmpliquidacion, variedades "
        SQL = SQL & " WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " and tmpliquidacion.codvarie = variedades.codvarie "
        SQL = SQL & " group by 1,2,3 "
        SQL = SQL & " order by 1,2,3 "
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                                        'codusu, codsocio, nomsocio, codvarie, nomvarie, neto, baseimpo, porceiva, imporiva,
        Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                       'porcerete, imporret, totalfac
        Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
        
        Set vSeccion = New CSeccion
        
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If Not vSeccion.AbrirConta Then
                Exit Function
            End If
        End If
    
        HayReg = False
        If Not RS.EOF Then
            SocioAnt = RS!Codsocio
            VarieAnt = RS!codvarie
            NVarieAnt = RS!nomvarie
            
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                NSocioAnt = vSocio.Nombre
                TipoIRPF = vSocio.TipoIRPF
            End If
        End If
        
        While Not RS.EOF
            '++monica:28/07/2009 añadida la segunda condicion
            If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
                
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
                
                VarieAnt = RS!codvarie
                
                baseimpo = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                KilosNet = 0
            End If
            
            If RS!Codsocio <> SocioAnt Then
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(RS!Codsocio) Then
                    If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                    NSocioAnt = vSocio.Nombre
                End If
                SocioAnt = vSocio.Codigo
                TipoIRPF = vSocio.TipoIRPF
            End If
            
            KilosNet = DBLet(RS!Kilos, "N")
            
'            Sql3 = "select sum(gastos) from tmpliquidacion1 where codusu = " & vUsu.Codigo
'            Sql3 = Sql3 & " and codsocio = " & DBSet(Rs!CodSocio, "N")
'            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'
'            Gastos = DevuelveValor(Sql3)
'
'            baseimpo = DBLet(Rs!Importe, "N") - Gastos
'
            baseimpo = DBLet(RS!Importe, "N")
                
            HayReg = True
    
    
    
            RS.MoveNext
        Wend
        ' ultimo registro si ha entrado
        If HayReg Then
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
        
        CargarTemporalAnticiposCatadau = True
        Exit Function
        
    End If ' end del if de cargar temporal new
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function HayPreciosVariedadesValsur(Tipo As Byte, ctabla As String, cwhere As String, TipoPrecio As Byte) As Boolean
'Comprobar si hay precios para cada una de las variedades seleccionadas
' tipo: 0=anticipos
'       1=liquidaciones
' tipoprecio: 0 = precio recolectado cooperativa
'             1 = precio recolectado socio
'             2 = precio recolectado socio y cooperativa
Dim SQL As String
Dim vPrecios As CPrecios
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim Sql2 As String

    On Error GoTo eHayPreciosVariedadesValsur
    
    HayPreciosVariedadesValsur = False
    
    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    SQL = "Select distinct rhisfruta.codvarie FROM " & QuitarCaracterACadena(ctabla, "_1")
    Sql2 = "Select distinct rhisfruta_clasif.codvarie, rhisfruta_clasif.codcalid FROM " & QuitarCaracterACadena(ctabla, "_1")
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
        Sql2 = Sql2 & " where " & cwhere
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    '[Monica]25/02/2011: Añadido esto en alzira quieren poder hacer una liquidacion complementaria
    ' si estamos hacendo una liquidacion complementaria, cogemos los precios correspondientes
    If Tipo = 1 And Check1(5).Value = 1 Then
        Tipo = 3
    End If

    b = True
    ' comprobamos que existen registros para todos las variedades / calidades seleccionadas
    While Not RS.EOF And b
        Set vPrecios = New CPrecios
        b = vPrecios.Leer(CStr(Tipo), CStr(RS.Fields(0).Value), txtcodigo(6).Text, txtcodigo(7).Text)
'        If b Then b = vPrecios.ExistenPreciosCalidades
        If b Then b = vPrecios.ExisteAlgunPrecioCalidad(Sql2, TipoPrecio)
        Set vPrecios = Nothing
        
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
    
    HayPreciosVariedadesValsur = b
    Exit Function
    
eHayPreciosVariedadesValsur:
    MuestraError Err.nume, "Comprobando si hay precios en variedades", Err.Description
End Function

Private Function HayPreciosVariedadesCatadau(Tipo As Byte, ctabla As String, cwhere As String, TipoPrecio As Byte) As Boolean
'Comprobar si hay precios para cada una de las variedades seleccionadas
' tipo: 0=anticipos
'       1=liquidaciones
' tipoprecio: 0 = precio recolectado cooperativa
'             1 = precio recolectado socio
'             2 = precio recolectado socio y cooperativa
Dim SQL As String
Dim vPrecios As CPrecios
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim b As Boolean
Dim Sql2 As String
Dim Sql5 As String
Dim VarieAnt As Long
Dim Numreg As Long

    On Error GoTo eHayPreciosVariedadesCatadau
    
    HayPreciosVariedadesCatadau = False
    
    conn.Execute " DROP TABLE IF EXISTS tmpVarie;"
    
    SQL = "CREATE TEMPORARY TABLE tmpVarie ( " 'TEMPORARY
    SQL = SQL & "codvarie INT(6) UNSIGNED  DEFAULT '0' NOT NULL) "
    conn.Execute SQL
    
    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    SQL = "Select distinct rhisfruta_clasif.codvarie, rhisfruta_clasif.codcalid ,rhisfruta.fecalbar FROM " & QuitarCaracterACadena(ctabla, "_1")
    
'    Sql2 = "Select distinct rhisfruta_clasif.codvarie, rhisfruta_clasif.codcalid FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
'[Monica]14/03/2014: añadida esta condicion y la de abajo
        SQL = SQL & " and rhisfruta_clasif.kilosnet <> 0 and  rhisfruta_clasif.kilosnet is not null"
    Else
        SQL = SQL & " where rhisfruta_clasif.kilosnet <> 0 and  rhisfruta_clasif.kilosnet is not null"
    End If
    
    SQL = SQL & " group by 1,2,3 "
    SQL = SQL & " order by 1,2,3 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True
    
    ' si estamos hacendo una liquidacion complementaria, cogemos los precios correspondientes
    If Tipo = 1 And Check1(5).Value = 1 Then
        Tipo = 3
    End If
    
    
    If Not RS.EOF Then VarieAnt = DBLet(RS!codvarie, "N")
    Numreg = 0
    ' comprobamos que existen registros para todos las variedades / calidades seleccionadas
    While Not RS.EOF And b
    
    
    
        Sql2 = "select * from rprecios where (codvarie, tipofact, contador) = ("
        Sql2 = Sql2 & "SELECT codvarie, tipofact, max(contador) FROM rprecios WHERE codvarie=" & DBSet(RS!codvarie, "N") & " and "
        Sql2 = Sql2 & " tipofact = " & Tipo & " and fechaini <= " & DBSet(RS!Fecalbar, "F")
        Sql2 = Sql2 & " and fechafin >= " & DBSet(RS!Fecalbar, "F")
        Sql2 = Sql2 & " group by 1, 2) "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Rs2.EOF Then
            b = False
            MsgBox "No existen precios para la variedad " & DBLet(RS!codvarie, "N") & " de fecha " & DBLet(RS!Fecalbar, "F") & ". Revise.", vbExclamation
        Else
            If DBLet(RS!codvarie, "N") <> VarieAnt Then
                b = (Numreg <> 0)
                
                If Numreg <> 0 Then
                    Sql5 = "insert into tmpVarie (codvarie) values (" & DBSet(VarieAnt, "N") & ")"
                    conn.Execute Sql5
                End If
                Numreg = 0
                VarieAnt = DBLet(RS!codvarie, "N")
            End If
            ' miramos si hay alguna calidad para facturar
            Sql2 = "select count(*) from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and contador = " & DBSet(Rs2!Contador, "N")
            Sql2 = Sql2 & " and tipofact = " & Tipo
            Sql2 = Sql2 & " and codcalid = " & DBSet(RS!codcalid, "N")
'07/07/2014
            Select Case TipoPrecio
                Case 0
                    Sql2 = Sql2 & " and (rprecios_calidad.precoop <> 0 and not rprecios_calidad.precoop is null)"
                Case 1
                    Sql2 = Sql2 & " and (rprecios_calidad.presocio <> 0 and not rprecios_calidad.presocio is null)"
                Case 2
                    Sql2 = Sql2 & " and ((rprecios_calidad.precoop <> 0 and not rprecios_calidad.precoop is null) and (rprecios_calidad.presocio <> 0 and not rprecios_calidad.presocio is null)) "
            End Select
            Numreg = Numreg + TotalRegistros(Sql2)
        End If
            
        Set Rs2 = Nothing
        
        
        
        RS.MoveNext
    Wend
    'ultimo registro
    If b Then
        b = (Numreg <> 0)
        If Numreg <> 0 Then
            Sql5 = "insert into tmpVarie (codvarie) values (" & DBSet(VarieAnt, "N") & ")"
            conn.Execute Sql5
        End If
    End If
    RS.Close
    Set RS = Nothing
    
    HayPreciosVariedadesCatadau = b
    Exit Function
    
eHayPreciosVariedadesCatadau:
    MuestraError Err.nume, "Comprobando si hay precios en variedades", Err.Description
End Function


Private Function TotalFacturas(ctabla As String, cwhere As String) As Long
Dim SQL As String

    TotalFacturas = 0
    
    SQL = "SELECT  count(distinct rhisfruta.codsocio) "
    SQL = SQL & " FROM  " & ctabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If

    TotalFacturas = TotalRegistros(SQL)

End Function

Private Function TotalFacturasNew(ctabla As String, cwhere As String, cCampos As String) As Long
Dim SQL As String

    TotalFacturasNew = 0
    
    SQL = "SELECT  count(distinct " & cCampos & ") "
    SQL = SQL & " FROM  " & ctabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If

    TotalFacturasNew = TotalRegistros(SQL)

End Function



Private Sub ActivarCLAVE()
Dim i As Integer
    
    For i = 9 To 11
        txtcodigo(i).Enabled = False
    Next i
    txtcodigo(8).Enabled = True
    imgFec(6).Enabled = False
    CmdAcepDesF.Enabled = False
    cmdCancelDesF.Enabled = True
    Combo1(1).Enabled = False
End Sub

Private Sub DesactivarCLAVE()
Dim i As Integer

    For i = 9 To 11
        txtcodigo(i).Enabled = True
    Next i
    txtcodigo(8).Enabled = False
    imgFec(6).Enabled = True
    CmdAcepDesF.Enabled = True
    Combo1(1).Enabled = True
End Sub

Private Sub CargaCombo()

    On Error GoTo ErrCarga
    
    ' Tipo de facturacion venta campo (anticipo o liquidacion)
    ' para generacion de factura
    Combo1(0).Clear
    Combo1(0).AddItem "Anticipo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Liquidación"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
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
    If vParamAplic.Cooperativa = 2 Then
        Combo1(3).AddItem "Calidad"
        Combo1(3).ItemData(Combo1(3).NewIndex) = 2
    End If
    
    'recolectado por
    Combo1(5).Clear
    Combo1(5).AddItem "Cooperativa"
    Combo1(5).ItemData(Combo1(5).NewIndex) = 0
    Combo1(5).AddItem "Socio"
    Combo1(5).ItemData(Combo1(5).NewIndex) = 1
    Combo1(5).AddItem "Todos"
    Combo1(5).ItemData(Combo1(5).NewIndex) = 2


ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub

Private Sub CargarListView(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

 '   ListView1.ColumnHeaders.Add , , "Tipo", 650
    ListView1(Index).ColumnHeaders.Add , , "Tipo Movimiento", 2750
    
    SQL = "SELECT codtipom, nomtipom "
    SQL = SQL & " FROM usuarios.stipom "
    SQL = SQL & " WHERE stipom.tipodocu in (1,2,3,4,5,6,7,8,9,10,11) or stipom.codtipom = 'FTR'"
    SQL = SQL & " ORDER BY codtipom "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Set ItmX = ListView1(Index).ListItems.Add
        ItmX.Text = RS.Fields(1).Value ' Format(Rs.Fields(0).Value)
        ItmX.Key = RS.Fields(0).Value
'        ItmX.SubItems(1) = Rs.Fields(1).Value
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Tipo de Facturas.", Err.Description
    End If
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
    End Select
    
    NroTotalMovimientos = TotalRegistrosConsulta(SQL)

End Function



Private Function GeneraFicheroModelo(Tipo As Byte, pTabla As String, pWhere As String) As Boolean
Dim NFic As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim RS As ADODB.Recordset
Dim Aux As String
Dim Aux2 As String
Dim Cad As String
Dim Pagos As Boolean
Dim Concepto As Byte
Dim vSocio As cSocio
Dim b As Boolean
Dim Nregs As Long
Dim Total As Variant
Dim SQL As String

Dim ctabla As String
Dim vWhere As String
Dim Nombre As String
Dim CPostal As String

    On Error GoTo EGen
    GeneraFicheroModelo = False
    
    ctabla = pTabla
    vWhere = pWhere
    
    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    ctabla = QuitarCaracterACadena(ctabla, "_1")
    If vWhere <> "" Then
        vWhere = QuitarCaracterACadena(vWhere, "{")
        vWhere = QuitarCaracterACadena(vWhere, "}")
        vWhere = QuitarCaracterACadena(vWhere, "_1")
    End If
    
    NFic = FreeFile
    
    Open App.Path & "\modelo.txt" For Output As #NFic
    
    Select Case Tipo
        Case 0 ' MODELO 190
            ' previamente hemos de cargar una tabla intermedia donde sumemos base retencion por nif,
            ' cargaremos el nif del socio / transportista y nombre para cargar la linea 190
            
            SQL = " DROP TABLE IF EXISTS tmpReten;"
            conn.Execute SQL
            
            SQL = "CREATE TEMPORARY TABLE tmpReten ( " '
            SQL = SQL & "`nif` char(9) NOT NULL ,"
            SQL = SQL & "`codpostal` varchar(6) NOT NULL,"
            SQL = SQL & "`nombre` varchar(40) NOT NULL,"
            SQL = SQL & "`tipo` tinyint(1) NOT NULL,"
            SQL = SQL & "`basereten` decimal(8,2) NOT NULL,"
            SQL = SQL & "`impreten` decimal(8,2) NOT NULL default '0')"
            
            conn.Execute SQL
            
'[Monica]20/01/2014: no enlazamos con los trasnportistas, pq en alzira han cambiado la codificacion
'            Sql = " insert into tmpReten (nif, codpostal, nombre, tipo, basereten, impreten) "
'            Sql = Sql & "select nifsocio, codpostal, tmprfactsoc.nomsocio, 0, sum(basereten), sum(impreten) "
'            Sql = Sql & " from tmprfactsoc, rsocios where codusu = " & vUsu.Codigo
'            Sql = Sql & " and tmprfactsoc.tipo = 0 " ' solo los socios
'            'Sql = Sql & " and tmprfactsoc.codsocio = cast(rsocios.codsocio as char) "
'            Sql = Sql & " and tmprfactsoc.codsocio = rsocios.codsocio "
'            Sql = Sql & " group by 1,2,3,4 "
'            Sql = Sql & " union"
'            Sql = Sql & " select niftrans, codpostal, nomtrans, if(tmprfactsoc.tipoirpf<=2,0,1), sum(basereten), sum(impreten) "
'            Sql = Sql & " from tmprfactsoc, rtransporte where codusu = " & vUsu.Codigo
'            Sql = Sql & " and tmprfactsoc.tipo = 1 " ' solo los transportistas
'            Sql = Sql & " and tmprfactsoc.codsocio = rtransporte.codtrans "
'            Sql = Sql & " group by 1,2,3,4 "
'            Sql = Sql & " order by 1,2,3,4 "
            
            SQL = " insert into tmpReten (nif, codpostal, nombre, tipo, basereten, impreten) "
            SQL = SQL & "select nif, codpostal, tmprfactsoc.nomsocio, 0, sum(basereten), sum(impreten) "
            SQL = SQL & " from tmprfactsoc where codusu = " & vUsu.Codigo
            SQL = SQL & " and tmprfactsoc.tipo = 0 " ' solo los socios
            SQL = SQL & " group by 1,2,3,4 "
            SQL = SQL & " union"
            SQL = SQL & " select nif, codpostal , nomsocio, if(tmprfactsoc.tipoirpf<=2,0,1), sum(basereten), sum(impreten) "
            SQL = SQL & " from tmprfactsoc where codusu = " & vUsu.Codigo
            SQL = SQL & " and tmprfactsoc.tipo = 1 " ' solo los transportistas
            SQL = SQL & " group by 1,2,3,4 "
            SQL = SQL & " union "
            SQL = SQL & " select nif, codpostal , nomsocio, if(tmprfactsoc.tipoirpf<=2,0,1), sum(basereten), sum(impreten) "
            SQL = SQL & " from tmprfactsoc where codusu = " & vUsu.Codigo
            SQL = SQL & " and tmprfactsoc.tipo = 2 " ' socios terceros
            SQL = SQL & " group by 1,2,3,4 "
            SQL = SQL & " order by 1,2,3,4 "
            
            conn.Execute SQL
            
            SQL = "select count(distinct nif, tipo), sum(basereten), sum(impreten) "
            SQL = SQL & " from tmpReten "
            
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
            ' CABECERA
            Cabecera190b NFic, CLng(DBLet(RS.Fields(0).Value, "N")), CCur(DBLet(RS.Fields(2).Value, "N")), CCur(DBLet(RS.Fields(1).Value, "N"))
            
            Set RS = Nothing
            
            ' LINEAS
            
            'Imprimimos las lineas
            Aux = "select tmpreten.nif, tmpreten.tipo, sum(tmpreten.basereten), sum(tmpreten.impreten) "
            Aux = Aux & " from tmpreten "
            Aux = Aux & " group by 1,2 "
            Aux = Aux & " having sum(tmpreten.basereten) <> 0 "
            Aux = Aux & " order by 1,2 "
            
            Set RS = New ADODB.Recordset
            RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If RS.EOF Then
                'No hayningun registro
            Else
                b = True
                Regs = 0
                While Not RS.EOF And b
                    Regs = Regs + 1
                    
                    Nombre = DevuelveValor("select nombre from tmpreten where nif = " & DBSet(RS!NIF, "T"))
                    CPostal = DevuelveValor("select codpostal from tmpreten where nif = " & DBSet(RS!NIF, "T"))
                    
                    Linea190new NFic, Nombre, CPostal, RS
                    
                    RS.MoveNext
                Wend
            End If
            RS.Close
            Set RS = Nothing
            
   
            
            
        Case 1 ' MODELO 346
            Aux = "select tmp346.codsocio, tmp346.codgrupo, sum(tmp346.importe) "
            Aux = Aux & " from tmp346 "
            Aux = Aux & " where " & Replace(vWhere, "tmprfactsoc", "tmp346") & " and tmp346.codgrupo in (4,5) " ' algarrobos y olivos
            Aux = Aux & " group by tmp346.codsocio, tmp346.codgrupo "
            Aux = Aux & "  union "
            Aux = Aux & " select tmp346.codsocio, 0, sum(tmp346.importe) "
            Aux = Aux & " from tmp346 "
            Aux = Aux & " where " & Replace(vWhere, "tmprfactsoc", "tmp346") & " and not tmp346.codgrupo in (4,5) " ' el resto
            Aux = Aux & " group by tmp346.codsocio, tmp346.codgrupo "
            Aux = Aux & " order by 1,2"
        
            Nregs = TotalRegistrosConsulta(Aux)
        
            If Nregs <> 0 Then
                Aux2 = "select sum(tmp346.importe) from tmp346 "
                Aux2 = Aux2 & " where " & Replace(vWhere, "tmprfactsoc", "tmp346")
                
                Total = DevuelveValor(Aux2)
            
                Cabecera346 NFic, Nregs, CCur(Total)
            
                Set RS = New ADODB.Recordset
                RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If RS.EOF Then
                    'No hayningun registro
                Else
                    b = True
                    Regs = 0
                    While Not RS.EOF And b
                        Regs = Regs + 1
                        Set vSocio = New cSocio
                        
                        If vSocio.LeerDatos(DBLet(RS!Codsocio, "N")) Then
                            Linea346 NFic, vSocio, RS
                        Else
                            b = False
                        End If
                        
                        Set vSocio = Nothing
                        RS.MoveNext
                    Wend
                End If
                RS.Close
                Set RS = Nothing
                
            End If
    End Select
    Close (NFic)
    
    If Regs > 0 Then GeneraFicheroModelo = True
    Exit Function
    
EGen:
    Set RS = Nothing
    Close (NFic)
    MuestraError Err.Number, Err.Description
End Function

Private Function CopiarFichero() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.CommonDialog1.DefaultExt = "txt"
    
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    ' copiamos el primer fichero
    Select Case OpcionListado
        Case 10
            CommonDialog1.FileName = "modelo190.txt"
        Case 11
            CommonDialog1.FileName = "modelo346.txt"
    End Select
        
    Me.CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        FileCopy App.Path & "\modelo.txt", CommonDialog1.FileName
    End If
    
    CopiarFichero = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function




Private Sub Cabecera190b(NFich As Integer, Nregs As Currency, ImpReten As Currency, BaseReten As Currency)
Dim Cad As String

'TIPO DE REGISTRO 1:REGISTRO DEL RETENEDOR}
    
    Cad = "1190"                                                  'p.1
    Cad = Cad & Format(txtcodigo(30).Text, "0000")                'p.5 año de ejercicio
    Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 9)       'p.9 cif empresa
    Cad = Cad & RellenaABlancos(SinCaracteresRaros(vParam.NombreEmpresa), True, 40)   'p.18 nombre de empresa
    '[Monica]20/01/2016: antes era una D
    Cad = Cad & "T"                                               'p.58 antes era D
    Cad = Cad & RellenaAceros(txtcodigo(37).Text, True, 9)        'p.59 telefono
    Cad = Cad & RellenaABlancos(SinCaracteresRaros(txtcodigo(36).Text), True, 40)     'p.68 persona de contacto
    Cad = Cad & RellenaAceros(txtcodigo(31).Text, True, 13)       'p.108 nro de justificante
    Cad = Cad & Space(2)                                          'p.121 ni es complementaria ni sustitutiva
    Cad = Cad & RellenaAceros("0", True, 13)                      'p.123 13 ceros (justificante de la complementaria o sustitutiva)
    Cad = Cad & Format(Nregs, "000000000")                        'p.136 nro de registros

    If BaseReten < 0 Then
        Cad = Cad & "N"                                           'p.145 signo de retenciones
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(BaseReten * (-1) * 100)), False, 15)    'p.146
    Else
        Cad = Cad & " "                                           'p.145
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(BaseReten * 100)), False, 15)           'p.146
    End If
              
    If ImpReten < 0 Then                                          'p.161
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(ImpReten * (-1) * 100)), False, 15)
    Else
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(ImpReten * 100)), False, 15)
    End If
    Cad = Cad & Space(322) 'p.176 a 487                    'antes:  Space(62)                             'p.176
    Cad = Cad & Space(3)   'p.488 a 500 firma digital      'antes:  Space(13)                                         'p.238

    Print #NFich, Cad

End Sub


Private Sub Linea190(NFich As Integer, vSocio As cSocio, ByRef RS As ADODB.Recordset)
Dim Cad As String

    Cad = "2190"                                                'p.1
    Cad = Cad & Format(txtcodigo(30).Text, "0000")              'p.5 año ejercicio
    Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 9)     'p.9 cif empresa
    Cad = Cad & RellenaABlancos(vSocio.NIF, True, 9)            'p.18 nifsocio
    Cad = Cad & Space(9)                                        'p.27 nif del representante legal
    Cad = Cad & RellenaABlancos(vSocio.Nombre, True, 40)        'p.36 nombre socio
    Cad = Cad & RellenaABlancos(Mid(vSocio.CPostal, 1, 2), True, 2) 'p.76 codpobla[1,2] codigo de provincia
    Cad = Cad & "H"                                             'p.78 clave de percepcion H=actividades agrícolas, ganaderas y forestales
    Cad = Cad & "01"                                            'p.79 subclave:
'                                                                       01 =  Se consignará esta subclave cuando se trate de percepciones
'                                                                        a las que resulte aplicable el tipo de retención establecido
'                                                                        con carácter general en el artículo 95.4.2º del Reglamento
'                                                                        del Impuesto.
   
'[Monica]: 14/01/2010
' antes no estaba en el if de abajo siempre era un blanco lo he cambiado según el signo.
'    cad = cad & " "                                             'p.81
    
    If DBLet(RS.Fields(1).Value, "N") < 0 Then                  'p.82 base de retencion
        Cad = Cad & "N"                                             'p.81
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(1).Value, "N") * (-1) * 100)), False, 13)
    Else
        Cad = Cad & " "                                             'p.81
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(1).Value, "N") * 100)), False, 13)
    End If
    
    If DBLet(RS.Fields(2).Value, "N") < 0 Then                  'p.95 importe de retencion
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * (-1) * 100)), False, 13)
    Else
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * 100)), False, 13)
    End If
    
    Cad = Cad & " "                                             'p.108
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.109
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.122
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.135
    Cad = Cad & RellenaAceros("0", True, 4)                     'p.148
    Cad = Cad & "0"                                             'p.152
    Cad = Cad & RellenaAceros("0", True, 5)                     'p.153
    Cad = Cad & RellenaABlancos(" ", True, 9)                   'p.158
    Cad = Cad & String(88, "0")                                 'p.167  antes eran 84 ceros
    Cad = Cad & Space(246)                                      'p.255 - 500 se rellenan a blancos
    
    Print #NFich, Cad
End Sub



Private Sub Linea190new(NFich As Integer, Nombre As String, CPostal As String, ByRef RS As ADODB.Recordset)
Dim Cad As String

    Cad = "2190"                                                'p.1
    Cad = Cad & Format(txtcodigo(30).Text, "0000")              'p.5 año ejercicio
    Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 9)     'p.9 cif empresa
    Cad = Cad & RellenaABlancos(RS!NIF, True, 9)            'p.18 nifsocio
    Cad = Cad & Space(9)                                        'p.27 nif del representante legal
    Cad = Cad & RellenaABlancos(SinCaracteresRaros(Nombre), True, 40)        'p.36 nombre socio
    Cad = Cad & RellenaABlancos(Mid(CPostal, 1, 2), True, 2) 'p.76 codpobla[1,2] codigo de provincia
    Cad = Cad & "H"                                             'p.78 clave de percepcion H=actividades agrícolas, ganaderas y forestales
    If RS!Tipo = 0 Then
        Cad = Cad & "01"                                            'p.79 subclave:
    Else
        Cad = Cad & "04"
    End If
'                                                                       01 =  Se consignará esta subclave cuando se trate de percepciones
'                                                                        a las que resulte aplicable el tipo de retención establecido
'                                                                        con carácter general en el artículo 95.4.2º del Reglamento
'                                                                        del Impuesto.
'                                                                       04 =  Cuando es regimen de transportista
'
'[Monica]: 14/01/2010
' antes no estaba en el if de abajo siempre era un blanco lo he cambiado según el signo.
'    cad = cad & " "                                             'p.81
    
    If DBLet(RS.Fields(2).Value, "N") < 0 Then                  'p.82 base de retencion
        Cad = Cad & "N"                                             'p.81
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * (-1) * 100)), False, 13)
    Else
        Cad = Cad & " "                                             'p.81
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * 100)), False, 13)
    End If
    
    If DBLet(RS.Fields(3).Value, "N") < 0 Then                  'p.95 importe de retencion
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(3).Value, "N") * (-1) * 100)), False, 13)
    Else
        Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(3).Value, "N") * 100)), False, 13)
    End If
    
    Cad = Cad & " "                                             'p.108
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.109
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.122
    Cad = Cad & RellenaAceros("0", True, 13)                    'p.135
    Cad = Cad & RellenaAceros("0", True, 4)                     'p.148
    Cad = Cad & "0"                                             'p.152
    Cad = Cad & RellenaAceros("0", True, 5)                     'p.153
    Cad = Cad & RellenaABlancos(" ", True, 9)                   'p.158
    '[Monica]20/01/2016: he puesto 2 ceros y 1 blanco, antes 3 ceros
    Cad = Cad & "00 "                                           'p.167
    Cad = Cad & String(85, "0")                                 'p.170  antes eran 88 ceros
    Cad = Cad & Space(246)                                      'p.255 - 500 se rellenan a blancos
    
    Print #NFich, Cad
End Sub





Private Sub Cabecera346(NFich As Integer, Nregs As Long, Total As Currency)
Dim Cad As String

   'TIPO DE REGISTRO 0:PRESENTACION COLECTIVA
    Cad = "1346"                                                'p.1
    Cad = Cad & Format(txtcodigo(30).Text, "0000")              'p.5 año ejercicio
    Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 9)     'p.9 cif empresa
    Cad = Cad & RellenaABlancos(vParam.NombreEmpresa, True, 40) 'p.18 nombre empresa
    Cad = Cad & "D"    'p.58 siglas
    Cad = Cad & RellenaAceros(txtcodigo(37).Text, False, 9)     'p.59 telefono
    Cad = Cad & RellenaABlancos(txtcodigo(36).Text, True, 40)   'p.68 persona de contacto
    Cad = Cad & RellenaAceros(txtcodigo(31).Text, False, 13)    'p.108 nro justificante
    Cad = Cad & Space(2)                                        ' contar posiciones en multibase
    Cad = Cad & String(13, "0")                                 'p.122
    Cad = Cad & RellenaAceros(CStr(Nregs), False, 9)            'p.136 numero de registros
    Cad = Cad & Space(1)                                        ' contar posiciones en multibase
    Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(Total * 100)), False, 17)  'p.146 importe total
    Cad = Cad & Space(87)                                       'p.163
    Cad = Cad & Space(251)
    
    Print #NFich, Cad
End Sub


Private Sub Linea346(NFich As Integer, vSocio As cSocio, ByRef RS As ADODB.Recordset)
Dim Cad As String
          
    Cad = "2346"                                                'p.1
    Cad = Cad & Format(txtcodigo(30).Text, "0000")              'p.5 año ejercicio
    Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 9)     'p.9 cif empresa
    Cad = Cad & RellenaABlancos(vSocio.NIF, True, 18)            'p.18 nifsocio
    Cad = Cad & RellenaABlancos(SinCaracteresRaros(vSocio.Nombre), True, 40)        'p.36 nombre socio
    Cad = Cad & RellenaABlancos(Mid(vSocio.CPostal, 1, 2), True, 2) 'p.76 codpobla[1,2]
    '[Monica]21/02/2012: la clave de percepcion es una B antes A
    Cad = Cad & "B"                                             'p.78
    
    Select Case DBLet(RS.Fields(1).Value, "N")
        Case 0
            '[Monica]21/02/2012: el tipo de percepcion un 2 antes era un 6
            Cad = Cad & "2"                                             'p.79
        Case 4
            Cad = Cad & "1"                                             'p.79
        Case 5
            Cad = Cad & "1"                                             'p.79
    End Select
    
    Cad = Cad & " "                                             ' contar posiciones en multibase
    Cad = Cad & RellenaAceros(ImporteSinFormato(CStr(DBLet(RS.Fields(2).Value, "N") * 100)), False, 14) 'p.81 base imponible
    Cad = Cad & RellenaAceros("0", True, 4)                     'p.95
    
    Select Case DBLet(RS.Fields(1).Value, "N")
        Case 0
            '[Monica]17/02/2012: en Alzira es subvencion fondo operativo
            If vParamAplic.Cooperativa = 4 Then
                Cad = Cad & RellenaABlancos("SUBVENCION FONDO OPERATIVO", True, 57)   'p.99
            Else
                Cad = Cad & RellenaABlancos("INDEMNIZACION AGROSEGURO", True, 57)   'p.99
            End If
        Case 4
            Cad = Cad & RellenaABlancos("CULTIVO ALGARROBO", True, 57)          'p.99
        Case 5
            Cad = Cad & RellenaABlancos("CULTIVO OLIVO", True, 57)              'p.99
    End Select
    '[Monica]21/02/2012: antes no habia nada en la clave de caracter de intervencion
    Cad = Cad & "2"  'clave caracter intervencion
    Cad = Cad & Space(94)                                       'p.156
    Cad = Cad & Space(250)
    
    Print #NFich, Cad
End Sub

Private Function SinCaracteresRaros(vCadena As String) As String

    SinCaracteresRaros = Replace(Replace(Replace(Replace(vCadena, "ª", " "), "º", " "), "ç", " "), "'", " ")

End Function



Private Function CargarTemporalLiquidacionValsur(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
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
Dim Cad As String
Dim HayReg As Boolean
Dim vGastos As Currency


    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionValsur = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid,"
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilos "
     SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilos "
    SQL = SQL & " FROM  " & ctabla

    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  bonificacion, gastos,  anticipos, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac, max(contador),tipofact, rprecios.fecini, rprecios.fecfin
    Sql2 = Sql2 & " porcen2, importeb1, importeb2, campo1, campo2, fecha1, fecha2) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!Codsocio
        VarieAnt = RS!codvarie
        NVarieAnt = RS!nomvarie
        CampoAnt = RS!codcampo
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(RS!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not RS.EOF
        If CampoAnt <> RS!codcampo Then
            ' gastos por campo
            
            ' [MONICA] : 08/09/2009 los gastos de transporte son una bonificacion para Valsur
            '            Se restan del resto de gastos
            'Sql4 = "select sum(if(isnull(imptrans),0,imptrans)) + sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) as gastos from rhisfruta "
            Sql4 = "select sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) as gastos from rhisfruta "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
            Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
            Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F")
            
            ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
            
            
            Sql4 = "select sum(if(isnull(imptrans),0,imptrans)) as bonif from rhisfruta "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
            Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
            Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F")
            
            ImpoBonif = ImpoBonif + DevuelveValor(Sql4)
             
            '[MONICA] : 15/03/2010 si es complementaria los gastos son 0
            If Check1(5).Value = 1 Then
                ImpoBonif = 0
            End If
            'fin 15/03/2010
            
            
            CampoAnt = RS!codcampo
        End If
    
        ' 23/07/2009: añadido el or con la segunda condicion
        If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
            
            Anticipos = DevuelveValor(Sql4)
            
            Bruto = baseimpo
            ImpoGastos = ImpoGastos
            
            ImpoBonif = ImpoBonif
            
'[MONICA] : 08/09/2009 he quitado lo de David pq los gastos de transporte los he quitado arriba
'           dejo lo original
'
'            'FALTA###
'            'DAVID###   20 Agosto 2009
'            'Si es para valsur los gastos se le suman, NO se le restan
'            ' Habria que ver:
'            '   -Los gastos del campo(el punto de arriba)
'            '   -Si en esta funcion solo entra valsur no haria falta poner vParamAplic.Cooperativa = 1
'
'            If vParamAplic.Cooperativa = 1 Then
'                baseimpo = baseimpo + ImpoGastos - Anticipos  'valsur
'            Else
'                baseimpo = baseimpo - ImpoGastos - Anticipos  'original
'            End If
            
            'DAVID###
            'El gasto de la cooperativa siempre se lo quito al total
'            baseimpo = baseimpo - Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)

'[Monica] 09/09/2009: el gasto de la cooperativa lo añado a la columna de gastos que no usa Valsur
            ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
            
            '[MONICA] : 15/03/2010 si es complementaria los gastos son 0
            If Check1(5).Value = 1 Then
                ImpoBonif = 0
            End If
            'fin 15/03/2010
            
            baseimpo = baseimpo + ImpoBonif - ImpoGastos - Anticipos
            
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
        
            ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoBonif, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N")
'02/09/2010
'            Sql1 = Sql1 & "),"
            Sql1 = Sql1 & ","
            Sql1 = Sql1 & DBSet(RS!Contador, "N") & "," & DBSet(RS!TipoFact, "N") & "," & DBSet(RS!FechaIni, "F") & "," & DBSet(RS!FechaFin, "F") & "),"
            
            
            VarieAnt = RS!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
            
            ImpoGastos = 0
            ImpoBonif = 0
            Anticipos = 0
            
        End If
        
        If RS!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        Dim vConta As String
        Dim vFecIni As String
        Dim vFecFin As String
        Dim vTipo As String
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(RS!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(RS!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(RS!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
        
            Recolect = DBLet(RS!Recolect, "N")
            Select Case Recolect
                Case 0
                    baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * PreCoop, 2)
                Case 1
                    baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * PreSocio, 2)
            End Select
        End If
        Set Rs9 = Nothing
        
        vConta = RS!Contador
        vFecIni = DBLet(RS!FechaIni, "F")
        vFecFin = DBLet(RS!FechaFin, "F")
        vTipo = RS!TipoFact
        'hasta aqui
            
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        Bruto = baseimpo
        
        
        ' anticipos
        Sql4 = "select sum(rfactsoc_variedad.imporvar) "
        Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
        Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
        Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
        Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
        
        Anticipos = DevuelveValor(Sql4)
            
        
        ' gastos por campo
        Sql4 = "select  sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) as gastos from rhisfruta "
        Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
        Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
        Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F")
        
        ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
                
        Sql4 = "select sum(if(isnull(imptrans),0,imptrans)) as bonif from rhisfruta "
        Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
        Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
        Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F")
        
                
        ImpoBonif = ImpoBonif + DevuelveValor(Sql4)


'        If vParamAplic.Cooperativa = 1 Then
'            baseimpo = baseimpo + ImpoGastos - Anticipos  'valsur
'        Else
'            baseimpo = baseimpo - ImpoGastos - Anticipos  'original
'        End If
        
        'DAVID###
        'El gasto de la cooperativa siempre se lo quito al total
'        baseimpo = baseimpo - Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
        
'[Monica] 09/09/2009: el gasto de la cooperativa lo añado a la columna de gastos que no usa Valsur
        ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
        
        '[MONICA] : 15/03/2010 si es complementaria los gastos son 0
        If Check1(5).Value = 1 Then
            ImpoBonif = 0
        End If
        'fin 15/03/2010
        
        
        baseimpo = baseimpo + ImpoBonif - ImpoGastos - Anticipos
        
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
    
        ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
    
        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
        Sql1 = Sql1 & DBSet(Bruto, "N") & ","
        Sql1 = Sql1 & DBSet(ImpoBonif, "N") & ","
        Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
        Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'        Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
'02/09/2010
'        Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "," & DBSet(vConta, "N") & "," & DBSet(vTipo, "N") & "," & DBSet(vFecIni, "F") & "," & DBSet(vFecFin, "F") & "),"
    
    
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalLiquidacionValsur = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function



Private Function CargarTemporalLiquidacionPicassent(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim SqlLiq As String

Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CampoAnt As Long
Dim AlbarAnt As Long
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
Dim Bonifica As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim PorcBoni As Currency
Dim vPorcIva As String
Dim vPorcGasto As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean
Dim vGastos As Currency


Dim BaseImpoFactura As Currency
Dim ImpoIvaFactura As Currency
Dim ImpoAporFactura As Currency
Dim ImpoRetenFactura As Currency
Dim ImpoGastosFactura As Currency
Dim ImpoTotalFactura As Currency
Dim ImpoFrasVarias As Currency

Dim SqlFactura As String
Dim sqlLiquid As String
Dim ImpBonif As Currency
Dim ImpTot As Currency

Dim PorcComi As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionPicassent = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql2 = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql2 = "delete from tmpfactura where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, rhisfruta.numalbar, rhisfruta.fecalbar, "
    SQL = SQL & "rhisfruta.recolect,  rhisfruta_clasif.codcalid, rcalidad.nomcalid,"
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilos "
    SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilos "
    SQL = SQL & " FROM  " & ctabla

    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  bonificacion, gastos,  anticipos, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac, max(contador),tipofact, rprecios.fecini, rprecios.fecfin
    Sql2 = Sql2 & " porcen2, importeb1, importeb2, campo1, campo2, fecha1, fecha2) values "
    
    'cargamos las bonificaciones para el informe de liquidacion
                                                                                'albaran            %bonif  impbonif, total
    SqlLiq = "insert into tmpliquidacion (codusu, codsocio, codvarie, codcampo, kilosnet, codcalid, precio, importe, gastos) values "
    
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!Codsocio
        VarieAnt = RS!codvarie
        NVarieAnt = RS!nomvarie
        CampoAnt = RS!codcampo
        AlbarAnt = RS!numalbar
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(RS!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    Bonifica = 0
    baseimpo = 0
    KilosNet = 0
    ImpoGastos = 0
    
    BaseImpoFactura = 0
    ImpoIvaFactura = 0
    ImpoAporFactura = 0
    ImpoRetenFactura = 0
    ImpoTotalFactura = 0
    ImpoGastosFactura = 0
    
    
    sqlLiquid = ""
    
    While Not RS.EOF
        If AlbarAnt <> RS!numalbar Or VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            ' gastos de los albaranes
            Sql4 = "select sum(rhisfruta_gastos.importe) "
            Sql4 = Sql4 & " from rhisfruta_gastos "
            Sql4 = Sql4 & " where rhisfruta_gastos.numalbar = " & DBSet(AlbarAnt, "N")
            
            ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
            
            '[Monica]23/07/2012: si es complementaria no hay gastos
            If Check1(5).Value = 1 Then ' si es complementaria no hay gastos
                ImpoGastos = 0
            End If
            
            ImpoGastosFactura = ImpoGastosFactura + DevuelveValor(Sql4)
            
            AlbarAnt = RS!numalbar
        End If
    
        ' 23/07/2009: añadido el or con la segunda condicion
        If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
            
            Anticipos = DevuelveValor(Sql4)
            
            Bruto = baseimpo - Bonifica
            
            ImpoBonif = Bonifica
            'ImpoBonif = BaseImpo - Bonifica
            
            baseimpo = baseimpo - Anticipos
            
            BaseImpoFactura = BaseImpoFactura + baseimpo
            
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
        
            If Check1(5).Value = 1 Then ' si es complementaria no hay importe de aportacion
                ImpoAport = 0
            Else
                ImpoAport = DevuelveValor("select importe from raporreparto where codsocio = " & DBSet(SocioAnt, "N") & " and tipoentr = 0")
            End If
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
            TotalFac = TotalFac - ImpoGastos
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoBonif, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N")
'02/09/2010
'            Sql1 = Sql1 & "),"
            Sql1 = Sql1 & ","
            Sql1 = Sql1 & DBSet(RS!Contador, "N") & "," & DBSet(RS!TipoFact, "N") & "," & DBSet(RS!FechaIni, "F") & "," & DBSet(RS!FechaFin, "F") & "),"
            
            
            VarieAnt = RS!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
            
            ImpoGastos = 0
            ImpoBonif = 0
            Anticipos = 0
            Bonifica = 0
            
        End If
        
        If RS!Codsocio <> SocioAnt Then
            ImpoIvaFactura = Round2((BaseImpoFactura) * ComprobarCero(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoRetenFactura = Round2((BaseImpoFactura + ImpoIvaFactura) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoRetenFactura = Round2(BaseImpoFactura * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoRetenFactura = 0
                    PorcReten = 0
            End Select
        
            If Check1(5).Value = 1 Then ' si es complementaria no hay importe de aportacion
                ImpoAporFactura = 0
            Else
                ImpoAporFactura = DevuelveValor("select importe from raporreparto where codsocio = " & DBSet(SocioAnt, "N") & " and tipoentr = 0")
            End If
            
            '[Monica]15/04/2013: si hay importe de facturas varias a descontar del socio
            ImpoFrasVarias = 0
            If Check1(14).Value = 1 Then                                                                                      'en liquidacion       no descontada         que no sea vtacampo
                ImpoFrasVarias = DevuelveValor("select sum(totalfac) from fvarcabfact where codsocio = " & DBSet(SocioAnt, "N") & " and enliquidacion = 1 and intliqui = 0 and envtacampo = 0")
            End If
            
            ImpoTotalFactura = BaseImpoFactura + ImpoIvaFactura - ImpoRetenFactura - ImpoAporFactura - ImpoGastosFactura '- ImpoFrasVarias
            
            SqlFactura = "insert into tmpfactura(codusu,codsocio,baseimpo,imporiva,impreten,impapor,impgastos,totalfac, impfrasvar) values ( "
            SqlFactura = SqlFactura & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(BaseImpoFactura, "N") & ","
            SqlFactura = SqlFactura & DBSet(ImpoIvaFactura, "N") & "," & DBSet(ImpoRetenFactura, "N") & ","
            SqlFactura = SqlFactura & DBSet(ImpoAporFactura, "N") & "," & DBSet(ImpoGastosFactura, "N") & ","
            SqlFactura = SqlFactura & DBSet(ImpoTotalFactura, "N") & "," & DBSet(ImpoFrasVarias, "N") & ")"
            
            conn.Execute SqlFactura
            
            BaseImpoFactura = 0
            ImpoIvaFactura = 0
            ImpoRetenFactura = 0
            ImpoAporFactura = 0
            ImpoGastosFactura = 0
            ImpoTotalFactura = 0
            ImpoFrasVarias = 0
            
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        Dim vConta As String
        Dim vFecIni As String
        Dim vFecFin As String
        Dim vTipo As String
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(RS!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(RS!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(RS!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
            PorcBoni = 0
            Recolect = DBLet(RS!Recolect, "N")
            PorcComi = 0
            Select Case Recolect
                Case 0
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreCoop > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                        
                        '[Monica]22/03/2012: Solo si le indicamos que no calcule comision no lo hace (solo prevision de liquidacion)
                        If Check1(13).Value Then
                            '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                            PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                            If CCur(PorcComi) <> 0 Then
                                PreCoop = PreCoop - Round2(PreCoop * PorcComi / 100, 4)
                            End If
                        End If
                    End If
                
'                    Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * PreCoop, 2)
'                    BaseImpo = BaseImpo + Round2(DBLet(RS!Kilos, "N") * PreCoop * (1 + (PorcBoni / 100)), 2)
                    
                    ImpBonif = Round2(DBLet(RS!Kilos, "N") * PreCoop * (PorcBoni / 100), 2)
                    ImpTot = Round2(DBLet(RS!Kilos, "N") * PreCoop, 2) + ImpBonif
                
                    Bonifica = Bonifica + ImpBonif
                    baseimpo = baseimpo + ImpTot
                
                
                Case 1
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreSocio > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                        
                        '[Monica]22/03/2012: Solo si le indicamos que no calcule comision no lo hace (solo prevision de liquidacion)
                        If Check1(13).Value Then
                            '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                            PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                            If CCur(PorcComi) <> 0 Then
                                PreSocio = PreSocio - Round2(PreSocio * PorcComi / 100, 4)
                            End If
                        End If
                    End If
                    
                    ImpBonif = Round2(DBLet(RS!Kilos, "N") * PreSocio * (PorcBoni / 100), 2)
                    ImpTot = Round2(DBLet(RS!Kilos, "N") * PreSocio, 2) + ImpBonif
                
                    Bonifica = Bonifica + ImpBonif
                    baseimpo = baseimpo + ImpTot
                
'                    Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * PreSocio, 2)
'                    BaseImpo = BaseImpo + Round2(DBLet(RS!Kilos, "N") * PreSocio * (1 + (PorcBoni / 100)), 2)
            End Select
        End If
        Set Rs9 = Nothing
        
        vConta = RS!Contador
        vFecIni = DBLet(RS!FechaIni, "F")
        vFecFin = DBLet(RS!FechaFin, "F")
        vTipo = RS!TipoFact
        'hasta aqui
            
        ' insertamos en tmpliquidacion la linea de calidad
        sqlLiquid = sqlLiquid & "(" & vUsu.Codigo & ", " & DBSet(RS!Codsocio, "N") & "," & DBSet(RS!codvarie, "N") & ","
        sqlLiquid = sqlLiquid & DBSet(RS!codcampo, "N") & "," & DBSet(RS!numalbar, "N") & "," & DBSet(RS!codcalid, "N") & "," & DBSet(PorcBoni, "N") & ","
        sqlLiquid = sqlLiquid & DBSet(ImpBonif, "N") & "," & DBSet(ImpTot, "N") & "),"
            
            
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' Metemos las bonificaciones
    If sqlLiquid <> "" Then
        conn.Execute SqlLiq & Mid(sqlLiquid, 1, Len(sqlLiquid) - 1)
    End If
    
    ' ultimo registro si ha entrado
    If HayReg Then
        ' gastos de los albaranes
        Sql4 = "select sum(rhisfruta_gastos.importe) "
        Sql4 = Sql4 & " from rhisfruta_gastos "
        Sql4 = Sql4 & " where rhisfruta_gastos.numalbar = " & DBSet(AlbarAnt, "N")
        
        ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
        
        '[Monica]23/07/2012: si es complementaria no hay gastos
        If Check1(5).Value = 1 Then ' si es complementaria no hay gastos
            ImpoGastos = 0
            ImpoGastosFactura = 0
        Else
            ImpoGastosFactura = ImpoGastosFactura + DevuelveValor(Sql4)
        End If
        
        
        ' anticipos
        Sql4 = "select sum(rfactsoc_variedad.imporvar) "
        Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
        Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
        Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
        Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
        
        Anticipos = DevuelveValor(Sql4)
        
        Bruto = baseimpo - Bonifica
        
        ImpoBonif = Bonifica
        
        baseimpo = baseimpo - Anticipos
        
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
    
    
        If Check1(5).Value = 1 Then
            ImpoAport = 0
        Else
            ImpoAport = DevuelveValor("select importe from raporreparto where codsocio = " & DBSet(SocioAnt, "N") & " and tipoentr = 0")
        End If
    
        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
        TotalFac = TotalFac - ImpoGastos
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
        Sql1 = Sql1 & DBSet(Bruto, "N") & ","
        Sql1 = Sql1 & DBSet(ImpoBonif, "N") & ","
        Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
        Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N")
'02/09/2010
'            Sql1 = Sql1 & "),"
        Sql1 = Sql1 & ","
        Sql1 = Sql1 & DBSet(vConta, "N") & "," & DBSet(vTipo, "N") & "," & DBSet(vFecIni, "F") & "," & DBSet(vFecFin, "F") & "),"
        
            
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
        
        BaseImpoFactura = BaseImpoFactura + baseimpo
        ImpoIvaFactura = Round2((BaseImpoFactura) * ComprobarCero(vPorcIva) / 100, 2)
    
        Select Case TipoIRPF
            Case 0
                ImpoRetenFactura = Round2((BaseImpoFactura + ImpoIvaFactura) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoRetenFactura = Round2(BaseImpoFactura * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoRetenFactura = 0
                PorcReten = 0
        End Select
    
        If Check1(5).Value = 1 Then
            ImpoAporFactura = 0
        Else
            ImpoAporFactura = DevuelveValor("select importe from raporreparto where codsocio = " & DBSet(SocioAnt, "N") & " and tipoentr = 0")
        End If
        
        '[Monica]15/04/2013: si hay importe de facturas varias a descontar del socio
        ImpoFrasVarias = 0
        If Check1(14).Value = 1 Then                                                                                          'liquidacion          no descontada    que no sea vtacampo
           ImpoFrasVarias = DevuelveValor("select sum(totalfac) from fvarcabfact where codsocio = " & DBSet(SocioAnt, "N") & " and enliquidacion = 1 and intliqui = 0 and envtacampo = 0")
        End If
        
        ImpoTotalFactura = BaseImpoFactura + ImpoIvaFactura - ImpoRetenFactura - ImpoAporFactura - ImpoGastosFactura ' - ImpoFrasVarias
        
        SqlFactura = "insert into tmpfactura(codusu,codsocio,baseimpo,imporiva,impreten,impapor,impgastos,totalfac,impfrasvar) values ( "
        SqlFactura = SqlFactura & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(BaseImpoFactura, "N") & ","
        SqlFactura = SqlFactura & DBSet(ImpoIvaFactura, "N") & "," & DBSet(ImpoRetenFactura, "N") & ","
        SqlFactura = SqlFactura & DBSet(ImpoAporFactura, "N") & "," & DBSet(ImpoGastosFactura, "N") & ","
        SqlFactura = SqlFactura & DBSet(ImpoTotalFactura, "N") & "," & DBSet(ImpoFrasVarias, "N") & ")"
        
        conn.Execute SqlFactura
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalLiquidacionPicassent = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalLiquidacionAlzira(ctabla As String, cwhere As String, TipoPrec As Byte) As Boolean
Dim RS As ADODB.Recordset
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
Dim Cad As String
Dim HayReg As Boolean
Dim vGastos As Currency


    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionAlzira = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid,"             '[Monica]28/03/2013: Añadida la condicion del if dentro del sum
    SQL = SQL & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(if(rhisfruta_clasif.kilosnet is null, 0, rhisfruta_clasif.kilosnet)) as kilos "
    SQL = SQL & " FROM  " & ctabla

    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "
    SQL = SQL & " having sum(if(rhisfruta_clasif.kilosnet is null, 0, rhisfruta_clasif.kilosnet)) <> 0 " '[Monica]28/03/2013: no tienen que salir los que son 0
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  bonificacion, gastos,  anticipos, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac
    Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!Codsocio
        VarieAnt = RS!codvarie
        NVarieAnt = RS!nomvarie
        CampoAnt = RS!codcampo
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(RS!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                vPorcIva = ""
                '[Monica]29/04/2011: INTERNAS
                If vSocio.EsFactADVInt Then
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                Else
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not RS.EOF
        If CampoAnt <> RS!codcampo Or VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            ' gastos por campo
'[MONICA] : 19/01/2010 los gastos que se aplican a Alzira no son los 4 imptrans, impacarr, imprecol, imppenal
'           si no los gastos que tenemos en rhisfruta_gastos
'            ' [MONICA] : 08/09/2009 los gastos de transporte son una bonificacion para Valsur
'            '            Se restan del resto de gastos
'            'Sql4 = "select sum(if(isnull(imptrans),0,imptrans)) + sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) as gastos from rhisfruta "
'            Sql4 = "select sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) as gastos from rhisfruta "
'            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
'            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
'            Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
'            Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
'            Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F")
'
'            ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
'
'
'
'            Sql4 = "select sum(if(isnull(imptrans),0,imptrans)) as bonif from rhisfruta "
'            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
'            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
'            Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
'            Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
'            Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F")
'
'            ImpoBonif = ImpoBonif + DevuelveValor(Sql4)
             
            '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
            If TipoPrec <> 3 Then
                 
                Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos from rhisfruta_gastos where numalbar in (select numalbar from rhisfruta "
                Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
                Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
                Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
                Sql4 = Sql4 & " and tipoentr <> 1 and tipoentr <> 3 "
                Select Case Combo1(2).ListIndex
                    Case 0      ' recolectado cooperativa
                        Sql4 = Sql4 & " and rhisfruta.recolect = 0"
                    Case 1      ' recolectado socio
                        Sql4 = Sql4 & " and rhisfruta.recolect = 1"
                End Select
                Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
                Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F") & ")"
                '[Monica]02/04/2013: cambio el devuelvevalor por el ObternerGastosAlbaranes
                ImpoGastos = ImpoGastos + ObtenerGastosAlbaranes(CStr(SocioAnt), CStr(VarieAnt), CStr(CampoAnt), ctabla, cwhere, 1)
                'ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
                
            End If
            ImpoBonif = 0
            
            CampoAnt = RS!codcampo
        End If
    
        ' 23/07/2009: añadido el or con la segunda condicion
        If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
            
            Anticipos = DevuelveValor(Sql4)
            
            Bruto = baseimpo
            ImpoGastos = ImpoGastos
            
            ImpoBonif = ImpoBonif
            
'[MONICA] : 08/09/2009 he quitado lo de David pq los gastos de transporte los he quitado arriba
'           dejo lo original
'
'            'FALTA###
'            'DAVID###   20 Agosto 2009
'            'Si es para valsur los gastos se le suman, NO se le restan
'            ' Habria que ver:
'            '   -Los gastos del campo(el punto de arriba)
'            '   -Si en esta funcion solo entra valsur no haria falta poner vParamAplic.Cooperativa = 1
'
'            If vParamAplic.Cooperativa = 1 Then
'                baseimpo = baseimpo + ImpoGastos - Anticipos  'valsur
'            Else
'                baseimpo = baseimpo - ImpoGastos - Anticipos  'original
'            End If
            
            'DAVID###
            'El gasto de la cooperativa siempre se lo quito al total
'            baseimpo = baseimpo - Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)

            '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
            If TipoPrec <> 3 Then
    
    '[Monica] 09/09/2009: el gasto de la cooperativa lo añado a la columna de gastos que no usa Valsur
                ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
            End If
            
            baseimpo = baseimpo + ImpoBonif - ImpoGastos - Anticipos
            
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
        
            ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoBonif, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
'            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
            
            VarieAnt = RS!codvarie
            
            baseimpo = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
            
            ImpoGastos = 0
            ImpoBonif = 0
            Anticipos = 0
            
        End If
        
        If RS!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    '[Monica]29/04/2011: INTERNAS
                    If vSocio.EsFactADVInt Then
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        Recolect = DBLet(RS!Recolect, "N")
        Select Case Recolect
'            Case 0
'                baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * RS!precoop, 2)
'            Case 1
'                baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * RS!presocio, 2)
            Case 0
                baseimpo = baseimpo + (DBLet(RS!Kilos, "N") * RS!PreCoop)
            Case 1
                baseimpo = baseimpo + (DBLet(RS!Kilos, "N") * RS!PreSocio)
        End Select
            
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        Bruto = baseimpo
        
        
        ' anticipos
        Sql4 = "select sum(rfactsoc_variedad.imporvar) "
        Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
        Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
        Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
        Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
        Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
        
        Anticipos = DevuelveValor(Sql4)
            
'[MONICA] 19/01/2010 Los gastos de campo se calculan de la rhisfruta_gastos
'        ' gastos por campo
'        Sql4 = "select  sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) as gastos from rhisfruta "
'        Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
'        Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
'        Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
'        Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
'        Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F")
'
'        ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
'
'        Sql4 = "select sum(if(isnull(imptrans),0,imptrans)) as bonif from rhisfruta "
'        Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
'        Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
'        Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
'        Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
'        Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F")
'
'
'        ImpoBonif = ImpoBonif + DevuelveValor(Sql4)


        '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
        If TipoPrec <> 3 Then

            Sql4 = "select  sum(if(isnull(importe),0,importe)) as gastos from rhisfruta_gastos where numalbar in ( select numalbar from rhisfruta "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
            Sql4 = Sql4 & " and fecalbar >= " & DBSet(txtcodigo(6).Text, "F")
            Sql4 = Sql4 & " and fecalbar <= " & DBSet(txtcodigo(7).Text, "F") & ")"
                
            '[Monica]02/04/2013: cambio el devuelvevalor por el ObternerGastosAlbaranes
            ImpoGastos = ImpoGastos + ObtenerGastosAlbaranes(CStr(SocioAnt), CStr(VarieAnt), CStr(CampoAnt), ctabla, cwhere, 1)
'           ImpoGastos = ImpoGastos + DevuelveValor(Sql4)

        End If

'        If vParamAplic.Cooperativa = 1 Then
'            baseimpo = baseimpo + ImpoGastos - Anticipos  'valsur
'        Else
'            baseimpo = baseimpo - ImpoGastos - Anticipos  'original
'        End If
        
        'DAVID###
        'El gasto de la cooperativa siempre se lo quito al total
'        baseimpo = baseimpo - Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
        
        '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
        If TipoPrec <> 3 Then
                
        '[Monica] 09/09/2009: el gasto de la cooperativa lo añado a la columna de gastos que no usa Valsur
                ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
        
        End If
        
        baseimpo = baseimpo + ImpoBonif - ImpoGastos - Anticipos
        
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
    
        ImpoAport = Round2((Bruto + ImpoBonif - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
    
        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
        
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
        Sql1 = Sql1 & DBSet(Bruto, "N") & ","
        Sql1 = Sql1 & DBSet(ImpoBonif, "N") & ","
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
    
    CargarTemporalLiquidacionAlzira = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalTotalesCatadau()
Dim SQL As String

    SQL = "delete from tmpinformes2 where codusu = " & vUsu.Codigo
    conn.Execute SQL
    'comerciales
    SQL = "insert into tmpinformes2 (codusu,codigo1,importe1) "
    SQL = SQL & "select " & vUsu.Codigo & ", codvarie, sum(coalesce(kilosnet,0)) from tmpliquidacion where codusu = " & DBSet(vUsu.Codigo, "N")
    SQL = SQL & " and tipoentr in (0,2)"
    SQL = SQL & " group by 1, 2 "
    conn.Execute SQL
    'venta campo
    SQL = "insert into tmpinformes2 (codusu,codigo1,importe2) "
    SQL = SQL & "select " & vUsu.Codigo & ", codvarie, sum(coalesce(kilosnet,0)) from tmpliquidacion where codusu = " & DBSet(vUsu.Codigo, "N")
    SQL = SQL & " and tipoentr = 1 "
    SQL = SQL & " group by 1, 2 "
    conn.Execute SQL
    'kilos aportacion
    SQL = "insert into tmpinformes2 (codusu,codigo1,importe3) "
    SQL = SQL & "select " & vUsu.Codigo & ", codvarie, sum(coalesce(kilosnet,0)) from tmpliquidacion where codusu = " & DBSet(vUsu.Codigo, "N")
    SQL = SQL & " and codcalid = 0"
    SQL = SQL & " group by 1, 2 "
    conn.Execute SQL


End Function


Private Function CargarTemporalLiquidacionCatadau(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
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
Dim Cad As String
Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionCatadau = False

    If CargarTemporalCatadau(ctabla, cwhere, 1) Then
        '[Monica]24/04/2013: pq en la anterior funcion se graba la tmpinformes
        Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
        conn.Execute Sql2

    
        '[Monica]27/01/2016: si es catadau y es complementaria sacamos otro report
        If vParamAplic.Cooperativa = 0 And Check1(5).Value = 1 Then
            CargarTemporalTotalesCatadau
        End If
    
    
    
        SQL = "SELECT tmpliquidacion.codsocio, tmpliquidacion.codvarie, variedades.nomvarie,"
        SQL = SQL & " sum(tmpliquidacion.kilosnet) as kilos , sum(tmpliquidacion.importe) as importe "
        SQL = SQL & " FROM tmpliquidacion, variedades where codusu = " & vUsu.Codigo
        SQL = SQL & " and tmpliquidacion.codvarie = variedades.codvarie "
        SQL = SQL & " group by 1, 2, 3 "
        SQL = SQL & " order by 1, 2, 3 "
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                                        'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  gastos,    anticipos, baseimpo, porceiva, imporiva,
        Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                       'porcerete, imporret, totalfac
        Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
        
        Set vSeccion = New CSeccion
        
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If Not vSeccion.AbrirConta Then
                Exit Function
            End If
        End If
        
        HayReg = False
        If Not RS.EOF Then
            SocioAnt = RS!Codsocio
            VarieAnt = RS!codvarie
            NVarieAnt = RS!nomvarie
            
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    '[Monica]05/03/2014:
                    If vParamAplic.Cooperativa = 4 Then
                        '[Monica]29/04/2011: INTERNAS
                        If vSocio.EsFactADVInt Then
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                        Else
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                    
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
                TipoIRPF = vSocio.TipoIRPF
            End If
        End If
        
        While Not RS.EOF
        
            ' 23/07/2009: añadido el or con la segunda condicion
            If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
                ' anticipos
                Sql4 = "select sum(rfactsoc_variedad.imporvar) "
                Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
                Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
                Sql4 = Sql4 & " and rfactsoc.esanticipogasto = 0 "
                Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
                Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
                
                If vParamAplic.Cooperativa = 9 Then
                '[Monica]11/12/2013: en el caso de natural solo decontamos los anticipos de las fechas que me hayan dicho
                ' si no seleccionamos ninguna no descontaremos ningun anticipo
                    If vFechas <> "" Then
                        Sql4 = Sql4 & " and rfactsoc.fecfactu in (" & vFechas & ")"
                    Else
                        Sql4 = Sql4 & " and rfactsoc.fecfactu = '1900-01-01' "
                    End If
                End If
                
                
                Anticipos = DevuelveValor(Sql4)
                
                Bruto = baseimpo
                
                ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
                
                baseimpo = baseimpo - ImpoGastos
                
                '[Monica]10/03/2014: esto solo seria para el caso de alzira
                '                    si no permitimos facturas negativas el valor de anticipos es mayor que la base imponible
                If Check1(21).Value = 1 And baseimpo < Anticipos Then
                    ' si no queremos que sea negativa no descuento los anticipos
                Else
                    baseimpo = baseimpo - Anticipos
                End If
                
                ImpoIva = Round2((baseimpo) * ImporteSinFormato(vPorcIva) / 100, 2)
            
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
            
                ImpoAport = Round2((Bruto - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
            
                TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
                
                Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
                Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
                Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
                Sql1 = Sql1 & DBSet(Bruto, "N") & ","
                Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
                Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
                Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
                Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
                Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
                Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
                Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
                
                VarieAnt = RS!codvarie
                
                baseimpo = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                KilosNet = 0
                
                ImpoGastos = 0
                Anticipos = 0
                
            End If
            
            If RS!Codsocio <> SocioAnt Then
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(RS!Codsocio) Then
                    If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        vPorcGasto = ""
                        vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                    End If
                    NSocioAnt = vSocio.Nombre
                End If
                SocioAnt = vSocio.Codigo
                TipoIRPF = vSocio.TipoIRPF
            End If
            
            KilosNet = DBLet(RS!Kilos, "N")
            
            baseimpo = DBLet(RS!Importe, "N")
                
            ' gastos
            Sql4 = "select sum(gastos) from tmpliquidacion1 "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codusu = " & vUsu.Codigo
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            
            ImpoGastos = DevuelveValor(Sql4)
                
            HayReg = True
            
            RS.MoveNext
        Wend
            
        ' ultimo registro si ha entrado
        If HayReg Then
            
            ' [Monica] 16/03/2010
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc.esanticipogasto = 0 "
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
            
            '[Monica]23/11/2012: en el caso de ser Natural solo descontamos los anticipos de las fechas seleccionadas
            '                    si no seleccionamos ninguna no descontaremos ningun anticipo
            If vParamAplic.Cooperativa = 9 Then
                    If vFechas <> "" Then
                        Sql4 = Sql4 & " and rfactsoc.fecfactu in (" & vFechas & ")"
                    Else
                        Sql4 = Sql4 & " and rfactsoc.fecfactu = '1900-01-01' "
                    End If
            End If
            
            Anticipos = DevuelveValor(Sql4)
            ' [Monica] 16/03/2010
            
            Bruto = baseimpo
            
            ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
            
            baseimpo = baseimpo - ImpoGastos
            
            '[Monica]10/03/2014: esto solo seria para el caso de alzira
            '                    si no permitimos facturas negativas el valor de anticipos es mayor que la base imponible
            If Check1(21).Value = 1 And baseimpo < Anticipos Then
                ' si no queremos que sea negativa no descuento los anticipos
            Else
                baseimpo = baseimpo - Anticipos
            End If
            
            ImpoIva = Round2((baseimpo) * ImporteSinFormato(vPorcIva) / 100, 2)
            
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
        
            ImpoAport = Round2((Bruto - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
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
        
        CargarTemporalLiquidacionCatadau = True
        Exit Function
    End If
        
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function ActualizarRegistrosFac(ctabla As String, cwhere As String) As Boolean
'Actualizar la marca de impreso
Dim SQL As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistrosFac = False
    SQL = "update " & ctabla & ", usuarios.stipom set impreso = 1 "
    SQL = SQL & " where usuarios.stipom.codtipom = rfactsoc.codtipom "
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " and " & cwhere
    End If
    
    conn.Execute SQL
    
    ActualizarRegistrosFac = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function



Private Function CargarTemporalAnticiposGastos(ctabla As String, cwhere As String, Cad As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim baseimpo As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim KilosGastos As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency
    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim HayReg As Boolean

Dim Sql3 As String
Dim Importe As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposGastos = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,  "
    SQL = SQL & "rcalidad.gastosrec, " ' sum(rhisfruta.imprecol) as importe, "
    SQL = SQL & "sum(rhisfruta_clasif.kilosnet) as kilos"
    SQL = SQL & " FROM  " & ctabla

    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5 "
    SQL = SQL & " order by 1, 2, 3, 4, 5 "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie,kggastos, baseimpo, porceiva, imporiva,
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                   'porcerete, imporret, totalfac, kgneto
    Sql2 = Sql2 & " porcen2, importeb1, importeb2, importeb3) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!Codsocio
        VarieAnt = RS!codvarie
        NVarieAnt = RS!nomvarie
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(RS!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        End If
    End If
    
    While Not RS.EOF
        If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            
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
            Sql1 = Sql1 & DBSet(KilosGastos, "N") & "," & DBSet(baseimpo, "N") & ","
            Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(TotalFac, "N") & "," & DBSet(KilosNet, "N") & "),"
            
            VarieAnt = RS!codvarie
            
            baseimpo = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
            KilosGastos = 0
        End If
        
        If RS!Codsocio <> SocioAnt Then
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        If DBLet(RS!gastosrec, "N") = 1 Then
            KilosGastos = KilosGastos + DBLet(RS!Kilos, "N")
        
        
            ' insertar linea de variedad, campo
            Sql3 = "select sum(imprecol) from rhisfruta where "
            If Cad <> "" Then Sql3 = Sql3 & Cad & " and "
            Sql3 = Sql3 & " rhisfruta.codvarie = " & DBSet(RS!codvarie, "N")
            Sql3 = Sql3 & " and codcampo = " & DBSet(RS!codcampo, "N") & " and codsocio = " & DBSet(RS!Codsocio, "N")
            
            Importe = DevuelveValor(Sql3)
        
        
        
        
            baseimpo = baseimpo + Importe  '+ DBLet(rs!Importe, "N")
        End If
        
            
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
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
        Sql1 = Sql1 & DBSet(KilosGastos, "N") & "," & DBSet(baseimpo, "N") & ","
        Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(TotalFac, "N") & "," & DBSet(KilosNet, "N") & "),"
    
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalAnticiposGastos = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalAnticiposGenericos(ctabla As String, cwhere As String, ConCampo As Boolean, DeRetirada As Boolean) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim baseimpo As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim KilosGastos As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency
    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim HayReg As Boolean

Dim Sql3 As String
Dim Importe As Currency
Dim Precio As Currency


Dim BaseImpoFactura As Currency
Dim ImpoIvaFactura As Currency
Dim ImpoRetenFactura As Currency
Dim ImpoTotalFactura As Currency
Dim SqlFactura As String
Dim SqlTempo As String
Dim Sql8 As String

    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposGenericos = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql2 = "delete from tmpfactura where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ' en la tabla intermedia tenemos tanto los registros de rclasifica como los de rhisfruta
    If ConCampo Then
        SQL = "SELECT codsocio, codvarie, nomvarie, codcampo, "
        SQL = SQL & "sum(kilosnet) as kilos"
        SQL = SQL & " FROM  tmpliquidacion "
        SQL = SQL & " WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " group by 1, 2, 3, 4 "
        SQL = SQL & " order by 1, 2, 3, 4 "
    Else
        SQL = "SELECT codsocio, codvarie, nomvarie, "
        SQL = SQL & "sum(kilosnet) as kilos"
        SQL = SQL & " FROM  tmpliquidacion "
        SQL = SQL & " WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " group by 1, 2, 3 "
        SQL = SQL & " order by 1, 2, 3 "
    End If


    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu, codsocio, nomsocio, codvarie, nomvarie, kilos, porceiva, porcerete,precio
    Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, porcen1, porcen2, precio1) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
            
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        SocioAnt = RS!Codsocio
        VarieAnt = RS!codvarie
        NVarieAnt = RS!nomvarie
        
        Set vSocio = Nothing
        Set vSocio = New cSocio
        If vSocio.LeerDatos(RS!Codsocio) Then
            If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
            End If
            NSocioAnt = vSocio.Nombre
            TipoIRPF = vSocio.TipoIRPF
        
            BaseImpoFactura = 0
            ImpoIvaFactura = 0
            ImpoRetenFactura = 0
            ImpoTotalFactura = 0
        
        End If
    End If
    
    Sql1 = ""
    KilosNet = 0
    While Not RS.EOF
        If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            Sql8 = "select precioindustria from rprecios where (codvarie, tipofact, contador) = ("
            Sql8 = Sql8 & "SELECT codvarie, tipofact, max(contador) FROM rprecios WHERE codvarie=" & DBSet(VarieAnt, "N") & " and "
            If DeRetirada Then
                Sql8 = Sql8 & " tipofact = 5 and fechaini = " & DBSet(txtcodigo(6).Text, "F")
            Else
                Sql8 = Sql8 & " tipofact = 4 and fechaini = " & DBSet(txtcodigo(6).Text, "F")
            End If
            Sql8 = Sql8 & " and fechafin = " & DBSet(txtcodigo(7).Text, "F") & " and precioindustria <> 0 and precioindustria is not null "
            Sql8 = Sql8 & " group by 1, 2) "
            
            Precio = DevuelveValor(Sql8)
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(vPorcIva, "N") & ","
            Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
            Sql1 = Sql1 & DBSet(Precio, "N") & "),"
            
            VarieAnt = RS!codvarie
            
            KilosNet = 0
        End If
        
        If RS!Codsocio <> SocioAnt Then
            
            ImpoIvaFactura = Round2((BaseImpoFactura) * ComprobarCero(vPorcIva) / 100, 2)
        
            Select Case TipoIRPF
                Case 0
                    ImpoRetenFactura = Round2((BaseImpoFactura + ImpoIvaFactura) * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoRetenFactura = Round2(BaseImpoFactura * vParamAplic.PorcreteFacSoc / 100, 2)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoRetenFactura = 0
                    PorcReten = 0
            End Select
        
            ImpoTotalFactura = BaseImpoFactura + ImpoIvaFactura - ImpoRetenFactura
            
            SqlFactura = "insert into tmpfactura(codusu,codsocio,baseimpo,imporiva,impreten,impapor,impgastos,totalfac) values ( "
            SqlFactura = SqlFactura & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(BaseImpoFactura, "N") & ","
            SqlFactura = SqlFactura & DBSet(ImpoIvaFactura, "N") & "," & DBSet(ImpoRetenFactura, "N") & ","
            SqlFactura = SqlFactura & "0,0,"
            SqlFactura = SqlFactura & DBSet(ImpoTotalFactura, "N") & ")"
            
            conn.Execute SqlFactura
            
            BaseImpoFactura = 0
            ImpoIvaFactura = 0
            ImpoRetenFactura = 0
            ImpoTotalFactura = 0
            
            
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                NSocioAnt = vSocio.Nombre
            End If
            
            SocioAnt = vSocio.Codigo
            TipoIRPF = vSocio.TipoIRPF
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        Sql8 = "select precioindustria from rprecios where (codvarie, tipofact, contador) = ("
        Sql8 = Sql8 & "SELECT codvarie, tipofact, max(contador) FROM rprecios WHERE codvarie=" & DBSet(VarieAnt, "N") & " and "
        If DeRetirada Then
            Sql8 = Sql8 & " tipofact = 5 and fechaini = " & DBSet(txtcodigo(6).Text, "F")
        Else
            Sql8 = Sql8 & " tipofact = 4 and fechaini = " & DBSet(txtcodigo(6).Text, "F")
        End If
        Sql8 = Sql8 & " and fechafin = " & DBSet(txtcodigo(7).Text, "F") & " and precioindustria <> 0 and precioindustria is not null "
        Sql8 = Sql8 & " group by 1, 2) "
        
        Precio = DevuelveValor(Sql8)
        
        Importe = Round2(RS!Kilos * Precio, 2)
    
        BaseImpoFactura = BaseImpoFactura + Importe
        
            
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
        Sql1 = Sql1 & DBSet(vPorcIva, "N") & ","
        Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
        Sql1 = Sql1 & DBSet(Precio, "N") & "),"
        
        ImpoIvaFactura = Round2((BaseImpoFactura) * ComprobarCero(vPorcIva) / 100, 2)
    
        Select Case TipoIRPF
            Case 0
                ImpoRetenFactura = Round2((BaseImpoFactura + ImpoIvaFactura) * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoRetenFactura = Round2(BaseImpoFactura * vParamAplic.PorcreteFacSoc / 100, 2)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoRetenFactura = 0
                PorcReten = 0
        End Select
    
        ImpoTotalFactura = BaseImpoFactura + ImpoIvaFactura - ImpoRetenFactura
        
        SqlFactura = "insert into tmpfactura(codusu,codsocio,baseimpo,imporiva,impreten,impapor,impgastos,totalfac) values ( "
        SqlFactura = SqlFactura & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(BaseImpoFactura, "N") & ","
        SqlFactura = SqlFactura & DBSet(ImpoIvaFactura, "N") & "," & DBSet(ImpoRetenFactura, "N") & ","
        SqlFactura = SqlFactura & "0,0,"
        SqlFactura = SqlFactura & DBSet(ImpoTotalFactura, "N") & ")"
        
        conn.Execute SqlFactura

        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
        
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalAnticiposGenericos = True
    Exit Function
    
eCargarTemporal:
    MuestraError "Cargando temporal Anticipos Genéricos/Retirada", Err.Description
End Function



Private Function CargarTemporalCatadau(ctabla As String, cwhere As String, Tipo As Byte) As Boolean
'tipo  0=anticipos
'      1=liquidacion
Dim RS As ADODB.Recordset
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
Dim Cad As String
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

Dim HayPrecio As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalCatadau = False

    Sql2 = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    Sql2 = "delete from tmpliquidacion1 where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql2 = "delete from tmpinformes2 where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    '[Monica]24/04/2013: Meto los albaranes y calidades que puede que liquide
    Sql2 = "insert into tmpinformes2 (codusu, importe1, fecha1, importe2, importe3, importe4) select " & vUsu.Codigo & ",rhisfruta.numalbar, rhisfruta.fecalbar,rhisfruta.codvarie, rhisfruta_clasif.codcalid, "
    Sql2 = Sql2 & " sum(rhisfruta_clasif.kilosnet) as kilos  "
    Sql2 = Sql2 & " FROM  (" & ctabla & ") inner join tmpvarie on rhisfruta.codvarie = tmpvarie.codvarie "
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql2 = Sql2 & " WHERE " & cwhere
    End If
    Sql2 = Sql2 & " group by 1, 2, 3, 4, 5"
    Sql2 = Sql2 & " having sum(rhisfruta_clasif.kilosnet) <> 0 "
    Sql2 = Sql2 & " order by 1, 2, 3, 4, 5 "
    
    conn.Execute Sql2
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, "
    SQL = SQL & " rhisfruta.recolect, rhisfruta.tipoentr, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, "
    SQL = SQL & " sum(rhisfruta_clasif.kilosnet) as kilos "
    SQL = SQL & " FROM  (" & ctabla & ") inner join tmpvarie on rhisfruta.codvarie = tmpvarie.codvarie "
    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8 "
    SQL = SQL & " having sum(rhisfruta_clasif.kilosnet) <> 0 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8 "

    Nregs = TotalRegistrosConsulta(SQL)
    
    Label2(10).Caption = "Cargando Tabla Temporal"
    Me.Pb1.visible = True
    Me.Pb1.Max = Nregs
    Me.Pb1.Value = 0
    Me.Refresh

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    '[Monica] 14/12/2009 si es liquidacion complementaria: no se descuentan gastos(complementaria) ponemos gastos = 0
    If Tipo = 1 And Check1(5).Value = 1 Then Tipo = 3 'seleccionamos los precios de liquidacion complementaria
                                    
                                    
    While Not RS.EOF
        Label2(12).Caption = "Socio " & RS!Codsocio & " Variedad " & RS!codvarie & "-" & RS!codcalid & " Campo " & RS!codcampo
        IncrementarProgresNew Pb1, 1
        Me.Refresh
        DoEvents
    
        Sql3 = "select fechaini, fechafin, precioindustria, max(contador) as contador from rprecios where codvarie = " & DBSet(RS!codvarie, "N")
        Sql3 = Sql3 & " and tipofact = " & DBSet(Tipo, "N")
        Sql3 = Sql3 & " and fechaini <= " & DBSet(RS!Fecalbar, "F")
        Sql3 = Sql3 & " and fechafin >= " & DBSet(RS!Fecalbar, "F")
        Sql3 = Sql3 & " group by 1,2,3"
                
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS1.EOF Then
            Contador = DBLet(RS1!Contador, "N")
            FechaIni = DBLet(RS1!FechaIni, "F")
            FechaFin = DBLet(RS1!FechaFin, "F")
        End If
        Set RS1 = Nothing
        
        If DBLet(RS!Recolect, "N") = 0 Then 'cooperativa
            Sql3 = "select precoop "
        Else
            Sql3 = "select presocio "
        End If
        
        Sql3 = Sql3 & " from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql3 = Sql3 & " and codcalid = " & DBSet(RS!codcalid, "N")
        Sql3 = Sql3 & " and tipofact = " & DBSet(Tipo, "N")
        Sql3 = Sql3 & " and contador = " & DBSet(Contador, "N")
        
        Precio = DevuelveValor(Sql3)
    
    
        '[monica]24/04/2013: miro si hay que liquidar
        HayPrecio = (TotalRegistrosConsulta(Sql3) <> 0)
        If Not HayPrecio Then
        
            Sql4 = "delete from tmpinformes2 where codusu = " & DBSet(vUsu.Codigo, "N") & " and importe2 = " & DBSet(RS!codvarie, "N")
            Sql4 = Sql4 & " and importe3 = " & DBSet(RS!codcalid, "N")
            Sql4 = Sql4 & " and fecha1 between " & DBSet(FechaIni, "F") & " and " & DBSet(FechaFin, "F")

            conn.Execute Sql4
        Else
            Dim vPrecio As Currency
            vPrecio = Precio
        
            Sql4 = "update tmpinformes2 set precio1 = " & DBSet(Precio, "N")
            Sql4 = Sql4 & " where codusu = " & DBSet(vUsu.Codigo, "N")
            Sql4 = Sql4 & " and importe2 = " & DBSet(RS!codvarie, "N")
            Sql4 = Sql4 & " and importe3 = " & DBSet(RS!codcalid, "N")
            Sql4 = Sql4 & " and fecha1 between " & DBSet(FechaIni, "F") & " and " & DBSet(FechaFin, "F")

            conn.Execute Sql4
        End If
        
        Sql4 = "select sum(if(isnull(imptrans),0,imptrans)) + sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) as gastos"
        Sql4 = Sql4 & "  from rhisfruta "
        Sql4 = Sql4 & " where rhisfruta.codsocio = " & DBSet(RS!Codsocio, "N") & "  and "
        Sql4 = Sql4 & " rhisfruta.codvarie = " & DBSet(RS!codvarie, "N") & "  and "
        Sql4 = Sql4 & " rhisfruta.codcampo = " & DBSet(RS!codcampo, "N") & " and "
        Sql4 = Sql4 & " rhisfruta.fecalbar >= " & DBSet(FechaIni, "F") & " and "
        Sql4 = Sql4 & " rhisfruta.fecalbar <= " & DBSet(FechaFin, "F") & " and "
        '[Monica]24/04/2013: tipo de entrada la seleccionada
        Sql4 = Sql4 & " rhisfruta.numalbar in (select distinct importe1 from tmpinformes2 where codusu = " & vUsu.Codigo & ")"
        Sql4 = Sql4 & " and rhisfruta.tipoentr <> 1 and rhisfruta.tipoentr <> 3 and rhisfruta.tipoentr <> 4 and rhisfruta.tipoentr <> 6 "
        
        '[Monica]03/06/2013: distinguimos entre entradas normales y entradas de p.integrado (solo para catadau)
        If vParamAplic.Cooperativa = 0 And (OpcionListado = 12 Or OpcionListado = 13 Or OpcionListado = 14) Then
            '[Monica]01/02/2016: antes Check1(16)
            If Check1(23).Value = 1 Then ' solo entradas normales
                Sql4 = Sql4 & " and rhisfruta.tipoentr = 0"
            End If
            If Check1(24).Value = 1 Then
            'Else ' solo producto integrado
                Sql4 = Sql4 & " and rhisfruta.tipoentr = 2"
            End If
        End If
        
        Gastos = DevuelveValor(Sql4)
        
        '[Monica]05/03/2014: tramos para alzira
        If vParamAplic.Cooperativa = 4 Then
            Gastos = ObtenerGastosAlbaranes(CStr(RS!Codsocio), CStr(RS!codvarie), CStr(RS!codcampo), ctabla, cwhere, 1)
        End If
        
        '[Monica] 03/12/2009 si es liquidacion: no se descuentan gastos(complementaria) ponemos gastos = 0
        If Tipo = 3 Or Tipo = 6 Then
            Gastos = 0
        End If
        ' end 03/12/2009
        
        Sql5 = "select count(*) from tmpliquidacion1 where codsocio = " & DBSet(RS!Codsocio, "N") & "  and "
        Sql5 = Sql5 & " tmpliquidacion1.codvarie = " & DBSet(RS!codvarie, "N") & "  and "
        Sql5 = Sql5 & " tmpliquidacion1.codcampo = " & DBSet(RS!codcampo, "N") & " and "
        Sql5 = Sql5 & " tmpliquidacion1.fechaini = " & DBSet(FechaIni, "F") & " and "
        Sql5 = Sql5 & " tmpliquidacion1.fechafin = " & DBSet(FechaFin, "F") & " and "
        Sql5 = Sql5 & " tmpliquidacion1.codusu = " & vUsu.Codigo
        
        If TotalRegistros(Sql5) = 0 Then
            Sql5 = "insert into tmpliquidacion1 values (" & vUsu.Codigo & "," & DBSet(RS!Codsocio, "N") & ","
            Sql5 = Sql5 & DBSet(RS!codvarie, "N") & ","
            Sql5 = Sql5 & DBSet(RS!codcampo, "N") & ","
            Sql5 = Sql5 & DBSet(FechaIni, "F") & ","
            Sql5 = Sql5 & DBSet(FechaFin, "F") & ","
            Sql5 = Sql5 & DBSet(Gastos, "N") & ")"
            
            conn.Execute Sql5
        End If
    

        '[Monica]24/04/2013: añadida la condicion
        If HayPrecio Then
            Sql2 = "select count(*) from tmpliquidacion where codusu = " & vUsu.Codigo
            Sql2 = Sql2 & " and codsocio = " & DBSet(RS!Codsocio, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(RS!codcampo, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and recolect = " & DBSet(RS!Recolect, "N")
            Sql2 = Sql2 & " and tipoentr = " & DBSet(RS!TipoEntr, "N")
            Sql2 = Sql2 & " and codcalid = " & DBSet(RS!codcalid, "N")
            Sql2 = Sql2 & " and contador = " & DBSet(Contador, "N")
            Sql2 = Sql2 & " and fechaini = " & DBSet(FechaIni, "F")
            Sql2 = Sql2 & " and fechafin = " & DBSet(FechaFin, "F")
            
            If TotalRegistros(Sql2) = 0 Then
                Kilos = 0
                
                Sql3 = "insert into tmpliquidacion (codusu,codsocio,codcampo,recolect,tipoentr,codvarie,codcalid,contador,kilosnet,precio,importe, "
                Sql3 = Sql3 & " nomvarie, fechaini, fechafin, gastos)"
                Sql3 = Sql3 & " values (" & vUsu.Codigo & "," & DBSet(RS!Codsocio, "N") & "," & DBSet(RS!codcampo, "N") & ","
                Sql3 = Sql3 & DBSet(RS!Recolect, "N") & "," & DBSet(RS!TipoEntr, "N") & ","
                Sql3 = Sql3 & DBSet(RS!codvarie, "N") & "," & DBSet(RS!codcalid, "N") & "," & DBSet(Contador, "N") & ","
                Sql3 = Sql3 & DBSet(RS!Kilos, "N") & "," & DBSet(Precio, "N") & "," & DBSet(0, "N") & ","
                Sql3 = Sql3 & DBSet(RS!nomvarie, "T") & "," & DBSet(FechaIni, "F") & ","
                Sql3 = Sql3 & DBSet(FechaFin, "F") & "," & DBSet(Gastos, "N") & ")"
                
                conn.Execute Sql3
                Kilos = RS!Kilos
            Else
                Kilos = Kilos + RS!Kilos
                Sql3 = "update tmpliquidacion set kilosnet = kilosnet + " & DBSet(RS!Kilos, "N")
                Sql3 = Sql3 & " where codusu = " & vUsu.Codigo
                Sql3 = Sql3 & " and codsocio = " & DBSet(RS!Codsocio, "N")
                Sql3 = Sql3 & " and codcampo = " & DBSet(RS!codcampo, "N")
                Sql3 = Sql3 & " and recolect = " & DBSet(RS!Recolect, "N")
                Sql3 = Sql3 & " and tipoentr = " & DBSet(RS!TipoEntr, "N")
                Sql3 = Sql3 & " and codvarie = " & DBSet(RS!codvarie, "N")
                Sql3 = Sql3 & " and codcalid = " & DBSet(RS!codcalid, "N")
                Sql3 = Sql3 & " and contador = " & DBSet(Contador, "N")
                Sql3 = Sql3 & " and fechaini = " & DBSet(FechaIni, "F")
                Sql3 = Sql3 & " and fechafin = " & DBSet(FechaFin, "F")
                
                conn.Execute Sql3
            End If
        '[Monica]24/04/2013: añadida la condicion he puesto end if
        End If
        
        RS.MoveNext
    Wend
                                    
'[Monica]27/01/2016: si es complementaria y se añaden lo de agroseguro
    If Check1(26).Value = 1 Then
        SQL = "insert into tmpliquidacion (codusu,codcampo,codsocio,codvarie,kilosnet,precio) "
        SQL = SQL & " select " & vUsu.Codigo & ",rcampos.codcampo,rcampos.codsocio,rcampos.codvarie,sum(coalesce(kilosaportacion,0))," & DBSet(vPrecio, "N")
        SQL = SQL & " from rcampos_seguros inner  join rcampos on rcampos_seguros.codcampo = rcampos.codcampo "
        SQL = SQL & " where rcampos.codvarie in (select distinct codvarie from tmpliquidacion where codusu = " & vUsu.Codigo & ") "
        SQL = SQL & " group by 1, 2, 3, 4 "
        conn.Execute SQL
        
'        SQL = "update tmpliquidacion  dd, tmpliquidacion ff "
'        SQL = SQL & " set dd.codsocio = ff.codsocio, dd.codvarie = ff.codvarie "
'        SQL = SQL & " where dd.codusu = " & DBSet(vUsu.Codigo, "N") & " and dd.codsocio = 0 "
'        SQL = SQL & " and dd.codusu = ff.codusu and dd.codcampo = ff.codcampo  "
'        conn.Execute SQL
    End If
                                    
                                    
    Sql3 = "update tmpliquidacion set importe = round(kilosnet * precio,2) where codusu = " & vUsu.Codigo
    conn.Execute Sql3
    
    
    Sql3 = "update tmpinformes2 set importe5 = round(importe4 * precio1,2) where codusu = " & vUsu.Codigo
    conn.Execute Sql3
    
    
    'guardamos los gastos
    Sql3 = "update tmpinformes2, rhisfruta  set importeb1 = (if(isnull(imptrans),0,imptrans) + if(isnull(impacarr),0,impacarr) + if(isnull(imprecol),0,imprecol) + if(isnull(imppenal),0,imppenal)) "
    Sql3 = Sql3 & " where codusu = " & vUsu.Codigo
    Sql3 = Sql3 & " and tmpinformes2.importe1 = rhisfruta.numalbar "
    conn.Execute Sql3
    
    
    Sql3 = "delete from tmpliquidacion1 where not (codusu, codsocio, codvarie, codcampo) in (select " & vUsu.Codigo & ", codsocio, codvarie, codcampo from tmpliquidacion where codusu = " & vUsu.Codigo & ") "
    
    conn.Execute Sql3
    
    
                                    
    Me.Pb1.visible = False
    Me.Label2(10).Caption = ""
    Me.Label2(12).Caption = ""
    Me.Refresh
                                    
    CargarTemporalCatadau = True
    Exit Function
    
eCargarTemporal:
    Me.Pb1.visible = False
    Me.Label2(10).Caption = ""
    Me.Label2(12).Caption = ""
    Me.Refresh
    
    MuestraError "Cargando temporal", Err.Description
End Function



Private Function CargarTemporalLiquidacionQuatretonda(ctabla As String, cwhere As String, Seccion As String) As Boolean
Dim RS As ADODB.Recordset
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
Dim Cad As String

Dim KilosRet As Currency
Dim ImporRet As Currency

Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionQuatretonda = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    If CargarTemporalQuatretonda(ctabla, cwhere, 1) Then
        SQL = "SELECT tmpliquidacion.codsocio, tmpliquidacion.codvarie, variedades.nomvarie,"
        SQL = SQL & " sum(tmpliquidacion.kilosnet) as kilos , sum(tmpliquidacion.importe) as importe "
        SQL = SQL & " FROM tmpliquidacion, variedades where codusu = " & vUsu.Codigo
        SQL = SQL & " and tmpliquidacion.codvarie = variedades.codvarie "
        SQL = SQL & " group by 1, 2, 3 "
        SQL = SQL & " order by 1, 2, 3 "
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                                        'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  retirada,    anticipos, baseimpo, porceiva, imporiva,
        Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                       'porcerete, imporret, totalfac
        Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
        
        Set vSeccion = New CSeccion
'[25/06/2012]: seccion
'        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.LeerDatos(Seccion) Then
            If Not vSeccion.AbrirConta Then
                Exit Function
            End If
        End If
        
        HayReg = False
        If Not RS.EOF Then
            SocioAnt = RS!Codsocio
            VarieAnt = RS!codvarie
            NVarieAnt = RS!nomvarie
            
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                '[25/06/2012]: seccion
'                If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.Seccionhorto) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), Seccion) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
                TipoIRPF = vSocio.TipoIRPF
            End If
        End If
        
        While Not RS.EOF
        
            ' 23/07/2009: añadido el or con la segunda condicion
            If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
                ' anticipos
                Sql4 = "select sum(rfactsoc_variedad.imporvar) "
                Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
                Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
                Sql4 = Sql4 & " and rfactsoc.esanticipogasto = 0 "
                Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
                Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
                Sql4 = Sql4 & " and rfactsoc.esretirada = 0"
                
                Anticipos = DevuelveValor(Sql4)
                
                Bruto = baseimpo
                
                ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
                
                baseimpo = baseimpo - ImpoGastos - Anticipos
                
                ImpoIva = Round2((baseimpo) * ImporteSinFormato(vPorcIva) / 100, 2)
            
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
            
                ImpoAport = Round2((Bruto - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
            
                TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
                
                Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
                Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
                Sql1 = Sql1 & DBSet(KilosNet + KilosRet, "N") & ","
                Sql1 = Sql1 & DBSet(Bruto + ImporRet, "N") & ","
                Sql1 = Sql1 & DBSet(ImporRet, "N") & ","
                Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
                Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
                Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
                Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
                Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
                Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
                
                VarieAnt = RS!codvarie
                
                baseimpo = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                KilosNet = 0
                KilosRet = 0
                ImporRet = 0
                
                ImpoGastos = 0
                Anticipos = 0
                
            End If
            
            If RS!Codsocio <> SocioAnt Then
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(RS!Codsocio) Then
                '[25/06/2012]: seccion
'                    If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), vParamAplic.Seccionhorto) Then
                    If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), Seccion) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        vPorcGasto = ""
                        vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                    End If
                    NSocioAnt = vSocio.Nombre
                End If
                SocioAnt = vSocio.Codigo
                TipoIRPF = vSocio.TipoIRPF
            End If
            
                
            ' kilos retirada
            Sql4 = "select sum(codcampo) from tmpliquidacion1 "
            Sql4 = Sql4 & " where codsocio = " & DBSet(RS!Codsocio, "N")
            Sql4 = Sql4 & " and codusu = " & vUsu.Codigo
            Sql4 = Sql4 & " and codvarie = " & DBSet(RS!codvarie, "N")
            
            ImpoGastos = DevuelveValor(Sql4)
                
            KilosNet = DBLet(RS!Kilos, "N") - ImpoGastos
            KilosRet = KilosRet + ImpoGastos
            
            ' importe retirada
            Sql4 = "select sum(gastos) from tmpliquidacion1 "
            Sql4 = Sql4 & " where codsocio = " & DBSet(RS!Codsocio, "N")
            Sql4 = Sql4 & " and codusu = " & vUsu.Codigo
            Sql4 = Sql4 & " and codvarie = " & DBSet(RS!codvarie, "N")
            
            ImpoGastos = DevuelveValor(Sql4)
            
            baseimpo = DBLet(RS!Importe, "N") - ImpoGastos
            ImporRet = ImporRet + ImpoGastos
            
            ImpoGastos = 0
            
            HayReg = True
            
            RS.MoveNext
        Wend
            
        ' ultimo registro si ha entrado
        If HayReg Then
            
            ' [Monica] 16/03/2010
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc.esanticipogasto = 0 "
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
            Sql4 = Sql4 & " and rfactsoc.esretirada = 0"
            
            Anticipos = DevuelveValor(Sql4)
            ' [Monica] 16/03/2010
            
            
            Bruto = baseimpo
            
            ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
            
            baseimpo = baseimpo - ImpoGastos - Anticipos
            
            ImpoIva = Round2((baseimpo) * ImporteSinFormato(vPorcIva) / 100, 2)
            
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
        
            ImpoAport = Round2((Bruto - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet + KilosRet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto + ImporRet, "N") & ","
            Sql1 = Sql1 & DBSet(ImporRet, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
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
        
        CargarTemporalLiquidacionQuatretonda = True
        Exit Function
    End If
        
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalQuatretonda(ctabla As String, cwhere As String, Tipo As Byte) As Boolean
'tipo  0=anticipos
'      1=liquidacion
Dim RS As ADODB.Recordset
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
Dim Cad As String
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
Dim KilosRet As Long
Dim ImporRet As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalQuatretonda = False

    Sql2 = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    Sql2 = "delete from tmpliquidacion1 where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, "
    SQL = SQL & " rhisfruta.recolect, rhisfruta.tipoentr, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, "
    SQL = SQL & " sum(if(rhisfruta_clasif.kilosnet is null,0,rhisfruta_clasif.kilosnet)) as kilos "
    SQL = SQL & " FROM  (" & ctabla & ") inner join tmpvarie on rhisfruta.codvarie = tmpvarie.codvarie "
    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8 "
    SQL = SQL & " having sum(rhisfruta_clasif.kilosnet) <> 0 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8 "


    Nregs = TotalRegistrosConsulta(SQL)
    
    Label2(10).Caption = "Cargando Tabla Temporal"
    Me.Pb1.visible = True
    Me.Pb1.Max = Nregs
    Me.Pb1.Value = 0
    Me.Refresh

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    '[Monica] 14/12/2009 si es liquidacion complementaria: no se descuentan gastos(complementaria) ponemos gastos = 0
    If Tipo = 1 And Check1(5).Value = 1 Then Tipo = 3 'seleccionamos los precios de liquidacion complementaria
                                    
                                    
    While Not RS.EOF
    
        Label2(12).Caption = "Socio " & RS!Codsocio & " Variedad " & RS!codvarie & "-" & RS!codcalid & " Campo " & RS!codcampo
        IncrementarProgresNew Pb1, 1
        Me.Refresh
        DoEvents
    
        Sql3 = "select fechaini, fechafin, max(contador) as contador from rprecios where codvarie = " & DBSet(RS!codvarie, "N")
        Sql3 = Sql3 & " and tipofact = " & DBSet(Tipo, "N")
        Sql3 = Sql3 & " and fechaini <= " & DBSet(RS!Fecalbar, "F")
        Sql3 = Sql3 & " and fechafin >= " & DBSet(RS!Fecalbar, "F")
        Sql3 = Sql3 & " group by 1,2"
                
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS1.EOF Then
            Contador = DBLet(RS1!Contador, "N")
            FechaIni = DBLet(RS1!FechaIni, "F")
            FechaFin = DBLet(RS1!FechaFin, "F")
        End If
        Set RS1 = Nothing
        If DBLet(RS!Recolect, "N") = 0 Then 'cooperativa
            Sql3 = "select precoop "
        Else
            Sql3 = "select presocio "
        End If
        
        Sql3 = Sql3 & " from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql3 = Sql3 & " and codcalid = " & DBSet(RS!codcalid, "N")
        Sql3 = Sql3 & " and tipofact = " & DBSet(Tipo, "N")
        Sql3 = Sql3 & " and contador = " & DBSet(Contador, "N")
        
        Precio = DevuelveValor(Sql3)
        
        
'        Sql4 = "select sum(if(isnull(imptrans),0,imptrans)) + sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) as gastos"
'        Sql4 = Sql4 & "  from rhisfruta "
'        Sql4 = Sql4 & " where rhisfruta.codsocio = " & DBSet(Rs!CodSocio, "N") & "  and "
'        Sql4 = Sql4 & " rhisfruta.codvarie = " & DBSet(Rs!CodVarie, "N") & "  and "
'        Sql4 = Sql4 & " rhisfruta.codcampo = " & DBSet(Rs!codcampo, "N") & " and "
'        Sql4 = Sql4 & " rhisfruta.fecalbar >= " & DBSet(FechaIni, "F") & " and "
'        Sql4 = Sql4 & " rhisfruta.fecalbar <= " & DBSet(FechaFin, "F")
        
        Sql4 = "select sum(kilosnet) from rfactsoc_variedad inner join rfactsoc on rfactsoc_variedad.codtipom=rfactsoc.codtipom and rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
        Sql4 = Sql4 & " where rfactsoc.codsocio = " & DBSet(RS!Codsocio, "N") & "  and "
        Sql4 = Sql4 & " rfactsoc_variedad.codvarie = " & DBSet(RS!codvarie, "N") & "  and "
        '[Monica]23/12/2014: el tipo de movimiento de veto ruso tb es de retirada
        Sql4 = Sql4 & " rfactsoc.codtipom in ('FAA','VAA') and rfactsoc.esretirada = 1"
        '[Monica]05/12/2011: los kilos de retirada solo se restan en la primera liquidacion
        Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0 "
        
        
        KilosRet = DevuelveValor(Sql4)
        
        Sql4 = "select sum(imporvar) from rfactsoc_variedad inner join rfactsoc on rfactsoc_variedad.codtipom=rfactsoc.codtipom and rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
        Sql4 = Sql4 & " where rfactsoc.codsocio = " & DBSet(RS!Codsocio, "N") & "  and "
        Sql4 = Sql4 & " rfactsoc_variedad.codvarie = " & DBSet(RS!codvarie, "N") & "  and "
        Sql4 = Sql4 & " rfactsoc.codtipom in ('FAA','VAA') and rfactsoc.esretirada = 1"
        '[Monica]05/12/2011: los kilos de retirada solo se restan en la primera liquidacion
        Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0 "
        
        ImporRet = DevuelveValor(Sql4)
        
        
        Sql5 = "select count(*) from tmpliquidacion1 where codsocio = " & DBSet(RS!Codsocio, "N") & "  and "
        Sql5 = Sql5 & " tmpliquidacion1.codvarie = " & DBSet(RS!codvarie, "N") & "  and "
        Sql5 = Sql5 & " tmpliquidacion1.codusu = " & vUsu.Codigo
        
        If TotalRegistros(Sql5) = 0 Then
            Sql5 = "insert into tmpliquidacion1 values (" & vUsu.Codigo & "," & DBSet(RS!Codsocio, "N") & ","
            Sql5 = Sql5 & DBSet(RS!codvarie, "N") & ","
            Sql5 = Sql5 & DBSet(KilosRet, "N") & ","
            Sql5 = Sql5 & "'0000-00-00',"
            Sql5 = Sql5 & "'0000-00-00',"
            Sql5 = Sql5 & DBSet(ImporRet, "N") & ")"
            
            conn.Execute Sql5
        End If

        ' si no tiene precio no insertamos en la tabla
        
'30/07/2009
'        If Precio <> 0 Then
'            Importe = Round2(Precio * DBLet(RS!kilos, "N"), 2)
            Sql2 = "select count(*) from tmpliquidacion where codusu = " & vUsu.Codigo
            Sql2 = Sql2 & " and codsocio = " & DBSet(RS!Codsocio, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(RS!codcampo, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and recolect = " & DBSet(RS!Recolect, "N")
            Sql2 = Sql2 & " and tipoentr = " & DBSet(RS!TipoEntr, "N")
            Sql2 = Sql2 & " and codcalid = " & DBSet(RS!codcalid, "N")
            Sql2 = Sql2 & " and contador = " & DBSet(Contador, "N")
            Sql2 = Sql2 & " and fechaini = " & DBSet(FechaIni, "F")
            Sql2 = Sql2 & " and fechafin = " & DBSet(FechaFin, "F")
            
            If TotalRegistros(Sql2) = 0 Then
                Kilos = 0
                
                Sql3 = "insert into tmpliquidacion (codusu,codsocio,codcampo,recolect,tipoentr,codvarie,codcalid,contador,kilosnet,precio,importe, "
                Sql3 = Sql3 & " nomvarie, fechaini, fechafin, gastos)"
                Sql3 = Sql3 & " values (" & vUsu.Codigo & "," & DBSet(RS!Codsocio, "N") & "," & DBSet(RS!codcampo, "N") & ","
                Sql3 = Sql3 & DBSet(RS!Recolect, "N") & "," & DBSet(RS!TipoEntr, "N") & ","
                Sql3 = Sql3 & DBSet(RS!codvarie, "N") & "," & DBSet(RS!codcalid, "N") & "," & DBSet(Contador, "N") & ","
                Sql3 = Sql3 & DBSet(RS!Kilos, "N") & "," & DBSet(Precio, "N") & "," & DBSet(0, "N") & ","
                Sql3 = Sql3 & DBSet(RS!nomvarie, "T") & "," & DBSet(FechaIni, "F") & ","
                Sql3 = Sql3 & DBSet(FechaFin, "F") & "," & DBSet(Gastos, "N") & ")"
                
                conn.Execute Sql3
                Kilos = RS!Kilos
            Else
                Kilos = Kilos + RS!Kilos
                Sql3 = "update tmpliquidacion set kilosnet = kilosnet + " & DBSet(RS!Kilos, "N")
                Sql3 = Sql3 & " where codusu = " & vUsu.Codigo
                Sql3 = Sql3 & " and codsocio = " & DBSet(RS!Codsocio, "N")
                Sql3 = Sql3 & " and codcampo = " & DBSet(RS!codcampo, "N")
                Sql3 = Sql3 & " and recolect = " & DBSet(RS!Recolect, "N")
                Sql3 = Sql3 & " and tipoentr = " & DBSet(RS!TipoEntr, "N")
                Sql3 = Sql3 & " and codvarie = " & DBSet(RS!codvarie, "N")
                Sql3 = Sql3 & " and codcalid = " & DBSet(RS!codcalid, "N")
                Sql3 = Sql3 & " and contador = " & DBSet(Contador, "N")
                Sql3 = Sql3 & " and fechaini = " & DBSet(FechaIni, "F")
                Sql3 = Sql3 & " and fechafin = " & DBSet(FechaFin, "F")
                
                conn.Execute Sql3
            End If
'30/07/2009
'        End If
        
        RS.MoveNext
    Wend
                                    
    Sql3 = "update tmpliquidacion set importe = round(kilosnet * precio,2) where codusu = " & vUsu.Codigo
    conn.Execute Sql3
                                    
'    'calculo de gastos
'    Sql4 = "delete from tmpliquidacion1 where codusu = " & vUsu.Codigo
'    conn.Execute Sql4
'
'    Sql4 = "insert into tmpliquidacion1 "
'    Sql4 = Sql4 & "select " & vUsu.Codigo & ", rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, tmpliquidacion.fechaini, tmpliquidacion.fechafin, sum(imptrans) + sum(impacarr) + sum(imprecol) + sum(imppenal)"
'    Sql4 = Sql4 & "  from rhisfruta, tmpliquidacion "
'    Sql4 = Sql4 & " where rhisfruta.codsocio = tmpliquidacion.codsocio  and "
'    Sql4 = Sql4 & " rhisfruta.codvarie = tmpliquidacion.codvarie  and"
'    Sql4 = Sql4 & " rhisfruta.codcampo = tmpliquidacion.codcampo  and"
'    Sql4 = Sql4 & " rhisfruta.fecalbar >= tmpliquidacion.fechaini  and"
'    Sql4 = Sql4 & " rhisfruta.fecalbar <= tmpliquidacion.fechafin group by 1,2,3,4,5,6"
'
'    conn.Execute Sql4
                                    
    Me.Pb1.visible = False
    Me.Label2(10).Caption = ""
    Me.Label2(12).Caption = ""
    Me.Refresh
                                    
    CargarTemporalQuatretonda = True
    Exit Function
    
eCargarTemporal:
    Me.Pb1.visible = False
    Me.Label2(10).Caption = ""
    Me.Label2(12).Caption = ""
    Me.Refresh
    
    MuestraError "Cargando temporal", Err.Description
End Function



Private Function ComprobarTiposIVA(Tabla As String, cSelect As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim vSeccion As CSeccion
Dim b As Boolean


    On Error GoTo eComprobarTiposIVA

    ComprobarTiposIVA = False
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            
            SQL = "select distinct codiva from rsocios_seccion where codsecci = " & vParamAplic.Seccionhorto
            SQL = SQL & " and codsocio in (select rhisfruta.codsocio from " & Trim(Tabla) & " where " & Trim(cSelect) & ")"
            SQL = SQL & " group by 1 "
            SQL = SQL & " order by 1 "
            
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            b = True
            
            While Not RS.EOF And b
                If DBLet(RS.Fields(0).Value, "N") = 0 Then
                    b = False
                    MsgBox "Hay socios sin iva en la sección hortofrutícola. Revise.", vbExclamation
                Else
                    SQL = ""
                    SQL = DevuelveDesdeBDNew(cConta, "tiposiva", "codigiva", "codigiva", DBLet(RS.Fields(0).Value, "N"), "N")
                    If SQL = "" Then
                        b = False
                        MsgBox "No existe el codigo de iva " & DBLet(RS.Fields(0).Value, "N") & ". Revise.", vbExclamation
                    End If
                End If
            
                RS.MoveNext
            Wend
        
            Set RS = Nothing
        
            ComprobarTiposIVA = b
        
            vSeccion.CerrarConta
            
            Set vSeccion = Nothing
        End If
    End If
    Exit Function
    
eComprobarTiposIVA:
    MuestraError Err.Number, "Comprobar Tipos Iva", Err.Description
End Function


Private Function HayFacturasConLineasDeDistintoGrupo(nTabla As String, cadSelect1 As String) As Boolean
Dim SQL As String
Dim Tabla As String

    Tabla = "(" & nTabla & ") INNER JOIN productos on variedades.codprodu = productos.codprodu "
    
    Tabla = QuitarCaracterACadena(Tabla, "{")
    Tabla = QuitarCaracterACadena(Tabla, "}")
    
    If cadSelect1 <> "" Then
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "{")
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "}")
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "_1")
    End If

    SQL = "select rfactsoc_variedad.codtipom, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu, count(distinct productos.gruporeten) "
    SQL = SQL & " from " & Tabla
    SQL = SQL & " where " & cadSelect1
    SQL = SQL & " group by 1,2,3 "
    SQL = SQL & " having count(distinct productos.gruporeten) > 1"
    SQL = SQL & " order by 1,2,3 "
    
    HayFacturasConLineasDeDistintoGrupo = (TotalRegistrosConsulta(SQL) > 0)

End Function



Private Function CargarFacturas(ctabla As String, cSelect As String, cTabla2 As String, cSelect2 As String, Optional cTabla3 As String, Optional cSelect3 As String) As Boolean
Dim SQL As String
Dim vCampAnt As CCampAnt
Dim ctabla1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim CadValues As String
Dim BaseIva As Currency
Dim ImpoIva As Currency
Dim SocioAnt As Long
Dim DifBase As Currency
Dim DifRete As Currency
Dim TipoIRPF As Byte
Dim ImpoReten As Currency
Dim BaseReten As Currency
Dim Producto As Long

Dim BDatos As Dictionary
Dim SqlBd As String
Dim RsBd As ADODB.Recordset

Dim cTablaAnticip As String


    On Error GoTo eCargarFacturas
    
    Screen.MousePointer = vbHourglass
    
    Set BDatos = New Dictionary
    
    ' quitamos las llaves de la tabla y where
    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    If cSelect <> "" Then
        cSelect = QuitarCaracterACadena(cSelect, "{")
        cSelect = QuitarCaracterACadena(cSelect, "}")
        cSelect = QuitarCaracterACadena(cSelect, "_1")
    End If
    
    ' idem de tabla2 y where2 - idem transporte
    cTabla2 = QuitarCaracterACadena(cTabla2, "{")
    cTabla2 = QuitarCaracterACadena(cTabla2, "}")
    If cSelect2 <> "" Then
        cSelect2 = QuitarCaracterACadena(cSelect2, "{")
        cSelect2 = QuitarCaracterACadena(cSelect2, "}")
        cSelect2 = QuitarCaracterACadena(cSelect2, "_1")
    End If
    
    ' idem de tabla3 y where3 - rcafter
    cTabla3 = QuitarCaracterACadena(cTabla3, "{")
    cTabla3 = QuitarCaracterACadena(cTabla3, "}")
    If cSelect3 <> "" Then
        cSelect3 = QuitarCaracterACadena(cSelect3, "{")
        cSelect3 = QuitarCaracterACadena(cSelect3, "}")
        cSelect3 = QuitarCaracterACadena(cSelect3, "_1")
    End If
    
    
    
    '[Monica]25/03/2013: la tabla de anticipos para el certificado
    cTablaAnticip = vEmpresa.BDAriagro & ".rfactsoc_anticipos"
    
    ' borramos las tablas temporales donde insertaremos las facturas para los listados
    SQL = "delete from tmprfactsoc where codusu= " & vUsu.Codigo
    conn.Execute SQL
    
    
    ' insertamos las facturas correspondientes a la campaña actual
    SQL = "insert into tmprfactsoc (`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
    SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
    SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, tipo, esretirada) "
    SQL = SQL & "select " & vUsu.Codigo & ", `rfactsoc`.`codtipom`, `rfactsoc`.`numfactu`, `rfactsoc`.`fecfactu`, cast(`codsocio` as char), "
    SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
    SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, 0, `esretirada` from " '[Monica]24/07/2012: metemos si es de retirada
    SQL = SQL & ctabla
    SQL = SQL & " where " & cSelect
    SQL = SQL & " group by 1,2,3,4 "
    conn.Execute SQL
    
    '[Monica]03/03/2011: cargamos el nif del socio
    SQL = "update tmprfactsoc set nif = (select nifsocio from rsocios where rsocios.codsocio = tmprfactsoc.codsocio)"
    SQL = SQL & " where tipo = 0"
    conn.Execute SQL
    
    '[Monica]03/03/2011: cargamos el nombre del socio
    SQL = "update tmprfactsoc set nomsocio = (select nomsocio from rsocios where rsocios.codsocio = tmprfactsoc.codsocio)"
    SQL = SQL & " where tipo = 0"
    conn.Execute SQL
    
    '[Monica]20/01/2014: cargamos el codigo postal del socio
    SQL = "update tmprfactsoc set codpostal = (select codpostal from rsocios where rsocios.codsocio = tmprfactsoc.codsocio)"
    SQL = SQL & " where tipo = 0"
    conn.Execute SQL
    
    
    '[Monica]15/10/2013: borramos los registros FTT que no tengan numero de factura asignada
    SQL = "delete from tmprfactsoc where codusu = " & vUsu.Codigo
    SQL = SQL & " and (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from rfactsoc where pdtenrofact = 1) " ' where codtipom in ('FTT','FAT','FLT') and numfacrec is null)"
    conn.Execute SQL
    
    
 '[Monica]12/01/2012: añado la condicion check1(9).value = 0 para que no me sume gastos en el informe de aportacion
    
    '[Monica]26/08/2011: Modificacion solo para Picassent
    '                    en las facturas de socios quiere que en la columna impapor estén tb los descuentos,
    '                    con lo cual el totalfac será el total a pagar
    '
    If vParamAplic.Cooperativa = 2 And Check1(9).Value = 0 Then
        SQL = "update tmprfactsoc set impapor = if(impapor is null,0,impapor) + (select if(sum(importe) is null,0,sum(importe)) from rfactsoc_gastos where "
        SQL = SQL & " rfactsoc_gastos.codtipom = tmprfactsoc.codtipom and rfactsoc_gastos.numfactu = tmprfactsoc.numfactu "
        SQL = SQL & " and rfactsoc_gastos.fecfactu = tmprfactsoc.fecfactu) "
        SQL = SQL & " where tmprfactsoc.codusu = " & vUsu.Codigo & " and tmprfactsoc.tipo = 0"
        
        conn.Execute SQL
    
        ' ahora el total factura es el total a pagar en Picassent
        SQL = "update tmprfactsoc set totalfac = baseimpo + if(imporiva is null,0,imporiva) - if(impreten is null,0,impreten) - if(impapor is null,0,impapor)  "
        SQL = SQL & " where tmprfactsoc.codusu = " & vUsu.Codigo & " and tmprfactsoc.tipo = 0"
        
        conn.Execute SQL
    
    End If
    
    
    If InStr(1, cSelect, "FTR") Or (cSelect2 <> "" And (OpcionListado = 10 Or OpcionListado = 11)) Then
        ' insertamos las facturas correspondientes a la campaña actual - idem transporte
        SQL = "insert into tmprfactsoc(`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
        SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
        SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, tipo) "
        SQL = SQL & "select " & vUsu.Codigo & ", `rfacttra`.`codtipom`, `rfacttra`.`numfactu`, `rfacttra`.`fecfactu`, `rfacttra`.`codtrans`, "
        SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
        SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, 1 from "
        SQL = SQL & cTabla2
        SQL = SQL & " where " & cSelect2
        SQL = SQL & " group by 1,2,3,4,5 "
        conn.Execute SQL
        
        '[Monica]03/03/2011: cargamos el nif del transportista
        SQL = "update tmprfactsoc set nif = (select niftrans from rtransporte where rtransporte.codtrans = tmprfactsoc.codsocio)"
        SQL = SQL & " where tipo = 1"
        conn.Execute SQL
        
        '[Monica]03/03/2011: cargamos el nombre del transportista
        SQL = "update tmprfactsoc set nomsocio = (select nomtrans from rtransporte where rtransporte.codtrans = tmprfactsoc.codsocio)"
        SQL = SQL & " where tipo = 1"
        conn.Execute SQL
        
        '[Monica]20/01/2014: cargamos el codigo postal del transportista
        SQL = "update tmprfactsoc set codpostal = (select codpostal from rtransporte where rtransporte.codtrans = tmprfactsoc.codsocio)"
        SQL = SQL & " where tipo = 1"
        conn.Execute SQL
        
    End If
    
    
    '[Monica]20/01/2015: en el caso de ser modelo 190, añadimos las de terceros
    If OpcionListado = 10 Then
        SQL = "insert into tmprfactsoc(`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
        SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
        SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, tipo) "
        SQL = SQL & "select " & vUsu.Codigo & ", 'AAA', 1, `rcafter`.`fecfactu`, `rcafter`.`codsocio`, "
        SQL = SQL & "coalesce(baseiva1,0) + coalesce(baseiva2,0) + coalesce(baseiva3,0),`tipoiva1`,`porciva1`,`impoiva1`,0,`basereten`,"
        SQL = SQL & "`retfacpr`,`trefacpr`,0,0,0,`totalfac`, 2 from "
        SQL = SQL & cTabla3
        SQL = SQL & " where " & cSelect3
        SQL = SQL & " group by 1,2,3,4,5 "
        conn.Execute SQL
    End If
    
    
    
    
    If OpcionListado = 11 Then ' caso del 346
        SQL = "delete from tmp346 where codusu= " & vUsu.Codigo
        conn.Execute SQL
        
        ctabla1 = "(" & ctabla & ") INNER JOIN variedades ON rfactsoc_variedad.codvarie = variedades.codvarie "
        ctabla1 = "(" & ctabla1 & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
        ctabla1 = "(" & ctabla1 & ") INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
        SQL = "insert into tmp346 (`codusu`,`codsocio`,`codgrupo`,`importe`) "
        SQL = SQL & " select " & vUsu.Codigo & ", rfactsoc.codsocio, grupopro.codgrupo, sum(rfactsoc_variedad.imporvar) "
        SQL = SQL & " from " & ctabla1
        SQL = SQL & " where " & cSelect & " and grupopro.codgrupo in (4,5) " ' algarrobos y olivos
        SQL = SQL & " group by rfactsoc.codsocio, grupopro.codgrupo  "
        SQL = SQL & " union "
        SQL = SQL & " select " & vUsu.Codigo & ", rfactsoc.codsocio, 0, sum(rfactsoc_variedad.imporvar)"
        SQL = SQL & " from " & ctabla1
        SQL = SQL & " where " & cSelect & " and not grupopro.codgrupo in (4,5) " ' el resto
        SQL = SQL & " group by rfactsoc.codsocio, grupopro.codgrupo  "
        SQL = SQL & " order by 1,2 "
    
        conn.Execute SQL
    End If
    
    If OpcionListado = 8 Or OpcionListado = 9 Then
        SQL = "delete from tmprfactsoc_variedad where codusu= " & vUsu.Codigo
        conn.Execute SQL
        
        SQL = "insert into tmprfactsoc_variedad (`codusu`,`codtipom`,`numfactu`,`fecfactu`,"
        SQL = SQL & "`codvarie`,`codcampo`,`kilosnet`,`preciomed`,`imporvar`,`descontado`) "
        SQL = SQL & " select " & vUsu.Codigo & ", rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu,"
        SQL = SQL & " rfactsoc_variedad.codvarie, rfactsoc_variedad.codcampo, rfactsoc_variedad.kilosnet,"
        SQL = SQL & " rfactsoc_variedad.preciomed, rfactsoc_variedad.imporvar, rfactsoc_variedad.descontado"
        SQL = SQL & " from " & ctabla
        SQL = SQL & " where " & cSelect
    
        conn.Execute SQL
    
        '24/01/2011
        If Check1(7).Value = 1 Then  ' caso del certificado de retenciones
            SQL = "insert into tmprfactsoc_variedad (`codusu`,`codtipom`,`numfactu`,`fecfactu`,"
            SQL = SQL & "`codvarie`,`codcampo`,`kilosnet`,`preciomed`,`imporvar`,`descontado`) "
            SQL = SQL & " select distinct " & vUsu.Codigo & ", rfactsoc_anticipos.codtipom, rfactsoc_anticipos.numfactu, rfactsoc_anticipos.fecfactu,"
            SQL = SQL & " rfactsoc_anticipos.codvarieanti, rfactsoc_anticipos.codcampoanti, 0,rfactsoc_anticipos.numfactuanti,rfactsoc_anticipos.baseimpo * (-1),0 "
            SQL = SQL & " from (" & ctabla & ") Inner join rfactsoc_anticipos on rfactsoc.codtipom = rfactsoc_anticipos.codtipom and rfactsoc.numfactu = rfactsoc_anticipos.numfactu and rfactsoc.fecfactu = rfactsoc_anticipos.fecfactu "
            SQL = SQL & " where " & cSelect
            
            conn.Execute SQL
        
        
        End If
    
        '[Monica]25/05/2015: tenemos que añadir los descuentos para el caso de montifrut
        If vParamAplic.Cooperativa = 12 And Check1(7).Value = 1 Then
            Dim Varie As Long
            Dim campo As Long
            Dim TotalKilos As Long
            Dim ImporteTot As Currency
            Dim vImporte As Currency
            Dim Importe As Currency
            Dim RS5 As ADODB.Recordset
            Dim Rs6 As ADODB.Recordset
            
            
            SQL = "select * from rfactsoc_gastos where (codtipom, numfactu, fecfactu) in (select rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu from " & ctabla & " where " & cSelect & ")"
            SQL = SQL & " order by codtipom, numfactu, fecfactu "
            
            Set RS5 = New ADODB.Recordset
            RS5.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not RS5.EOF
                SQL = "select sum(kilosnet) from tmprfactsoc_variedad where codusu = " & vUsu.Codigo & " and codtipom = " & DBSet(RS5!CodTipom, "T")
                SQL = SQL & " and numfactu = " & DBSet(RS5!numfactu, "N")
                SQL = SQL & " and fecfactu = " & DBSet(RS5!fecfactu, "F")
                
                TotalKilos = DevuelveValor(SQL)
                ImporteTot = DBLet(RS5!Importe)
                vImporte = 0
                
                SQL = "select * from tmprfactsoc_variedad where codusu = " & vUsu.Codigo & " and codtipom = " & DBSet(RS5!CodTipom, "T")
                SQL = SQL & " and numfactu = " & DBSet(RS5!numfactu, "N")
                SQL = SQL & " and fecfactu = " & DBSet(RS5!fecfactu, "F")
                
                Set Rs6 = New ADODB.Recordset
                Rs6.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs6.EOF
                    Importe = Round2(DBLet(Rs6!KilosNet) * DBLet(RS5!Importe) / TotalKilos, 2)
                    vImporte = vImporte + Importe
                    
                    Varie = DBLet(Rs6!codvarie)
                    campo = DBLet(Rs6!codcampo)
                    
                    SQL = "update tmprfactsoc_variedad set imporvar = imporvar - " & DBSet(Importe, "N")
                    SQL = SQL & " where codusu = " & vUsu.Codigo
                    SQL = SQL & " and codtipom = " & DBSet(RS5!CodTipom, "T")
                    SQL = SQL & " and numfactu = " & DBSet(RS5!numfactu, "N")
                    SQL = SQL & " and fecfactu = " & DBSet(RS5!fecfactu, "F")
                    SQL = SQL & " and codvarie = " & DBSet(Varie, "N")
                    SQL = SQL & " and codcampo = " & DBSet(campo, "N")
                    
                    conn.Execute SQL
                    
                    Rs6.MoveNext
                Wend
                Set Rs6 = Nothing
                
                ' si hay diferencia en el ultimo ponemos la diferencia
                If vImporte <> ImporteTot Then
                    SQL = "update tmprfactsoc_variedad set imporvar = impovar + " & DBSet(vImporte - ImporteTot, "N")
                    SQL = SQL & " where codusu = " & vUsu.Codigo
                    SQL = SQL & " and codtipom = " & DBSet(RS5!CodTipom, "T")
                    SQL = SQL & " and numfactu = " & DBSet(RS5!numfactu, "N")
                    SQL = SQL & " and fecfactu = " & DBSet(RS5!fecfactu, "F")
                    SQL = SQL & " and codvarie = " & DBSet(Varie, "N")
                    SQL = SQL & " and codcampo = " & DBSet(campo, "N")
                
                    conn.Execute SQL
                End If
            
                RS5.MoveNext
            Wend
            
            Set RS5 = Nothing
        End If
    
    
    
        ' - idem transporte
        If InStr(1, cSelect, "FTR") Then
            SQL = "insert into tmprfactsoc_variedad (`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
            SQL = SQL & "`codvarie`,`codcampo`,`kilosnet`,`preciomed`,`imporvar`,`descontado`) "
            SQL = SQL & " select " & vUsu.Codigo & ", rfacttra.codtipom, rfacttra.numfactu, rfacttra.fecfactu, rfacttra.codtrans,"
            SQL = SQL & " rfacttra_albaran.codvarie, rfacttra_albaran.codcampo, sum(rfacttra_albaran.kilosnet) kilosnet,"
            SQL = SQL & " 0, sum(rfacttra_albaran.importe) importe, 0"
            SQL = SQL & " from " & cTabla2
            SQL = SQL & " where " & cSelect2
            SQL = SQL & " group by 1,2,3,4,5,6,7,9,11"
            conn.Execute SQL
        End If
        
        
        '[Monica]15/10/2013: borramos los registros FTT que no tengan numero de factura asignada
        SQL = "delete from tmprfactsoc_variedad where codusu = " & vUsu.Codigo
        SQL = SQL & " and (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from rfactsoc where pdtenrofact = 1) " 'where codtipom in ('FTT','FAT','FLT') and numfacrec is null)"
        conn.Execute SQL
    
    End If
    
  
' TODAS LAS CAMPAÑAS ANTERIORES

'[Monica]17/10/2013: Si y solo si no es Montifrut que tiene en otro ariagro otra cosa
    '[Monica]21/01/2015: añado el caso de Natural
    '[Monica]25/01/2016: añado el caso de bolbaite
If vParamAplic.Cooperativa <> 12 And vParamAplic.Cooperativa <> 9 And vParamAplic.Cooperativa <> 14 Then
    SqlBd = "SHOW DATABASES like 'ariagro%' "
    Set RsBd = New ADODB.Recordset
    RsBd.Open SqlBd, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RsBd.EOF
        If Trim(DBLet(RsBd.Fields(0).Value)) <> vEmpresa.BDAriagro And Trim(DBLet(RsBd.Fields(0).Value)) <> "" And InStr(1, DBLet(RsBd.Fields(0).Value), "ariagroutil") = 0 Then
        
            ' borramos la tabla temporal de la campaña anterior
            SQL = "delete from " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc where codusu= " & vUsu.Codigo
            conn.Execute SQL
            
            
            SQL = "insert into " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc (`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
            SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
            SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, tipo, esretirada) "
            SQL = SQL & "select " & vUsu.Codigo & ", `rfactsoc`.`codtipom`, `rfactsoc`.`numfactu`, `rfactsoc`.`fecfactu`, `codsocio`, "
            SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
            SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, 0, `esretirada` " '[Monica]24/07/2012: metemos si es de retirada
            SQL = SQL & " from " & Replace(ctabla, vEmpresa.BDAriagro, RsBd.Fields(0).Value)
            SQL = SQL & " where " & cSelect
            SQL = SQL & " group by 1,2,3,4 "
            conn.Execute SQL
        
            
            '[Monica]26/08/2011: Modificacion solo para Picassent
            '                    en las facturas de socios quiere que en la columna impapor estén tb los descuentos,
            '                    con lo cual el totalfac será el total a pagar
            '
            If vParamAplic.Cooperativa = 2 Then
                SQL = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc set impapor = if(impapor is null,0,impapor) + (select if(sum(importe) is null,0,sum(importe)) from rfactsoc_gastos where "
                SQL = SQL & " rfactsoc_gastos.codtipom = tmprfactsoc.codtipom and rfactsoc_gastos.numfactu = tmprfactsoc.numfactu "
                SQL = SQL & " and rfactsoc_gastos.fecfactu = tmprfactsoc.fecfactu) "
                SQL = SQL & " where tmprfactsoc.codusu = " & vUsu.Codigo & " and tmprfactsoc.tipo = 0"
                
                conn.Execute SQL
            
                ' ahora el total factura es el total a pagar en Picassent
                SQL = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc set totalfac = baseimpo + if(imporiva is null,0,imporiva) - if(impreten is null,0,impreten) - if(impapor is null,0,impapor)  "
                SQL = SQL & " where tmprfactsoc.codusu = " & vUsu.Codigo & " and tmprfactsoc.tipo = 0"
                
                conn.Execute SQL
            End If
            
            
            '[Monica]03/03/2011: cargamos el nif del socio
'            Sql = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc set nif = (select nifsocio from rsocios where rsocios.codsocio = tmprfactsoc.codsocio)"
            SQL = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc dd, " & Trim(RsBd.Fields(0).Value) & ".rsocios ff set dd.nif = ff.nifsocio where ff.codsocio = dd.codsocio"
            SQL = SQL & " and dd.tipo = 0"
            conn.Execute SQL
            
            '[Monica]03/03/2011: cargamos el nif del socio
'            Sql = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc set nomsocio = (select nomsocio from rsocios where rsocios.codsocio = tmprfactsoc.codsocio)"
            SQL = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc dd, " & Trim(RsBd.Fields(0).Value) & ".rsocios ff set dd.nomsocio = ff.nomsocio where ff.codsocio = dd.codsocio"
            SQL = SQL & " and dd.tipo = 0"
            conn.Execute SQL
            
            '[Monica]20/01/2014: cargamos el codigo postal del socio
            SQL = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc dd, " & Trim(RsBd.Fields(0).Value) & ".rsocios ff set dd.codpostal = ff.codpostal where ff.codsocio = dd.codsocio"
            SQL = SQL & " and dd.tipo = 0"
            conn.Execute SQL
            
            
            
            '[Monica]15/10/2013: borramos los registros FTT que no tengan numero de factura asignada
            SQL = "delete from " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc where codusu = " & vUsu.Codigo
            SQL = SQL & " and (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from " & Trim(RsBd.Fields(0).Value) & ".rfactsoc where pdtenrofact = 1) " 'where codtipom in ('FTT','FAT','FLT') and numfacrec is null)"
            conn.Execute SQL
            
            
            If InStr(1, cSelect, "FTR") Or (cSelect2 <> "" And (OpcionListado = 10 Or OpcionListado = 11)) Then
                SQL = "insert into " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc (`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
                SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
                SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, tipo) "
                SQL = SQL & "select " & vUsu.Codigo & ", `rfacttra`.`codtipom`, `rfacttra`.`numfactu`, `rfacttra`.`fecfactu`, `rfacttra`.`codtrans`, "
                SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
                SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, 1 "
                SQL = SQL & " from " & Replace(cTabla2, vEmpresa.BDAriagro, RsBd.Fields(0).Value)
                SQL = SQL & " where " & cSelect2
                SQL = SQL & " group by 1,2,3,4,5 "
                conn.Execute SQL
            
                '[Monica]03/03/2011: cargamos el nif del transportista
'                Sql = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc set nif = (select niftrans from rtransporte where rtransporte.codtrans = tmprfactsoc.codsocio)"
                SQL = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc dd, " & Trim(RsBd.Fields(0).Value) & ".rtransporte ff set dd.nif = ff.niftrans where ff.codtrans = dd.codsocio"
                SQL = SQL & " and dd.tipo = 1"
                conn.Execute SQL
            
                '[Monica]03/03/2011: cargamos el nif del transportista
'                Sql = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc set nomsocio = (select nomtrans from rtransporte where rtransporte.codtrans = tmprfactsoc.codsocio)"
                SQL = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc dd, " & Trim(RsBd.Fields(0).Value) & ".rtransporte ff set dd.nomsocio = ff.nomtrans where ff.codtrans = dd.codsocio"
                SQL = SQL & " and dd.tipo = 1"
                conn.Execute SQL
            
                '[Monica]20/01/2014: cargamos el codigo postal del transportista
                SQL = "update " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc dd, " & Trim(RsBd.Fields(0).Value) & ".rtransporte ff set dd.codpostal = ff.codpostal where ff.codtrans = dd.codsocio"
                SQL = SQL & " and dd.tipo = 1"
                conn.Execute SQL
            
            End If
        
            '[Monica]20/01/2015: en el caso de ser modelo 190, añadimos las de terceros
            If OpcionListado = 10 Then
                SQL = "insert into " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc(`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
                SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
                SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, tipo) "
                SQL = SQL & "select " & vUsu.Codigo & ", 'AAA', 1, `rcafter`.`fecfactu`, `rcafter`.`codsocio`, "
                SQL = SQL & "coalesce(baseiva1,0) + coalesce(baseiva2,0) + coalesce(baseiva3,0),`tipoiva1`,`porciva1`,`impoiva1`,0,`basereten`,"
                SQL = SQL & "`retfacpr`,`trefacpr`,0,0,0,`totalfac`, 2 from "
                SQL = SQL & Replace(cTabla3, vEmpresa.BDAriagro, RsBd.Fields(0).Value)
                SQL = SQL & " where " & cSelect3
                SQL = SQL & " group by 1,2,3,4,5 "
                conn.Execute SQL
            End If
        
            If OpcionListado = 11 Then ' caso del 346
                SQL = "delete from " & Trim(RsBd.Fields(0).Value) & ".tmp346 where codusu= " & vUsu.Codigo
                conn.Execute SQL
                
                ctabla1 = "(" & Replace(ctabla, vEmpresa.BDAriagro, Trim(RsBd.Fields(0).Value)) & ") INNER JOIN " & Trim(RsBd.Fields(0).Value) & ".variedades ON rfactsoc_variedad.codvarie = variedades.codvarie "
                ctabla1 = "(" & ctabla1 & ") INNER JOIN " & Trim(RsBd.Fields(0).Value) & ".productos ON variedades.codprodu = productos.codprodu "
                ctabla1 = "(" & ctabla1 & ") INNER JOIN " & Trim(RsBd.Fields(0).Value) & ".grupopro ON productos.codgrupo = grupopro.codgrupo "
                
                SQL = "insert into " & Trim(RsBd.Fields(0).Value) & ".tmp346 (`codusu`,`codsocio`,`codgrupo`,`importe`) "
                SQL = SQL & " select " & vUsu.Codigo & ", rfactsoc.codsocio, grupopro.codgrupo, sum(rfactsoc_variedad.imporvar) "
                SQL = SQL & " from " & ctabla1 '04/02/2014:antes estaba mal, Replace(ctabla1, vEmpresa.BDAriagro, Trim(RsBd.Fields(0).Value))
                SQL = SQL & " where " & cSelect & " and grupopro.codgrupo in (4,5) " ' algarrobos y olivos
                SQL = SQL & " group by rfactsoc.codsocio, grupopro.codgrupo  "
                SQL = SQL & " union "
                SQL = SQL & " select " & vUsu.Codigo & ", rfactsoc.codsocio, 0, sum(rfactsoc_variedad.imporvar)"
                SQL = SQL & " from " & ctabla1 '04/02/2014:antes estaba mal, Replace(ctabla1, vEmpresa.BDAriagro, Trim(RsBd.Fields(0).Value))
                SQL = SQL & " where " & cSelect & " and not grupopro.codgrupo in (4,5) " ' el resto
                SQL = SQL & " group by rfactsoc.codsocio, grupopro.codgrupo  "
                SQL = SQL & " order by 1,2 "
                
                conn.Execute SQL
            End If
            
            If OpcionListado = 8 Or OpcionListado = 9 Then
                SQL = "delete from " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc_variedad where codusu= " & vUsu.Codigo
                conn.Execute SQL
                
                SQL = "insert into " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc_variedad (`codusu`,`codtipom`,`numfactu`,`fecfactu`,"
                SQL = SQL & "`codvarie`,`codcampo`,`kilosnet`,`preciomed`,`imporvar`,`descontado`) "
                SQL = SQL & " select " & vUsu.Codigo & ", rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu,"
                SQL = SQL & " rfactsoc_variedad.codvarie, rfactsoc_variedad.codcampo, rfactsoc_variedad.kilosnet,"
                SQL = SQL & " rfactsoc_variedad.preciomed, rfactsoc_variedad.imporvar, rfactsoc_variedad.descontado "
                SQL = SQL & " from " & Replace(ctabla, vEmpresa.BDAriagro, Trim(RsBd.Fields(0).Value))
                SQL = SQL & " where " & cSelect
            
                conn.Execute SQL
                                       
                If Check1(7).Value = 1 Then  ' caso del certificado de retenciones
                    SQL = "insert into " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc_variedad (`codusu`,`codtipom`,`numfactu`,`fecfactu`,"
                    SQL = SQL & "`codvarie`,`codcampo`,`kilosnet`,`preciomed`,`imporvar`,`descontado`) "
                    SQL = SQL & " select distinct " & vUsu.Codigo & ", rfactsoc_anticipos.codtipom, rfactsoc_anticipos.numfactu, rfactsoc_anticipos.fecfactu,"
                    SQL = SQL & " rfactsoc_anticipos.codvarieanti, rfactsoc_anticipos.codcampoanti, 0,rfactsoc_anticipos.numfactuanti,rfactsoc_anticipos.baseimpo * (-1),0 "
                    SQL = SQL & " from (" & Replace(ctabla, vEmpresa.BDAriagro, Trim(RsBd.Fields(0).Value)) & ") Inner join " & Replace(cTablaAnticip, vEmpresa.BDAriagro, Trim(RsBd.Fields(0).Value)) & " on rfactsoc.codtipom = rfactsoc_anticipos.codtipom and rfactsoc.numfactu = rfactsoc_anticipos.numfactu and rfactsoc.fecfactu = rfactsoc_anticipos.fecfactu "
                    SQL = SQL & " where " & cSelect
                    conn.Execute SQL
                
                
                End If
            
            
                If InStr(1, cSelect, "FTR") Then
                    SQL = "insert into " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc_variedad (`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
                    SQL = SQL & "`codvarie`,`codcampo`,`kilosnet`,`preciomed`,`imporvar`,`descontado`) "
                    SQL = SQL & " select " & vUsu.Codigo & ", rfacttra.codtipom, rfacttra.numfactu, rfacttra.fecfactu, rfacttra.codtrans,"
                    SQL = SQL & " rfacttra_albaran.codvarie, rfacttra_albaran.codcampo, sum(rfacttra_albaran.kilosnet) kilosnet,"
                    SQL = SQL & " 0, sum(rfacttra_albaran.importe) importe , 0 "
                    SQL = SQL & " from " & Replace(cTabla2, vEmpresa.BDAriagro, Trim(RsBd.Fields(0).Value))
                    SQL = SQL & " where " & cSelect2
                    SQL = SQL & " group by 1,2,3,4,5,6,7,9,11 "
                    conn.Execute SQL
                End If
            
                '[Monica]15/10/2013: borramos los registros FTT que no tengan numero de factura asignada
                SQL = "delete from " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc_variedad where codusu = " & vUsu.Codigo
                SQL = SQL & " and (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from " & Trim(RsBd.Fields(0).Value) & ".rfactsoc where pdtenrofact = 1) " 'codtipom in ('FTT','FAT','FLT') and numfacrec is null)"
                conn.Execute SQL
            End If
        
        
            ' introducimos las facturas de la campaña anterior en la temporal de la
            ' campaña actual
            SQL = "insert into tmprfactsoc select * from " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc "
            SQL = SQL & " where codusu = " & vUsu.Codigo
            
            conn.Execute SQL
            
            
            If OpcionListado = 11 Then
                SQL = "insert into tmp346 select * from " & Trim(RsBd.Fields(0).Value) & ".tmp346 "
                SQL = SQL & " where codusu = " & vUsu.Codigo
                
                conn.Execute SQL
            End If
            
            If OpcionListado = 8 Or OpcionListado = 9 Then
                SQL = "insert into tmprfactsoc_variedad select * from " & Trim(RsBd.Fields(0).Value) & ".tmprfactsoc_variedad "
                SQL = SQL & " where codusu = " & vUsu.Codigo
                
                conn.Execute SQL
            
            End If
        End If
    
        RsBd.MoveNext
    Wend
  
    Set RsBd = Nothing
End If ' Por Montifrut que tiene la otra en el ariagro2
  
    ' [Monica] 11/05/2010: hacemos todos los calculos aunque luego no los impriman
    If Check1(7).Value = 1 Then ' And vParamAplic.Cooperativa = 0  Then ' si estamos en el certificado de retenciones de Catadau
        SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
        conn.Execute SQL
    
    
        '[Monica]19/01/2012: en Quatretonda lo quieren ordenado por variedad
        If vParamAplic.Cooperativa = 7 Then
            SQL = "select nif, tipo , tmprfactsoc.codsocio codigo , variedades.codvarie codprodu, tmprfactsoc.porc_iva, max(tmprfactsoc.porc_ret) porc_ret, sum(tmprfactsoc_variedad.imporvar) importe "
        Else
            SQL = "select nif, tipo , tmprfactsoc.codsocio codigo , variedades.codprodu, tmprfactsoc.porc_iva, max(tmprfactsoc.porc_ret) porc_ret, sum(tmprfactsoc_variedad.imporvar) importe "
        End If
        SQL = SQL & " from (tmprfactsoc inner join tmprfactsoc_variedad on tmprfactsoc.codtipom = tmprfactsoc_variedad.codtipom "
        SQL = SQL & " and tmprfactsoc.codusu = tmprfactsoc_variedad.codusu "
        SQL = SQL & " and tmprfactsoc.numfactu = tmprfactsoc_variedad.numfactu  and tmprfactsoc.fecfactu = tmprfactsoc_variedad.fecfactu) "
        SQL = SQL & " inner join variedades on tmprfactsoc_variedad.codvarie = variedades.codvarie "
        SQL = SQL & " where tmprfactsoc.codusu = " & vUsu.Codigo
        SQL = SQL & " group by 1,2,3,4,5 "
        SQL = SQL & " order by 1,2,3,4,5"

        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                                                'nif,    0/1   soc/trans, codprodu, basereten,impreten
        Sql2 = "insert into tmpinformes (codusu, nombre1,codigo1, nombre2, importe1, importe2, importe3) values "
        
        
        CadValues = ""
        
        While Not RS.EOF
            Select Case RS.Fields(1).Value
                Case 0
                    TipoIRPF = DevuelveValor("select tipoirpf from rsocios where codsocio = " & DBSet(RS!Codigo, "N"))
                Case 1
                    TipoIRPF = DevuelveValor("select tipoirpf from rtranspor where codtrans = " & DBSet(RS!Codigo, "T"))
            End Select
            
            BaseIva = DBLet(RS!Importe, "N")
            ImpoIva = Round(BaseIva * DBLet(RS!porc_iva, "N") / 100, 2)
            Select Case TipoIRPF
                Case 0
                    ImpoReten = Round2((BaseIva + ImpoIva) * DBLet(RS!porc_ret, "N") / 100, 2)
                    BaseReten = (BaseIva + ImpoIva)
                Case 1
                    ImpoReten = Round2(BaseIva * DBLet(RS!porc_ret, "N") / 100, 2)
                    BaseReten = BaseIva
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
            End Select
            
            Sql4 = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo & " and nombre1 = " & DBSet(RS.Fields(0).Value, "T")
            Sql4 = Sql4 & " and codigo1 = " & DBSet(RS.Fields(1).Value, "N")
            Sql4 = Sql4 & " and nombre2 = " & DBSet(RS!Codigo, "T")
            Sql4 = Sql4 & " and importe1 = " & DBSet(RS!codprodu, "N")
            If TotalRegistros(Sql4) <> 0 Then
                Sql4 = "update tmpinformes set importe2 = importe2 + " & DBSet(BaseReten, "N") & ","
                Sql4 = Sql4 & " importe3 = importe3 + " & DBSet(ImpoReten, "N")
                Sql4 = Sql4 & " where codusu = " & vUsu.Codigo & " and nombre1 = " & DBSet(RS.Fields(0).Value, "T")
                Sql4 = Sql4 & " and codigo1 = " & DBSet(RS.Fields(1).Value, "N")
                Sql4 = Sql4 & " and nombre2 = " & DBSet(RS!Codigo, "T")
                Sql4 = Sql4 & " and importe1 = " & DBSet(RS!codprodu, "N")
                
                conn.Execute Sql4
                
            Else
            
                CadValues = "(" & vUsu.Codigo & "," & DBSet(RS.Fields(0).Value, "T") & "," & DBSet(RS.Fields(1).Value, "N") & "," & DBSet(RS!Codigo, "T") & "," & DBSet(RS!codprodu, "N") & ","
                CadValues = CadValues & DBSet(BaseReten, "N") & ","
                CadValues = CadValues & DBSet(ImpoReten, "N") & ")"
                
                conn.Execute Sql2 & CadValues
                    
            End If
            RS.MoveNext
        Wend
        Set RS = Nothing
        If CadValues <> "" Then
            
            ' comprobamos por socio si cuadra con la base de retencion e importe de retencion del socio
            Sql2 = "select codigo1, nombre2 codigo, sum(importe2) basereten, sum(importe3) impreten from tmpinformes where codusu = " & vUsu.Codigo
            Sql2 = Sql2 & " group by 1, 2 order by 1, 2 "
            
            Set RS = New ADODB.Recordset
            RS.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                Select Case RS.Fields(0).Value
                    Case 0
                        SQL = "select sum(basereten) base, sum(impreten) reten from tmprfactsoc where codusu = " & vUsu.Codigo
                        SQL = SQL & " and codsocio = " & DBSet(RS!Codigo, "N")
                        SQL = SQL & " and tipo = 0"
                    Case 1 '"B"
                        SQL = "select sum(basereten) base, sum(impreten) reten from tmprfactsoc where codusu = " & vUsu.Codigo
                        SQL = SQL & " and codsocio = " & DBSet(RS!Codigo, "T")
                        SQL = SQL & " and tipo = 1"
                End Select
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs2.EOF Then
                    DifBase = DBLet(Rs2!Base, "N") - DBLet(RS!BaseReten, "N")
                    DifRete = DBLet(Rs2!Reten, "N") - DBLet(RS!ImpReten, "N")
                
                    If DifBase <> 0 Or DifRete <> 0 Then
                        'cogemos el maximo producto para actualizarle el redondeo
                        Sql2 = "select max(importe1) from tmpinformes where codusu = " & vUsu.Codigo
                        Sql2 = Sql2 & " and codigo1 = " & DBSet(RS.Fields(0).Value, "N")
                        Sql2 = Sql2 & " and nombre2 = " & DBSet(RS!Codigo, "T")
                        
                        Producto = DevuelveValor(Sql2)
                        
                        Sql2 = "update tmpinformes set importe2 = importe2 + (" & DBSet(DifBase, "N") & "),"
                        Sql2 = Sql2 & " importe3 = importe3 + (" & DBSet(DifRete, "N") & ")"
                        Sql2 = Sql2 & " where codusu = " & DBSet(vUsu.Codigo, "N")
                        Sql2 = Sql2 & " and codigo1 = " & DBSet(RS.Fields(0).Value, "N")
                        Sql2 = Sql2 & " and nombre2 = " & DBSet(RS!Codigo, "T")
                        Sql2 = Sql2 & " and importe1 = " & DBSet(Producto, "N")
                        
                        conn.Execute Sql2
                    End If
                End If
                Set Rs2 = Nothing
            
                RS.MoveNext
            Wend
            Set RS = Nothing
            
        End If
    
    End If
    
    
    CargarFacturas = True
    Screen.MousePointer = vbDefault
    Exit Function
    
eCargarFacturas:
    MuestraError Err.Number, "Cargar Facturas", Err.Description
    Screen.MousePointer = vbDefault
    CargarFacturas = False
End Function


Private Function CargarTemporalLiquidacionIndustria(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
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
Dim Cad As String
Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionIndustria = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    If CargarTemporalIndustria(ctabla, cwhere) Then
        SQL = "SELECT tmpliquidacion.codsocio, tmpliquidacion.codcampo, tmpliquidacion.codvarie, variedades.nomvarie,"
        SQL = SQL & " sum(tmpliquidacion.kilosnet) as kilos , sum(tmpliquidacion.importe) as importe "
        SQL = SQL & " FROM tmpliquidacion, variedades where codusu = " & vUsu.Codigo
        SQL = SQL & " and tmpliquidacion.codvarie = variedades.codvarie "
        SQL = SQL & " group by 1, 2, 3, 4 "
        SQL = SQL & " order by 1, 2, 3, 4 "
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                                        'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  gastos,    codcampo, baseimpo, porceiva, imporiva,
        Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                       'porcerete, imporret, totalfac
        Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
        
        Set vSeccion = New CSeccion
        
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If Not vSeccion.AbrirConta Then
                Exit Function
            End If
        End If
        
        HayReg = False
        If Not RS.EOF Then
            SocioAnt = RS!Codsocio
            VarieAnt = RS!codvarie
            NVarieAnt = RS!nomvarie
            CampoAnt = RS!codcampo
            
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                NSocioAnt = vSocio.Nombre
                TipoIRPF = vSocio.TipoIRPF
            End If
        End If
        
        While Not RS.EOF
        
            If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Or CampoAnt <> RS!codcampo Then
                
                Bruto = baseimpo
                
                baseimpo = baseimpo - ImpoGastos
                
                ImpoIva = Round2((baseimpo) * ImporteSinFormato(vPorcIva) / 100, 2)
            
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
            
                ' No hay fondo de aportacion
                ' ImpoAport = Round2((Bruto - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
            
                TotalFac = baseimpo + ImpoIva - ImpoReten '- ImpoAport
                
                Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
                Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
                Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
                Sql1 = Sql1 & DBSet(Bruto, "N") & ","
                Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
                Sql1 = Sql1 & DBSet(CampoAnt, "N") & ","
                Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
                Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
                Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
                Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
                Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
                
                VarieAnt = RS!codvarie
                CampoAnt = RS!codcampo
                
                baseimpo = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                KilosNet = 0
                
                ImpoGastos = 0
                
            End If
            
            If RS!Codsocio <> SocioAnt Then
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(RS!Codsocio) Then
                    If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                    NSocioAnt = vSocio.Nombre
                End If
                SocioAnt = vSocio.Codigo
                TipoIRPF = vSocio.TipoIRPF
            End If
            
            KilosNet = DBLet(RS!Kilos, "N")
            
            baseimpo = DBLet(RS!Importe, "N")
                
            ' gastos
            Sql4 = "select sum(gastos) from tmpliquidacion1 "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codusu = " & vUsu.Codigo
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and codcampo = " & DBSet(CampoAnt, "N")
            
            ImpoGastos = DevuelveValor(Sql4)
                
                
            HayReg = True
            
            RS.MoveNext
        Wend
            
        ' ultimo registro si ha entrado
        If HayReg Then
            Bruto = baseimpo
            
            baseimpo = baseimpo - ImpoGastos
            
            ImpoIva = Round2((baseimpo) * ImporteSinFormato(vPorcIva) / 100, 2)
            
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
            
            ' No hay fondo de aportacion
            'ImpoAport = Round2((Bruto - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten '- ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(CampoAnt, "N") & ","
            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
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
        
        CargarTemporalLiquidacionIndustria = True
        Exit Function
    End If
    

eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function





Private Function HayPreciosVariedadesIndustria(ctabla As String, cwhere As String) As Boolean
'Comprobar si hay precios para cada una de las variedades seleccionadas
Dim SQL As String
Dim vPrecios As CPrecios
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim b As Boolean
Dim Sql2 As String
Dim Sql5 As String
Dim VarieAnt As Long
Dim Numreg As Long

    On Error GoTo eHayPreciosVariedadesIndustria
    
    HayPreciosVariedadesIndustria = False
    
    conn.Execute " DROP TABLE IF EXISTS tmpVarie;"
    
    SQL = "CREATE TEMPORARY TABLE tmpVarie ( " 'TEMPORARY
    SQL = SQL & "codvarie INT(6) UNSIGNED  DEFAULT '0' NOT NULL) "
    conn.Execute SQL
    
    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    SQL = "Select distinct rhisfruta.codvarie, rhisfruta.fecalbar FROM " & QuitarCaracterACadena(ctabla, "_1")
    
'    Sql2 = "Select distinct rhisfruta_clasif.codvarie, rhisfruta_clasif.codcalid FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
'        Sql2 = Sql2 & " where " & cWhere
    End If
    SQL = SQL & " group by 1,2 "
    SQL = SQL & " order by 1,2 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True
    
    If Not RS.EOF Then VarieAnt = DBLet(RS!codvarie, "N")
    Numreg = 0
    ' comprobamos que existen registros para todos las variedades seleccionadas
    While Not RS.EOF And b
        Sql2 = "select * from rprecios where (codvarie, tipofact, contador) = ("
        Sql2 = Sql2 & "SELECT codvarie, tipofact, max(contador) FROM rprecios WHERE codvarie=" & DBSet(RS!codvarie, "N") & " and "
        Sql2 = Sql2 & " tipofact = 2 and fechaini <= " & DBSet(RS!Fecalbar, "F")
        Sql2 = Sql2 & " and fechafin >= " & DBSet(RS!Fecalbar, "F") & " and precioindustria <> 0 and precioindustria is not null "
        Sql2 = Sql2 & " group by 1, 2) "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Rs2.EOF Then
            b = False
            MsgBox "No existe precio de Industria para la variedad " & DBLet(RS!codvarie, "N") & " de fecha " & DBLet(RS!Fecalbar, "F") & ". Revise.", vbExclamation
        Else
            Sql5 = "select count(*) from tmpvarie where codvarie = " & DBSet(RS!codvarie, "N")
            If TotalRegistros(Sql5) = 0 Then
                Sql5 = "insert into tmpVarie (codvarie) values (" & DBSet(RS!codvarie, "N") & ")"
                conn.Execute Sql5
            End If
        End If
            
        Set Rs2 = Nothing
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    HayPreciosVariedadesIndustria = b
    Exit Function
    
eHayPreciosVariedadesIndustria:
    MuestraError Err.nume, "Comprobando si hay precios de Industria en variedades", Err.Description
End Function



Private Function CargarTemporalIndustria(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
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
Dim Cad As String
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
    
    CargarTemporalIndustria = False
    
    
    Sql2 = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    Sql2 = "delete from tmpliquidacion1 where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, "
    SQL = SQL & " rhisfruta_clasif.codcalid, rhisfruta.fecalbar, "
    SQL = SQL & " sum(rhisfruta_clasif.kilosnet) as kilos "
    SQL = SQL & " FROM  (" & ctabla & ") inner join tmpvarie on rhisfruta.codvarie = tmpvarie.codvarie "
    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6 "
    SQL = SQL & " having sum(rhisfruta_clasif.kilosnet) <> 0 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6 "


    Nregs = TotalRegistrosConsulta(SQL)
    
    Label2(10).Caption = "Cargando Tabla Temporal"
    Me.Pb1.visible = True
    Me.Pb1.Max = Nregs
    Me.Pb1.Value = 0
    Me.Refresh

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    
    While Not RS.EOF
    
        Label2(12).Caption = "Socio " & RS!Codsocio & " Variedad " & RS!codvarie & "-" & RS!codcalid & " Campo " & RS!codcampo
        IncrementarProgresNew Pb1, 1
        Me.Refresh
        DoEvents
    
        Sql3 = "select fechaini, fechafin, max(contador) as contador from rprecios where codvarie = " & DBSet(RS!codvarie, "N")
        Sql3 = Sql3 & " and tipofact = 2 "
        Sql3 = Sql3 & " and fechaini <= " & DBSet(RS!Fecalbar, "F")
        Sql3 = Sql3 & " and fechafin >= " & DBSet(RS!Fecalbar, "F")
        Sql3 = Sql3 & " group by 1,2"
                
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS1.EOF Then
            Contador = DBLet(RS1!Contador, "N")
            FechaIni = DBLet(RS1!FechaIni, "F")
            FechaFin = DBLet(RS1!FechaFin, "F")
        End If
        Set RS1 = Nothing
        
        Sql3 = "select precioindustria from rprecios where codvarie = " & DBSet(RS!codvarie, "N")
        Sql3 = Sql3 & " and tipofact = 2 "
        Sql3 = Sql3 & " and contador = " & DBSet(Contador, "N")
        
        Precio = DevuelveValor(Sql3)
        
        
        Sql4 = "select sum(if(isnull(importe),0,importe)) as gastos"
        Sql4 = Sql4 & "  from rhisfruta, rhisfruta_gastos "
        Sql4 = Sql4 & " where rhisfruta.codsocio = " & DBSet(RS!Codsocio, "N") & "  and "
        Sql4 = Sql4 & " rhisfruta.codvarie = " & DBSet(RS!codvarie, "N") & "  and "
        Sql4 = Sql4 & " rhisfruta.codcampo = " & DBSet(RS!codcampo, "N") & " and "
        Sql4 = Sql4 & " rhisfruta.fecalbar >= " & DBSet(FechaIni, "F") & " and "
        Sql4 = Sql4 & " rhisfruta.fecalbar <= " & DBSet(FechaFin, "F") & " and "
        Sql4 = Sql4 & " rhisfruta.tipoentr = 3 and "
        Sql4 = Sql4 & " rhisfruta.numalbar = rhisfruta_gastos.numalbar "
         
        Gastos = DevuelveValor(Sql4)
        
        '[Monica]23/05/2013: Catadau pasa a tener entradas de industria
        If vParamAplic.Cooperativa = 0 Then
            Sql4 = "select sum(if(isnull(imptrans),0,imptrans)) + sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) as gastos"
            Sql4 = Sql4 & "  from rhisfruta "
            Sql4 = Sql4 & " where rhisfruta.codsocio = " & DBSet(RS!Codsocio, "N") & "  and "
            Sql4 = Sql4 & " rhisfruta.codvarie = " & DBSet(RS!codvarie, "N") & "  and "
            Sql4 = Sql4 & " rhisfruta.codcampo = " & DBSet(RS!codcampo, "N") & " and "
            Sql4 = Sql4 & " rhisfruta.fecalbar >= " & DBSet(FechaIni, "F") & " and "
            Sql4 = Sql4 & " rhisfruta.fecalbar <= " & DBSet(FechaFin, "F") & " and "
            Sql4 = Sql4 & " rhisfruta.tipoentr = 3"
                        
            Gastos = DevuelveValor(Sql4)
        End If
        
        Sql5 = "select count(*) from tmpliquidacion1 where codsocio = " & DBSet(RS!Codsocio, "N") & "  and "
        Sql5 = Sql5 & " tmpliquidacion1.codvarie = " & DBSet(RS!codvarie, "N") & "  and "
        Sql5 = Sql5 & " tmpliquidacion1.codcampo = " & DBSet(RS!codcampo, "N") & " and "
        Sql5 = Sql5 & " tmpliquidacion1.fechaini = " & DBSet(FechaIni, "F") & " and "
        Sql5 = Sql5 & " tmpliquidacion1.fechafin = " & DBSet(FechaFin, "F") & " and "
        Sql5 = Sql5 & " tmpliquidacion1.codusu = " & vUsu.Codigo
        
        If TotalRegistros(Sql5) = 0 Then
            Sql5 = "insert into tmpliquidacion1 values (" & vUsu.Codigo & "," & DBSet(RS!Codsocio, "N") & ","
            Sql5 = Sql5 & DBSet(RS!codvarie, "N") & ","
            Sql5 = Sql5 & DBSet(RS!codcampo, "N") & ","
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
            Sql2 = Sql2 & " and codsocio = " & DBSet(RS!Codsocio, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(RS!codcampo, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and codcalid = " & DBSet(RS!codcalid, "N")
            Sql2 = Sql2 & " and contador = " & DBSet(Contador, "N")
            Sql2 = Sql2 & " and fechaini = " & DBSet(FechaIni, "F")
            Sql2 = Sql2 & " and fechafin = " & DBSet(FechaFin, "F")
            
            If TotalRegistros(Sql2) = 0 Then
                Kilos = 0
                
                Sql3 = "insert into tmpliquidacion (codusu,codsocio,codcampo,codvarie,codcalid,contador,kilosnet,precio,importe, "
                Sql3 = Sql3 & " nomvarie, fechaini, fechafin, gastos)"
                Sql3 = Sql3 & " values (" & vUsu.Codigo & "," & DBSet(RS!Codsocio, "N") & "," & DBSet(RS!codcampo, "N") & ","
                Sql3 = Sql3 & DBSet(RS!codvarie, "N") & "," & DBSet(RS!codcalid, "N") & "," & DBSet(Contador, "N") & ","
                Sql3 = Sql3 & DBSet(RS!Kilos, "N") & "," & DBSet(Precio, "N") & "," & DBSet(0, "N") & ","
                Sql3 = Sql3 & DBSet(RS!nomvarie, "T") & "," & DBSet(FechaIni, "F") & ","
                Sql3 = Sql3 & DBSet(FechaFin, "F") & "," & DBSet(Gastos, "N") & ")"
                
                conn.Execute Sql3
                Kilos = RS!Kilos
            Else
                Kilos = Kilos + RS!Kilos
                Sql3 = "update tmpliquidacion set kilosnet = kilosnet + " & DBSet(RS!Kilos, "N")
                Sql3 = Sql3 & " where codusu = " & vUsu.Codigo
                Sql3 = Sql3 & " and codsocio = " & DBSet(RS!Codsocio, "N")
                Sql3 = Sql3 & " and codcampo = " & DBSet(RS!codcampo, "N")
                Sql3 = Sql3 & " and codvarie = " & DBSet(RS!codvarie, "N")
                Sql3 = Sql3 & " and codcalid = " & DBSet(RS!codcalid, "N")
                Sql3 = Sql3 & " and contador = " & DBSet(Contador, "N")
                Sql3 = Sql3 & " and fechaini = " & DBSet(FechaIni, "F")
                Sql3 = Sql3 & " and fechafin = " & DBSet(FechaFin, "F")
                
                conn.Execute Sql3
            End If
'30/07/2009
'        End If
        
        RS.MoveNext
    Wend
                                    
    Sql3 = "update tmpliquidacion set importe = round(kilosnet * precio,2) where codusu = " & vUsu.Codigo
    conn.Execute Sql3
                                    
                                    
    Me.Pb1.visible = False
    Me.Label2(10).Caption = ""
    Me.Label2(12).Caption = ""
    Me.Refresh
                                    
    CargarTemporalIndustria = True
    Exit Function
    
eCargarTemporal:
    Me.Pb1.visible = False
    Me.Label2(10).Caption = ""
    Me.Label2(12).Caption = ""
    Me.Refresh
    
    MuestraError "Cargando temporal Industria", Err.Description
End Function

Private Sub CertificadoRetencionesVisible()
' si se trata de un certificado de retenciones no tiene sentido el check de salta página por socio
    If Check1(7).Value = 1 Then
        Check1(6).Enabled = False
        Check1(6).Value = False
        Check1(4).Enabled = False
        Check1(4).Value = False
        Check1(9).Enabled = False
        Check1(9).Value = False
        Check1(20).Enabled = False
        Check1(20).Value = False
    
    Else
        Check1(6).Enabled = True
        Check1(4).Enabled = True
        Check1(9).Enabled = True
        Check1(20).Enabled = True
    End If
    FrameFechaCertif.visible = (Check1(7).Value = 1)
    FrameFechaCertif.Enabled = (Check1(7).Value = 1)
    
End Sub



Private Sub AportacionesFondoOperativoVisible()
' si se trata de un certificado de retenciones no tiene sentido el check de salta página por socio
    If Check1(9).Value = 1 Then
        Check1(6).Enabled = False
        Check1(6).Value = False
        Check1(4).Enabled = False
        Check1(4).Value = False
        Check1(7).Enabled = False
        Check1(7).Value = False
        Check1(20).Enabled = False
        Check1(20).Value = False
    Else
        Check1(6).Enabled = True
        Check1(4).Enabled = True
        Check1(7).Enabled = True
        Check1(20).Enabled = True
    End If
    FrameFechaCertif.visible = False
    FrameFechaCertif.Enabled = False
    
End Sub

Private Sub EpigrafeVisible()
' si se trata de un informe de epigrafe visible puedo o no saltar por socio
    If Check1(20).Value = 1 Then
        Check1(6).Enabled = False
        Check1(6).Value = False
        Check1(7).Enabled = False
        Check1(7).Value = False
        Check1(4).Enabled = False
        Check1(4).Value = False
        Check1(9).Enabled = False
        Check1(9).Value = False
    Else
        Check1(6).Enabled = True
        Check1(7).Enabled = True
        Check1(4).Enabled = True
        Check1(9).Enabled = True
    End If
    FrameFechaCertif.visible = (Check1(7).Value = 1)
    FrameFechaCertif.Enabled = (Check1(7).Value = 1)
    
End Sub



Private Function TipoFacturaOk() As Boolean
Dim b As Boolean
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
Dim OtrosTipos As Boolean
Dim Tipos As String
Dim i As Integer
    
    
    b = True
    'Tipo de movimiento: INDUSTRIA
    Tipos = ""
    Industria = False
    For i = 1 To ListView1(0).ListItems.Count
        If ListView1(0).ListItems(i).Checked Then
            Tipos = Tipos & DBSet(ListView1(0).ListItems(i).Key, "T") & ","
        End If
    Next i
    If Len(Tipos) > 0 Then Tipos = Mid(Tipos, 1, Len(Tipos) - 1)
    If InStr(1, Tipos, "FLI") Then
        If Len(Tipos) > 5 Then
            MsgBox "Si selecciona las facturas de industria no puede meter más tipos de factura.", vbExclamation
            b = False
        Else
            Industria = True
        End If
    End If
    
    ' Tipo de movimiento: BODEGA / ALMAZARA
    ' si selecciona facturas de bodega/almazara, únicamente de bodega/almazara.
    Bodega = False
    OtrosTipos = False
    If b Then
        For i = 1 To ListView1(0).ListItems.Count
            If ListView1(0).ListItems(i).Checked Then 'And Mid(ListView1(0).ListItems(i).Key, 3, 1) = "B" Then
                Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(ListView1(0).ListItems(i).Key, "T"))
                If Tipo >= 7 And Tipo <= 10 Then
                    Bodega = True
                Else
                    OtrosTipos = True
                End If
            End If
        Next i
        
        If Bodega And OtrosTipos Then
            MsgBox "Si selecciona las facturas de bodega/almazara no puede meter más tipos de factura.", vbExclamation
            b = False
        End If
    End If
            
    TipoFacturaOk = b

End Function


Private Sub PonerCamposSocio()
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim RS As ADODB.Recordset


    If txtcodigo(49).Text = "" Then Exit Sub
    
    Cad = "rcampos.codsocio = " & DBSet(txtcodigo(49).Text, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select count(*) from rcampos where " & Cad
     
    NumRegis = TotalRegistros(Cad1)
    
    If NumRegis = 0 Then Exit Sub
    If NumRegis = 1 Then
        Cad1 = "select codcampo from rcampos where " & Cad
        Set RS = New ADODB.Recordset
        RS.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            txtcodigo(50).Text = DBLet(RS.Fields(0).Value)
            PonerDatosCampo txtcodigo(50).Text
        End If
    Else
        Set frmMens1 = New frmMensajes
        frmMens1.cadWhere = " and " & Cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens1.campo = txtcodigo(50).Text
        frmMens1.OpcionMensaje = 7
        frmMens1.Show vbModal
        Set frmMens1 = Nothing
    End If
    
End Sub


Private Sub PonerDatosCampo(campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim RS As ADODB.Recordset


    If campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla, rcampos.nrocampo from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set RS = New ADODB.Recordset
    RS.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
'    Text1(5).Text = ""
    Text2(0).Text = ""
    Text2(4).Text = ""
    Text2(3).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    Text2(5).Text = ""
    If Not RS.EOF Then
        txtcodigo(50).Text = campo
        PonerFormatoEntero txtcodigo(50)
        Text2(4).Text = DBLet(RS.Fields(0).Value, "N") ' codigo de partida
        If Text2(4).Text <> "" Then Text2(4).Text = Format(Text2(4).Text, "0000")
        Text2(3).Text = DBLet(RS.Fields(1).Value, "T") ' nombre de partida
        Text2(1).Text = DBLet(RS.Fields(2).Value, "N") ' codigo de zona
        If Text2(1).Text <> "" Then Text2(1).Text = Format(Text2(1).Text, "0000")
        Text2(2).Text = DBLet(RS.Fields(3).Value, "T") ' nombre de zona
        Text2(5).Text = DBLet(RS.Fields(4).Value, "T") ' descripcion de poblacion
        Text2(0).Text = DBLet(RS.Fields(5).Value, "N") ' Nro de campo
    End If
    
    Set RS = Nothing
    
End Sub

 

Private Function RecalculoImportes(Albar As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ImpTotal As Currency
Dim Albaran As Long
Dim KilosNet As Long
Dim KilosTot As Long
Dim ImporteTotal As Currency
Dim Importe As Currency

    On Error GoTo eRecalculoImportes

    RecalculoImportes = False

    If Not BloqueaRegistro("rhisfruta", "numalbar in (" & Albar & ")") Then
        MsgBox "No se pueden actualizar Entradas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        conn.BeginTrans

        ' Kilos Totales
        SQL = "select sum(kilosnet) from rhisfruta where numalbar in (" & Trim(Albar) & ")"
        KilosTot = DevuelveValor(SQL)
    
        ImporteTotal = txtcodigo(46).Text
    
        SQL = "select numalbar, kilosnet from rhisfruta where numalbar in (" & Trim(Albar) & ")"
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        ImpTotal = 0
        
        While Not RS.EOF
            Albaran = DBLet(RS!numalbar, "N")
            Importe = Round2(RS!KilosNet * ImporteTotal / KilosTot, 2)
            
            ImpTotal = ImpTotal + Importe
            
            ' actualizamos la entrada
            SQL = "update rhisfruta set impentrada = " & DBSet(Importe, "N")
            SQL = SQL & " where numalbar = " & DBSet(Albaran, "N")
            
            conn.Execute SQL
            
            RS.MoveNext
        Wend
        Set RS = Nothing
        
        If ImpTotal <> ImporteTotal And ImpTotal <> 0 Then
            SQL = "update rhisfruta set impentrada = impentrada + " & DBSet(ImporteTotal - ImpTotal, "N")
            SQL = SQL & " where numalbar = " & DBSet(Albaran, "N")
        
            conn.Execute SQL
        End If
        
        RecalculoImportes = True
        conn.CommitTrans
        Exit Function
        
    End If
    
eRecalculoImportes:
    MuestraError Err.Number, "Recalculo de Importes", Err.Description
    conn.RollbackTrans
    TerminaBloquear
End Function



Private Function CargarTemporalLiquidacionesCalidadPicassent(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CalidAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bonifica As Currency
Dim Importe As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

Dim PorcBoni As Currency
Dim PrecioAnt As Currency
Dim PorcComi As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionesCalidadPicassent = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    SQL = "SELECT rhisfruta.codsocio,  rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid, rhisfruta.fecalbar, "
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilos "
     SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilos "
    
    SQL = SQL & " FROM  " & ctabla
    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "


    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                    'codusu,  codvarie, nomvarie, calidad, Kneto,  Precio, importe, bonificacion,
    Sql2 = "insert into tmpinformes (codusu,  importe1, nombre1, campo1, importe2, precio1, importe3, importe4, "
                   'importetotal
    Sql2 = Sql2 & " importe5) values "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    If Not RS.EOF Then
        VarieAnt = RS!codvarie
        NVarieAnt = RS!nomvarie
        CalidAnt = RS!codcalid
    End If
    
    While Not RS.EOF
        '++monica:28/07/2009 añadida la segunda condicion
        If VarieAnt <> RS!codvarie Or CalidAnt <> RS!codcalid Then
            Sql1 = Sql1 & "(" & vUsu.Codigo & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(CalidAnt, "N") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(PrecioAnt, "N") & ","
            Sql1 = Sql1 & DBSet(Importe, "N") & "," & DBSet(Bonifica - Importe, "N") & ","
            Sql1 = Sql1 & DBSet(Bonifica, "N") & "),"
            
            VarieAnt = RS!codvarie
            CalidAnt = RS!codcalid
            
            baseimpo = 0
            Bonifica = 0
            Importe = 0
            Neto = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            KilosNet = 0
        End If
        
        KilosNet = KilosNet + DBLet(RS!Kilos, "N")
        
        Recolect = DBLet(RS!Recolect, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(RS!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(RS!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(RS!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
            PorcBoni = 0
            PorcComi = 0
            Select Case Recolect
                Case 0
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreCoop > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                        
                        '[Monica]22/03/2012: Solo si le indicamos que no calcule comision no lo hace (solo prevision de liquidacion)
                        If Check1(13).Value Then
                            '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                            PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                            If CCur(PorcComi) <> 0 Then
                                PreCoop = PreCoop - Round2(PreCoop * PorcComi / 100, 4)
                            End If
                        End If
                    End If
                    PrecioAnt = PreCoop
                    Importe = Importe + Round2(DBLet(RS!Kilos, "N") * PreCoop, 2)
                    Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * PreCoop, 2) + Round2(DBLet(RS!Kilos, "N") * PreCoop * PorcBoni / 100, 2)
                Case 1
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreSocio > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                        
                        '[Monica]22/03/2012: Solo si le indicamos que no calcule comision no lo hace (solo prevision de liquidacion)
                        If Check1(13).Value Then
                            '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                            PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                            If CCur(PorcComi) <> 0 Then
                                PreSocio = PreSocio - Round2(PreSocio * PorcComi / 100, 4)
                            End If
                        End If
                    End If
                    PrecioAnt = PreSocio
                    Importe = Importe + Round2(DBLet(RS!Kilos, "N") * PreSocio, 2)
                    Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * PreSocio, 2) + Round2(DBLet(RS!Kilos, "N") * PreSocio * PorcBoni / 100, 2)
            End Select
            
        End If
        Set Rs9 = Nothing
        'hasta aqui
        
        
        
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If HayReg Then
        Sql1 = Sql1 & "(" & vUsu.Codigo & ","
        Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
        Sql1 = Sql1 & DBSet(CalidAnt, "N") & ","
        Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(PrecioAnt, "N") & ","
        Sql1 = Sql1 & DBSet(Importe, "N") & "," & DBSet(Bonifica - Importe, "N") & ","
        Sql1 = Sql1 & DBSet(Bonifica, "N") & "),"
        
        ' quitamos la ultima coma e insertamos
        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        conn.Execute Sql2 & Sql1
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
    CargarTemporalLiquidacionesCalidadPicassent = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargando temporal", Err.Description
End Function


Private Function InsertarTablaIntermedia(ctabla As String, cwhere As String, ConCampo As Boolean) As Boolean
Dim Sql2 As String
Dim SQL As String
Dim SqlTempo As String
Dim KilosEntrados As Long
Dim KilosRetirados As Long
Dim TKilosEntrados As Long
Dim TKilosRetirados As Long
Dim Kilos As Long
Dim RS As ADODB.Recordset
Dim SocioAnt As String
Dim VarieAnt As String
Dim CampoAnt As String

    On Error GoTo eInsertarTablaIntermedia

    InsertarTablaIntermedia = False


    Sql2 = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    conn.Execute Sql2


    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    If ConCampo Then
        SQL = "SELECT " & vUsu.Codigo & ", rclasifica.codsocio, rclasifica.codvarie, variedades.nomvarie, rclasifica.codcampo, "
    Else
        SQL = "SELECT " & vUsu.Codigo & ", rclasifica.codsocio, rclasifica.codvarie, variedades.nomvarie, "
    End If
    SQL = SQL & "sum(rclasifica.kilosnet) as kilos"
    SQL = SQL & " FROM  " & ctabla

    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    If ConCampo Then
        SQL = SQL & " group by 1, 2, 3, 4, 5 "
        SQL = SQL & " order by 1, 2, 3, 4, 5 "
    Else
        SQL = SQL & " group by 1, 2, 3, 4 "
        SQL = SQL & " order by 1, 2, 3, 4 "
    End If
    
    If ConCampo Then
        SqlTempo = "insert into tmpliquidacion (codusu, codsocio, codvarie, nomvarie, codcampo, kilosnet) "
    Else
        SqlTempo = "insert into tmpliquidacion (codusu, codsocio, codvarie, nomvarie, kilosnet) "
    End If
    SqlTempo = SqlTempo & SQL
    conn.Execute SqlTempo
    
    If ConCampo Then
        SqlTempo = "insert into tmpliquidacion (codusu, codsocio, codvarie, nomvarie, codcampo, kilosnet) "
    Else
        SqlTempo = "insert into tmpliquidacion (codusu, codsocio, codvarie, nomvarie, kilosnet) "
    End If
    SqlTempo = SqlTempo & Replace(Replace(SQL, "rclasifica", "rhisfruta"), "fechaent", "fecalbar")
    conn.Execute SqlTempo

    '[Monica]19/10/2011: si la factura es de retirada los kilos deben coincidir con los kilos de Retirada
    If Check1(12).Value = 1 Then
        'comprobamos que los kilos retirados es una cantidad inferior a la que hay de kilos entrados
        TKilosRetirados = CLng(ImporteSinFormato(txtcodigo(59).Text))
        TKilosEntrados = DevuelveValor("select sum(kilosnet) from tmpliquidacion where codusu = " & vUsu.Codigo)
        If TKilosEntrados < TKilosRetirados Then
            MsgBox "Los kilos de Retirada son superiores a los entrados. Revise.", vbExclamation
            InsertarTablaIntermedia = False
            Exit Function
        End If
    
        KilosRetirados = 0
    
        ' prorrateamos los kilos
        SQL = "select * from tmpliquidacion where codusu = " & vUsu.Codigo & " order by codsocio"
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            Kilos = Round2(DBLet(RS!KilosNet, "N") * TKilosRetirados / TKilosEntrados, 0)
            KilosRetirados = KilosRetirados + Kilos
            
            SocioAnt = RS!Codsocio
            VarieAnt = RS!codvarie
            CampoAnt = RS!codcampo
            
            Sql2 = "update tmpliquidacion set kilosnet = " & DBSet(Kilos, "N")
            Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
            Sql2 = Sql2 & " and codsocio = " & DBSet(RS!Codsocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(RS!codvarie, "N")
            If ConCampo Then
                Sql2 = Sql2 & " and codcampo = " & DBSet(RS!codcampo, "N")
            End If
            
            conn.Execute Sql2
            
            RS.MoveNext
        Wend
        Set RS = Nothing
    
        If KilosRetirados <> TKilosRetirados Then
            Sql2 = "update tmpliquidacion set kilosnet = kilosnet + " & DBSet(TKilosRetirados - KilosRetirados, "N")
            Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
            Sql2 = Sql2 & " and codsocio = " & DBSet(SocioAnt, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(VarieAnt, "N")
            If ConCampo Then
                Sql2 = Sql2 & " and codcampo = " & DBSet(CampoAnt, "N")
            End If
            conn.Execute Sql2
        End If
    End If ' fin del prorrateo de kilos retirados
    
    InsertarTablaIntermedia = True
    Exit Function

eInsertarTablaIntermedia:
    MuestraError Err.Number, "Insertar Tabla Intermedia", Err.Description
End Function


Private Sub KilosRetiradaVisible()
' si se trata de un anticipo de retirada
    If Check1(12).Value Then
        txtcodigo(59).Enabled = True
        '[Monica]23/12/2014: veto ruso
        Check1(22).Enabled = True
        

    Else
        txtcodigo(59).Enabled = False
        '[Monica]23/12/2014: veto ruso
        Check1(22).Enabled = False
        Check1(22).Value = 0
    End If
End Sub


Private Function HayVariedadesAlmazaraconHorto(mTabla As String, mSelect As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Totales As Integer

    SQL = "select distinct rhisfruta.codvarie from " & mTabla
    If mSelect <> "" Then SQL = SQL & " where " & Replace(Replace(mSelect, "{", ""), "}", "")
    
    Sql2 = "select count(*) from productos where codgrupo = 5 and codprodu in (select codprodu from variedades where codvarie in (" & SQL & ")) "
    Totales = TotalRegistros(Sql2)
    Sql2 = "select count(*) from productos where codgrupo <> 5 and codprodu in (select codprodu from variedades where codvarie in (" & SQL & ")) "
    If TotalRegistros(Sql2) > 0 Then Totales = Totales + 1

    HayVariedadesAlmazaraconHorto = (Totales = 2)

End Function

Private Function HayVariedadesAlmazara(mTabla As String, mSelect As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Totales As Integer

    SQL = "select distinct rhisfruta.codvarie from " & mTabla
    If mSelect <> "" Then SQL = SQL & " where " & Replace(Replace(mSelect, "{", ""), "}", "")
    
    Sql2 = "select count(*) from productos where codgrupo = 5 and codprodu in (select codprodu from variedades where codvarie in (" & SQL & ")) "
    Totales = TotalRegistros(Sql2)

    HayVariedadesAlmazara = (Totales = 1)

End Function

Private Function AlbaranesFacturados(ctabla As String, cwhere As String, Optional CadenaAlbaranes As String) As Boolean
Dim SQL As String
Dim Cadena As String
Dim Cadena2 As String
Dim RS As ADODB.Recordset
    
    On Error GoTo eAlbaranesFacturados
    
    AlbaranesFacturados = True
    
    SQL = "select rfactsoc_albaran.numalbar, rfactsoc_albaran.fecalbar "
    SQL = SQL & " from rfactsoc_albaran "
    SQL = SQL & " where numalbar in (select rhisfruta.numalbar from " & ctabla & " where " & cwhere & ")"
    SQL = SQL & " order by 1"
            
    If TotalRegistros(SQL) > 0 Then
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        Cadena = ""
    
        While Not RS.EOF
            Cadena = Cadena & Format(DBLet(RS!numalbar, "N"), "0000000") & ", "
        
            RS.MoveNext
        Wend
        Set RS = Nothing
        
        Set frmMens4 = New frmMensajes
        
        frmMens4.OpcionMensaje = 35
        frmMens4.cadWhere = "rhisfruta.numalbar in (" & Mid(Cadena, 1, Len(Cadena) - 2) & ")"
        frmMens4.Show vbModal
        
        Set frmMens4 = Nothing
        
        Select Case vReturn
            Case 0
                ' indicamos como si no hubieran albaranes facturados para poder continuar con el proceso
                ' de liquidacion o de anticipos
                AlbaranesFacturados = True
            
            Case 1
                ' se liquidan todos los albaranes no facturados
                AlbaranesFacturados = True

                cwhere = cwhere & " and rhisfruta.numalbar not in (" & Mid(Cadena, 1, Len(Cadena) - 2) & ")"
            
                CadenaAlbaranes = "not rhisfruta.numalbar  in (" & Mid(Cadena, 1, Len(Cadena) - 2) & ")"
            
            Case 2
                ' abortamos el proceso
                AlbaranesFacturados = False
        
        End Select
    End If
    Exit Function
    
eAlbaranesFacturados:
    AlbaranesFacturados = False
    MensError = "Albaranes Facturados"
    MuestraError Err.Number, MensError
End Function


'****************************************************************
'*******ANTIGUAS CAMPAÑAS ANTERIORES EN EL LISTADO DE RETENCIONES
'****************************************************************
'
'    ' insertamos las facturas correspondientes a la campaña anterior
'    Set vCampAnt = New CCampAnt
'    If vCampAnt.Leer(True) = 0 Then
'        If AbrirConexionCampAnterior(vCampAnt.BaseDatos) Then
'
'            ' borramos la tabla temporal de la campaña anterior
'            SQL = "delete from tmprfactsoc where codusu= " & vUsu.Codigo
'            ConnCAnt.Execute SQL
'
''            ' borramos la tabla temporal de la campaña anterior
''            sql = "delete from tmprfacttra where codusu= " & vUsu.Codigo
''            ConnCAnt.Execute sql
'
'
'            SQL = "insert into tmprfactsoc (`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
'            SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
'            SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, tipo, esretirada) "
'            SQL = SQL & "select " & vUsu.Codigo & ", `rfactsoc`.`codtipom`, `rfactsoc`.`numfactu`, `rfactsoc`.`fecfactu`, `codsocio`, "
'            SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
'            SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, 0, `esretirada` " '[Monica]24/07/2012: metemos si es de retirada
'            SQL = SQL & " from " & cTabla
'            SQL = SQL & " where " & cSelect
'            SQL = SQL & " group by 1,2,3,4 "
'            ConnCAnt.Execute SQL
'
'
'            '[Monica]26/08/2011: Modificacion solo para Picassent
'            '                    en las facturas de socios quiere que en la columna impapor estén tb los descuentos,
'            '                    con lo cual el totalfac será el total a pagar
'            '
'            If vParamAplic.Cooperativa = 2 Then
'                SQL = "update tmprfactsoc set impapor = if(impapor is null,0,impapor) + (select if(sum(importe) is null,0,sum(importe)) from rfactsoc_gastos where "
'                SQL = SQL & " rfactsoc_gastos.codtipom = tmprfactsoc.codtipom and rfactsoc_gastos.numfactu = tmprfactsoc.numfactu "
'                SQL = SQL & " and rfactsoc_gastos.fecfactu = tmprfactsoc.fecfactu) "
'                SQL = SQL & " where tmprfactsoc.codusu = " & vUsu.Codigo & " and tmprfactsoc.tipo = 0"
'
'                ConnCAnt.Execute SQL
'
'                ' ahora el total factura es el total a pagar en Picassent
'                SQL = "update tmprfactsoc set totalfac = baseimpo + if(imporiva is null,0,imporiva) - if(impreten is null,0,impreten) - if(impapor is null,0,impapor)  "
'                SQL = SQL & " where tmprfactsoc.codusu = " & vUsu.Codigo & " and tmprfactsoc.tipo = 0"
'
'                ConnCAnt.Execute SQL
'            End If
'
'
'            '[Monica]03/03/2011: cargamos el nif del socio
'            SQL = "update tmprfactsoc set nif = (select nifsocio from rsocios where rsocios.codsocio = tmprfactsoc.codsocio)"
'            SQL = SQL & " where tipo = 0"
'            ConnCAnt.Execute SQL
'
'            '[Monica]03/03/2011: cargamos el nif del socio
'            SQL = "update tmprfactsoc set nomsocio = (select nomsocio from rsocios where rsocios.codsocio = tmprfactsoc.codsocio)"
'            SQL = SQL & " where tipo = 0"
'            ConnCAnt.Execute SQL
'
'            If InStr(1, cSelect, "FTR") Or (cSelect2 <> "" And (OpcionListado = 10 Or OpcionListado = 11)) Then
'                SQL = "insert into tmprfactsoc (`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
'                SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
'                SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, tipo) "
'                SQL = SQL & "select " & vUsu.Codigo & ", `rfacttra`.`codtipom`, `rfacttra`.`numfactu`, `rfacttra`.`fecfactu`, `rfacttra`.`codtrans`, "
'                SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,`basereten`,"
'                SQL = SQL & "`porc_ret`,`impreten`,`baseaport`,`porc_apo`,`impapor`,`totalfac`, 1 "
'                SQL = SQL & " from " & cTabla2
'                SQL = SQL & " where " & cSelect2
'                SQL = SQL & " group by 1,2,3,4,5 "
'                ConnCAnt.Execute SQL
'
'                '[Monica]03/03/2011: cargamos el nif del transportista
'                SQL = "update tmprfactsoc set nif = (select niftrans from rtransporte where rtransporte.codtrans = tmprfactsoc.codsocio)"
'                SQL = SQL & " where tipo = 1"
'                ConnCAnt.Execute SQL
'
'                '[Monica]03/03/2011: cargamos el nif del transportista
'                SQL = "update tmprfactsoc set nomsocio = (select nomtrans from rtransporte where rtransporte.codtrans = tmprfactsoc.codsocio)"
'                SQL = SQL & " where tipo = 1"
'                ConnCAnt.Execute SQL
'
'
'            End If
'
'
'            If OpcionListado = 11 Then ' caso del 346
'                SQL = "delete from tmp346 where codusu= " & vUsu.Codigo
'                ConnCAnt.Execute SQL
'
'                ctabla1 = "(" & cTabla & ") INNER JOIN variedades ON rfactsoc_variedad.codvarie = variedades.codvarie "
'                ctabla1 = "(" & ctabla1 & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
'                ctabla1 = "(" & ctabla1 & ") INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
'
'                SQL = "insert into tmp346 (`codusu`,`codsocio`,`codgrupo`,`importe`) "
'                SQL = SQL & " select " & vUsu.Codigo & ", rfactsoc.codsocio, grupopro.codgrupo, sum(rfactsoc_variedad.imporvar) "
'                SQL = SQL & " from " & ctabla1
'                SQL = SQL & " where " & cSelect & " and grupopro.codgrupo in (4,5) " ' algarrobos y olivos
'                SQL = SQL & " group by rfactsoc.codsocio, grupopro.codgrupo  "
'                SQL = SQL & " union "
'                SQL = SQL & " select " & vUsu.Codigo & ", rfactsoc.codsocio, 0, sum(rfactsoc_variedad.imporvar)"
'                SQL = SQL & " from " & ctabla1
'                SQL = SQL & " where " & cSelect & " and not grupopro.codgrupo in (4,5) " ' el resto
'                SQL = SQL & " group by rfactsoc.codsocio, grupopro.codgrupo  "
'                SQL = SQL & " order by 1,2 "
'
'                ConnCAnt.Execute SQL
'            End If
'
'            If OpcionListado = 8 Or OpcionListado = 9 Then
'                SQL = "delete from tmprfactsoc_variedad where codusu= " & vUsu.Codigo
'                ConnCAnt.Execute SQL
'
'
''                sql = "delete from tmprfacttra_variedad where codusu= " & vUsu.Codigo
''                ConnCAnt.Execute sql
''
'
'                SQL = "insert into tmprfactsoc_variedad (`codusu`,`codtipom`,`numfactu`,`fecfactu`,"
'                SQL = SQL & "`codvarie`,`codcampo`,`kilosnet`,`preciomed`,`imporvar`,`descontado`) "
'                SQL = SQL & " select " & vUsu.Codigo & ", rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu,"
'                SQL = SQL & " rfactsoc_variedad.codvarie, rfactsoc_variedad.codcampo, rfactsoc_variedad.kilosnet,"
'                SQL = SQL & " rfactsoc_variedad.preciomed, rfactsoc_variedad.imporvar, rfactsoc_variedad.descontado "
'                SQL = SQL & " from " & cTabla
'                SQL = SQL & " where " & cSelect
'
'                ConnCAnt.Execute SQL
'
'                If Check1(7).Value = 1 Then ' caso del certificado de retenciones
'                    SQL = "insert into tmprfactsoc_variedad (`codusu`,`codtipom`,`numfactu`,`fecfactu`,"
'                    SQL = SQL & "`codvarie`,`codcampo`,`kilosnet`,`preciomed`,`imporvar`,`descontado`) "
'                    SQL = SQL & " select distinct " & vUsu.Codigo & ", rfactsoc_anticipos.codtipom, rfactsoc_anticipos.numfactu, rfactsoc_anticipos.fecfactu,"
'                    SQL = SQL & " rfactsoc_anticipos.codvarieanti, rfactsoc_anticipos.codcampoanti, 0,0,rfactsoc_anticipos.baseimpo * (-1),0 "
'                    SQL = SQL & " from (" & cTabla & ") Inner join rfactsoc_anticipos on rfactsoc.codtipom = rfactsoc_anticipos.codtipom and rfactsoc.numfactu = rfactsoc_anticipos.numfactu and rfactsoc.fecfactu = rfactsoc_anticipos.fecfactu "
'                    SQL = SQL & " where " & cSelect
'
'                    ConnCAnt.Execute SQL
'
'                End If
'
'
'
'
'                If InStr(1, cSelect, "FTR") Then
'                    SQL = "insert into tmprfactsoc_variedad (`codusu`,`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
'                    SQL = SQL & "`codvarie`,`codcampo`,`kilosnet`,`preciomed`,`imporvar`,`descontado`) "
'                    SQL = SQL & " select " & vUsu.Codigo & ", rfacttra.codtipom, rfacttra.numfactu, rfacttra.fecfactu, rfacttra.codtrans,"
'                    SQL = SQL & " rfacttra_albaran.codvarie, rfacttra_albaran.codcampo, sum(rfacttra_albaran.kilosnet) kilosnet,"
'                    SQL = SQL & " 0, sum(rfacttra_albaran.importe) importe , 0 "
'                    SQL = SQL & " from " & cTabla2
'                    SQL = SQL & " where " & cSelect2
'                    SQL = SQL & " group by 1,2,3,4,5,6,7,9,11 "
'                    ConnCAnt.Execute SQL
'                End If
'
'            End If
'
'        End If
'
'        ' introducimos las facturas de la campaña anterior en la temporal de la
'        ' campaña actual
'        SQL = "insert into tmprfactsoc select * from " & vCampAnt.BaseDatos & ".tmprfactsoc "
'        SQL = SQL & " where codusu = " & vUsu.Codigo
'
'        conn.Execute SQL
'
'
''        ' introducimos las facturas de la campaña anterior en la temporal de la
''        ' campaña actual
''        sql = "insert into tmprfacttra select * from " & vCampAnt.BaseDatos & ".tmprfacttra "
''        sql = sql & " where codusu = " & vUsu.Codigo
''
''        conn.Execute sql
'
'
'        If OpcionListado = 11 Then
'            SQL = "insert into tmp346 select * from " & vCampAnt.BaseDatos & ".tmp346 "
'            SQL = SQL & " where codusu = " & vUsu.Codigo
'
'            conn.Execute SQL
'        End If
'
'        If OpcionListado = 8 Or OpcionListado = 9 Then
'            SQL = "insert into tmprfactsoc_variedad select * from " & vCampAnt.BaseDatos & ".tmprfactsoc_variedad "
'            SQL = SQL & " where codusu = " & vUsu.Codigo
'
'            conn.Execute SQL
'
''
''            sql = "insert into tmprfacttra_variedad select * from " & vCampAnt.BaseDatos & ".tmprfacttra_variedad "
''            sql = sql & " where codusu = " & vUsu.Codigo
''
''            conn.Execute sql
''
'
'        End If
'
'        CerrarConexionCampAnterior
'
'    End If
'    Set vCampAnt = Nothing


'********************************************************
'********** PROCESO PICASSENT *****************03/10/2013
'********************************************************
Private Sub ProcesoPicassent()
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
Dim b As Boolean
Dim Sql2 As String

Dim MaxContador As String

Dim Tabla1 As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtcodigo(12).Text)
        cHasta = Trim(txtcodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(20).Text)
        cHasta = Trim(txtcodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(20).Text, "N")
        If txtcodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(6).Text)
        cHasta = Trim(txtcodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
            
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        'Socio que no sea tercero
        If Check1(11).Value = 0 Then
            If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} <> 1") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} <> 1") Then Exit Sub
        Else
            ' socio tercero de modulos
            If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} = 1") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} = 1") Then Exit Sub
        End If
        
        'sólo entradas distintas de VENTA CAMPO y distintas de INDUSTRIA y distintas de RETIRADA
        If Not AnyadirAFormula(cadSelect, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3 and {rhisfruta.tipoentr} <> 4") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.tipoentr} <> 1 and {rhisfruta.tipoentr} <> 3 and {rhisfruta.tipoentr} <> 4") Then Exit Sub
        
        '++monica: 28/07/2009 dependiendo del tipo de recoleccion (0=coop 1=socio 2=todos)
        Select Case Combo1(2).ListIndex
            Case 0      ' recolectado cooperativa
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 0") Then Exit Sub
            Case 1      ' recolectado socio
                If Not AnyadirAFormula(cadSelect, "{rhisfruta.recolect} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadFormula, "{rhisfruta.recolect} = 1") Then Exit Sub
            Case 2      ' ambos
            
        End Select
        
        nTabla = "(((((rhisfruta INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio) "
        nTabla = nTabla & " INNER JOIN rhisfruta_clasif ON rhisfruta.numalbar = rhisfruta_clasif.numalbar) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        cadSelect1 = cadSelect
        Tabla1 = nTabla
        
        Select Case OpcionListado
            Case 1 ' Listado de anticipos
                'Nombre fichero .rpt a Imprimir
                indRPT = 24 ' informe de anticipos
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatAnticipos.rpt"
                cadTitulo = "Informe de Anticipos"
            
            Case 2 ' Prevision de pago de anticipos
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    cadNombreRPT = "rPrevPagosAnt.rpt"
                Else
                    If Combo1(3).ListIndex = 1 Then ' agrupado por variedad
                        cadNombreRPT = "rPrevPagosAnt1.rpt"
                    Else ' por calidad
                        cadNombreRPT = "rPrevPagosAnt2.rpt"
                    End If
                End If
                cadTitulo = "Previsión de Pago de Anticipos"
            
            Case 3 ' Facturación de Anticipos
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Anticipos"
            
            Case 12 ' Listado de Liquidacion
                'Nombre fichero .rpt a Imprimir
                indRPT = 26 ' informe de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"CatLiquidacion.rpt"
                cadTitulo = "Informe de Liquidación"
                
            Case 13 ' Prevision de pago de liquidacion
                'Nombre fichero .rpt a Imprimir
                indRPT = 33 ' informe de prevision de pago de liquidacion
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu '"ValPrevPagosLiq.rpt"
                If Combo1(3).ListIndex = 0 Then ' agrupado por socio
                    ' no hacemos nada dejamos el nombre de fichero como estaba
                    
                Else
                    If Combo1(3).ListIndex = 1 Then ' agrupado por variedad
                        cadNombreRPT = Replace(cadNombreRPT, "PrevPagosLiq.rpt", "PrevPagosLiq1.rpt")
                    Else ' por calidad
                        cadNombreRPT = Replace(cadNombreRPT, "PrevPagosLiq.rpt", "PrevPagosLiq2.rpt")
                    End If
                End If
                
                cadTitulo = "Previsión de Pago de Liquidación"
                
            Case 14 ' Facturación de Liquidacion
                cadNombreRPT = "rResumFacturas.rpt"
                
                cadTitulo = "Resumen de Facturas de Liquidación"
                
        End Select
                    
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWhere = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
            Select Case OpcionListado
                Case 1, 2, 3
                    TipoPrec = 0 ' ANTICIPOS
                Case 12, 13, 14
                    TipoPrec = 1 ' LIQUIDACIONES
            End Select
            
            'comprobamos que los tipos de iva existen en la contabilidad de horto
            If Not ComprobarTiposIVA(nTabla, cadSelect) Then Exit Sub
            
            
'            '[Monica]27/04/2011: de momento solo alzira comprobamos si los albaranes seccionado ya estan liquidados
'            If vParamAplic.Cooperativa = 4 Then
'                If Not AlbaranesFacturados(nTabla, cadSelect) Then Exit Sub
'                ' volvemos a comprobar si hay albaranes pendientes de liquidar
'                If Not HayRegParaInforme(nTabla, cadSelect) Then Exit Sub
'            End If
            
            If HayPreciosVariedadesCatadau(TipoPrec, nTabla, cadSelect, Combo1(2).ListIndex) Then
                'D/H fecha
                cDesde = Trim(txtcodigo(6).Text)
                cHasta = Trim(txtcodigo(7).Text)
                cadDesde = CDate(cDesde)
                cadhasta = CDate(cHasta)
                cadAux = "{rprecios.fechaini}= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechaini}=" & DBSet(txtcodigo(6).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                cadAux = "{rprecios.fechafin}= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")"
                If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                cadAux = "{rprecios.fechafin}=" & DBSet(txtcodigo(7).Text, "F")
                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                If Check1(5).Value = 0 Then
                    ' si se trata de anticipos--> seleccionamos los precios de anticipos
                    ' sino los de liquidaciones
                    If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = " & TipoPrec) Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = " & TipoPrec) Then Exit Sub
                Else
                    If Not AnyadirAFormula(cadFormula, "{rprecios.tipofact} = 3") Then Exit Sub
                    If Not AnyadirAFormula(cadSelect, "{rprecios.tipofact} = 3") Then Exit Sub
                End If
                
'                '02/09/2010
'                cadAux = "{rprecios.contador} = (select max(p.contador) from rprecios p where p.codvarie = rhisfruta.codvarie and "
'                cadAux = cadAux & " p.tipofact = " & TipoPrec & " and p.fechaini = " & DBSet(txtcodigo(6).Text, "F")
'                cadAux = cadAux & " and p.fechafin = " & DBSet(txtcodigo(7).Text, "F") & ")"
'                If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                
                Select Case OpcionListado
                    Case 1  '1 - informe de anticipos
                        If CargarTemporalCatadau(Tabla1, cadSelect1, TipoPrec) Then
                            'pasamos como parametro la fecha de anticipo
                            cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
                            numParam = numParam + 1
                            ConSubInforme = True

                            cadFormula = ""
                            'InsertarTemporal (Variedades)
                            If Not AnyadirAFormula(cadFormula, "{tmpliquidacion.codusu} = " & vUsu.Codigo) Then Exit Sub

                            LlamarImprimir
                        End If

'                        If CargarTemporalCatadau(Tabla1, cadSelect1, TipoPrec) Then
'                            Nregs = TotalFacturas(Tabla1, cadSelect1)
'                            If Nregs <> 0 Then
'
'                                Me.Pb1.visible = True
'                                Me.Pb1.Max = Nregs
'                                Me.Pb1.Value = 0
'                                Me.Refresh
'                                b = False
'                                b = InformeAnticiposPicassentNew(Tabla1, cadSelect1, txtcodigo(15).Text, Me.Pb1, Check1(14).Value = 1, Check1(11).Value = 1)
'
'
'                                cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
'                                numParam = numParam + 1
'                                ConSubInforme = True
'
'                                cadFormula = ""
'                                'InsertarTemporal (Variedades)
'                                If Not AnyadirAFormula(cadFormula, "{tmpliquidacion.codusu} = " & vUsu.Codigo) Then Exit Sub
'
'                                LlamarImprimir
'
'                            End If
'                         End If


                    
                    Case 12 '12- informe de liquidaciones
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
' Picassent
                        If CargarTemporalLiquidacionPicassentNew(Tabla1, cadSelect1) Then
                            cadFormula = ""
                            If Not AnyadirAFormula(cadFormula, "{tmpliquidacion.codusu} = " & vUsu.Codigo) Then Exit Sub
                                                                
                            ConSubInforme = True
                            'pasamos como parametro la fecha de anticipo
                            cadParam = cadParam & "pFecAnt=""" & txtcodigo(15).Text & """|"
                            numParam = numParam + 1
                            cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
                            numParam = numParam + 1
                            LlamarImprimir
                        End If
                            
                    
                    Case 2  '2 - listado de prevision de pagos de anticipos
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "
                        nTabla = "(" & nTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        ' Picassent
                        If Combo1(3).ListIndex = 2 Then
                            If CargarTemporalAnticiposCalidadPicassentNew(Tabla1, cadSelect1) Then
                                cadFormula = ""
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                ConSubInforme = True
                                
                                LlamarImprimir
                            End If
                        Else
                            If CargarTemporalAnticiposPicassentNew(Tabla1, cadSelect1) Then
                                cadFormula = ""
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                ConSubInforme = True
                                
                                cadParam = cadParam & "pConBonifica=1|"
                                numParam = numParam + 1
                                LlamarImprimir
                            End If
                        End If
                                                        
                        
                    Case 13 '13- listado de prevision de pagos de liquidaciones
'                        nTabla = "(" & nTabla & ") INNER JOIN rprecios_calidad ON rhisfruta_clasif.codvarie = rprecios_calidad.codvarie and rhisfruta_clasif.codcalid = rprecios_calidad.codcalid "
'                        nTabla = "(" & nTabla & ") INNER JOIN rprecios ON rprecios.codvarie = rprecios_calidad.codvarie and rprecios.tipofact = rprecios_calidad.tipofact and rprecios.contador = rprecios_calidad.contador "

'                        NTabla = "(" & NTabla & ") INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio "
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        ' Picassent
                        '[Monica]22/03/2012: indicamos en el informe si hacemos o no el descuento de comision segun el check1(13)
                        If Check1(13).Value = 1 Then
                            cadParam = cadParam & "pTipo=""Cálculo con comisión""|"
                        Else
                            cadParam = cadParam & "pTipo=""Cálculo sin comisión""|"
                        End If
                        numParam = numParam + 1
                        
                        If Combo1(3).ListIndex = 2 Then
                            ' es igual que la cargatempporal de anticipos pero aqui coge los precios de liquidacion
                            If CargarTemporalLiquidacionesCalidadPicassentNew(Tabla1, cadSelect1) Then
                                cadFormula = ""
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                ConSubInforme = False
                                
                                LlamarImprimir
                            End If
                        Else
                            If CargarTemporalLiquidacionPicassentNew(Tabla1, cadSelect1) Then
                                cadFormula = ""
                                If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                ConSubInforme = False
                                
                                LlamarImprimir
                            End If
                        End If
                    
                    Case 3, 14 '3 .- factura de anticipos
                               '14.- factura de liquidaciones
                        nTabla = "(" & nTabla & ") INNER JOIN rcalidad ON rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
                        
                        If CargarTemporalPicassent(Tabla1, cadSelect1, TipoPrec) Then
                            Nregs = TotalFacturas(Tabla1, cadSelect1)
                            If Nregs <> 0 Then
                                If Not ComprobarTiposMovimiento(TipoPrec, Tabla1, cadSelect1) Then
                                    Exit Sub
                                End If
                                
                                Me.Pb1.visible = True
                                Me.Pb1.Max = Nregs
                                Me.Pb1.Value = 0
                                Me.Refresh
                                b = False
                                If TipoPrec = 0 Then
                                    b = FacturacionAnticiposPicassentNew(Tabla1, cadSelect1, txtcodigo(15).Text, Me.Pb1, Check1(14).Value = 1, Check1(11).Value = 1)
                                Else
                                    b = FacturacionLiquidacionesPicassentNew(Tabla1, cadSelect1, txtcodigo(15).Text, Me.Pb1, TipoPrec, Check1(14).Value = 1, Check1(11).Value = 1)
                                End If
                                If b Then
                                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                                   
                                    'IMPRESION DEL RESUMEN DE LA FACTURACION DE ANTICIPOS/LIQUIDACIONES
                                    If Me.Check1(2).Value Then
                                        cadFormula = ""
                                        cadParam = cadParam & "pFecFac= """ & txtcodigo(15).Text & """|"
                                        numParam = numParam + 1
                                        If TipoPrec = 0 Then
                                            cadParam = cadParam & "pTitulo= ""Resumen Facturación de Anticipos""|"
                                        Else
                                            cadParam = cadParam & "pTitulo= ""Resumen Facturación de Liquidaciones""|"
                                        End If
                                        numParam = numParam + 1
                                        
                                        FecFac = CDate(txtcodigo(15).Text)
                                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                                        ConSubInforme = True
                                        
                                        LlamarImprimir
                                    End If
                                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE ANTICIPOS/LIQUIDACION
                                    If Me.Check1(3).Value Then
                                        cadFormula = ""
                                        cadSelect = ""
                                        If TipoPrec = 0 Then 'Tipo de Factura: Anticipo
                                            cadAux = "({stipom.tipodocu} = 1)"
                                        Else  'Tipo de Factura: Liquidación
                                            cadAux = "({stipom.tipodocu} = 2)"
                                        End If
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
                                        'Nº Factura
                                        If TipoPrec = 0 Then
                                            cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(2) & "])"
                                        Else
                                            cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradas(3) & "])"
                                        End If
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
    
                                        'Fecha de Factura
                                        FecFac = CDate(txtcodigo(15).Text)
                                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                                        cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub
    
                                        indRPT = 23 'Impresion de facturas de socios
                                        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                                        'Nombre fichero .rpt a Imprimir
                                        cadNombreRPT = nomDocu
                                        'Nombre fichero .rpt a Imprimir
                                        If TipoPrec = 0 Then
                                            cadTitulo = "Reimpresión de Facturas Anticipos"
                                        Else
                                            cadTitulo = "Reimpresión de Facturas Liquidaciones"
                                        End If
                                        ConSubInforme = True
    
                                        If indRPT = 23 And vParamAplic.Cooperativa = 2 Then
                                            Dim PrecioApor As Double
                                            PrecioApor = DevuelveValor("select min(precio) from raporreparto")
                                            
                                            cadParam = cadParam & "pPrecioApor=""" & Replace(Format(PrecioApor, "#0.000000"), ",", ".") & """|"
                                            numParam = numParam + 1
                                        End If
    
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
                    End If
                    
                End Select
            '++monica:27/07/2009
            Else
                MsgBox "No hay precios para las calidades en este rango. Revise.", vbExclamation
            End If
        End If
    End If
End Sub

Private Function CargarTemporalPicassent(ctabla As String, cwhere As String, Tipo As Byte) As Boolean
'tipo  0=anticipos
'      1=liquidacion
Dim RS As ADODB.Recordset
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
Dim Cad As String
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

Dim HayPrecio As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalPicassent = False

    Sql2 = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    Sql2 = "delete from tmpliquidacion1 where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql2 = "delete from tmpinformes2 where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    
    '[Monica]24/04/2013: Meto los albaranes y calidades que puede que liquide
    Sql2 = "insert into tmpinformes2 (codusu, importe1, fecha1, importe2, importe3, importe5, importeb1, importe4) select " & vUsu.Codigo & ",rhisfruta.numalbar, rhisfruta.fecalbar,rhisfruta.codvarie, rhisfruta_clasif.codcalid, rhisfruta.codcampo, rhisfruta.codsocio, "
    Sql2 = Sql2 & " sum(coalesce(rhisfruta_clasif.kilosnet,0)) as kilos  "
    Sql2 = Sql2 & " FROM  (" & ctabla & ") inner join tmpvarie on rhisfruta.codvarie = tmpvarie.codvarie "
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        Sql2 = Sql2 & " WHERE " & cwhere
    End If
    Sql2 = Sql2 & " group by 1, 2, 3, 4, 5, 6, 7"
    Sql2 = Sql2 & " having sum(coalesce(rhisfruta_clasif.kilosnet,0)) <> 0 "
    Sql2 = Sql2 & " order by 1, 2, 3, 4, 5, 6, 7 "
    
    conn.Execute Sql2
    
    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, "
    SQL = SQL & " rhisfruta.recolect, rhisfruta.tipoentr, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, "
    SQL = SQL & " sum(coalesce(rhisfruta_clasif.kilosnet,0)) as kilos "
    SQL = SQL & " FROM  (" & ctabla & ") inner join tmpvarie on rhisfruta.codvarie = tmpvarie.codvarie "
    
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8 "
    SQL = SQL & " having sum(coalesce(rhisfruta_clasif.kilosnet,0)) <> 0 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8 "


    Nregs = TotalRegistrosConsulta(SQL)
    
    Label2(10).Caption = "Cargando Tabla Temporal"
    Me.Pb1.visible = True
    Me.Pb1.Max = Nregs
    Me.Pb1.Value = 0
    Me.Refresh

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    '[Monica] 14/12/2009 si es liquidacion complementaria: no se descuentan gastos(complementaria) ponemos gastos = 0
    If Tipo = 1 And Check1(5).Value = 1 Then Tipo = 3 'seleccionamos los precios de liquidacion complementaria
                                    
                                    
    While Not RS.EOF
    
        Label2(12).Caption = "Socio " & RS!Codsocio & " Variedad " & RS!codvarie & "-" & RS!codcalid & " Campo " & RS!codcampo
        IncrementarProgresNew Pb1, 1
        Me.Refresh
        DoEvents
    
        Sql3 = "select fechaini, fechafin, max(contador) as contador from rprecios where codvarie = " & DBSet(RS!codvarie, "N")
        Sql3 = Sql3 & " and tipofact = " & DBSet(Tipo, "N")
        Sql3 = Sql3 & " and fechaini <= " & DBSet(RS!Fecalbar, "F")
        Sql3 = Sql3 & " and fechafin >= " & DBSet(RS!Fecalbar, "F")
        Sql3 = Sql3 & " group by 1,2"
                
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS1.EOF Then
            Contador = DBLet(RS1!Contador, "N")
            FechaIni = DBLet(RS1!FechaIni, "F")
            FechaFin = DBLet(RS1!FechaFin, "F")
        End If
        Set RS1 = Nothing
        If DBLet(RS!Recolect, "N") = 0 Then 'cooperativa
            Sql3 = "select precoop "
        Else
            Sql3 = "select presocio "
        End If
        
        Sql3 = Sql3 & " from rprecios_calidad where codvarie = " & DBSet(RS!codvarie, "N")
        Sql3 = Sql3 & " and codcalid = " & DBSet(RS!codcalid, "N")
        Sql3 = Sql3 & " and tipofact = " & DBSet(Tipo, "N")
        Sql3 = Sql3 & " and contador = " & DBSet(Contador, "N")
        
        Precio = DevuelveValor(Sql3)
        
        '[monica]24/04/2013: miro si hay que liquidar
        HayPrecio = (TotalRegistrosConsulta(Sql3) <> 0)
        If Not HayPrecio Then
        
            Sql4 = "delete from tmpinformes2 where codusu = " & DBSet(vUsu.Codigo, "N") & " and importe2 = " & DBSet(RS!codvarie, "N")
            Sql4 = Sql4 & " and importe3 = " & DBSet(RS!codcalid, "N")
            Sql4 = Sql4 & " and fecha1 between " & DBSet(FechaIni, "F") & " and " & DBSet(FechaFin, "F")

            conn.Execute Sql4
            
        Else
        
            Sql4 = "update tmpinformes2 set precio1 = " & DBSet(Precio, "N")
            Sql4 = Sql4 & ", fecha2 = " & DBSet(FechaIni, "F")
            Sql4 = Sql4 & ", fecha3 = " & DBSet(FechaFin, "F")
            Sql4 = Sql4 & ", campo1 = " & DBSet(Contador, "N")
            Sql4 = Sql4 & ", campo2 = " & DBSet(Tipo, "N")
            Sql4 = Sql4 & " where codusu = " & DBSet(vUsu.Codigo, "N")
            Sql4 = Sql4 & " and importe2 = " & DBSet(RS!codvarie, "N")
            Sql4 = Sql4 & " and importe3 = " & DBSet(RS!codcalid, "N")
            Sql4 = Sql4 & " and fecha1 between " & DBSet(FechaIni, "F") & " and " & DBSet(FechaFin, "F")

            conn.Execute Sql4
        
        End If
        
        RS.MoveNext
    Wend
                                    
                                        
                                    
                                    
    Me.Pb1.visible = False
    Me.Label2(10).Caption = ""
    Me.Label2(12).Caption = ""
    Me.Refresh
                                    
    CargarTemporalPicassent = True
    Exit Function
    
eCargarTemporal:
    Me.Pb1.visible = False
    Me.Label2(10).Caption = ""
    Me.Label2(12).Caption = ""
    Me.Refresh
    
    MuestraError "Cargando temporal", Err.Description
End Function

Private Function CargarTemporalAnticiposPicassentNew(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SqlVar As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bonifica As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

Dim PorcBoni As Currency
Dim PorcComi As Currency

Dim ImporteFVar As Currency
Dim HayPrecio As Boolean


    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposPicassentNew = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    '[Monica]15/04/2013: introducimos las facturas varias
    Sql2 = "delete from tmpsuperficies where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    If CargarTemporalPicassent(ctabla, cwhere, 0) Then
    
        SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
        SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid, rhisfruta.fecalbar, sum(coalesce(rhisfruta_clasif.kilosnet,0)) as kilos "
        SQL = SQL & " from (" & ctabla & ") inner join rcalidad on rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid "
        SQL = SQL & " where " & cwhere
        SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "
        SQL = SQL & " having sum(coalesce(rhisfruta_clasif.kilosnet,0)) <> 0 "
        SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "
        
        
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
                                        'codusu, codsocio, nomsocio, codvarie, nomvarie, Kneto, baseimpo, porceiva, imporiva,
        Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, campo2, importe3, importe4, porcen1, importe5, "
                       'porcerete, imporret, totalfac, bonificacion
        Sql2 = Sql2 & " porcen2, importeb1, importeb2, importeb3) values "
        
        Set vSeccion = New CSeccion
        
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If Not vSeccion.AbrirConta Then
                Exit Function
            End If
        End If
        
        HayReg = False
        If Not RS.EOF Then
            SocioAnt = RS!Codsocio
            VarieAnt = RS!codvarie
            NVarieAnt = RS!nomvarie
            
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                NSocioAnt = vSocio.Nombre
                TipoIRPF = vSocio.TipoIRPF
            End If
        End If
        
        While Not RS.EOF
            '++monica:28/07/2009 añadida la segunda condicion
            If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
                '[Monica]24/02/2014: añadida condicion
                If KilosNet <> 0 Then
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
                    Sql1 = Sql1 & DBSet(TotalFac, "N") & "," & DBSet(Bonifica, "N") & "),"
                End If
                
                VarieAnt = RS!codvarie
                                    
                baseimpo = 0
                Bonifica = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                KilosNet = 0
            End If
            
            If RS!Codsocio <> SocioAnt Then
                '[Monica]15/04/2013: descontamos las facturas varias
                If Check1(14).Value Then                                                                                                 'anticipos            no descontados   q no sean de ventacampo
                    ImporteFVar = DevuelveValor("select sum(totalfac) from fvarcabfact where codsocio = " & DBSet(SocioAnt, "N") & " and enliquidacion = 2 and intliqui = 0 and envtacampo = 0 ")
                                                        'usuario, codsocio, importe facturas varias
                    SqlVar = "insert into tmpsuperficies (codusu, codvarie, superficie1) values (" & vUsu.Codigo & ","
                    SqlVar = SqlVar & DBSet(SocioAnt, "N") & ","
                    SqlVar = SqlVar & DBSet(ImporteFVar, "N") & ")"
                    conn.Execute SqlVar
                End If
            
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(RS!Codsocio) Then
                    If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                    NSocioAnt = vSocio.Nombre
                End If
                SocioAnt = vSocio.Codigo
                TipoIRPF = vSocio.TipoIRPF
            End If
            
            KilosNet = KilosNet + DBLet(RS!Kilos, "N")
            
            Recolect = DBLet(RS!Recolect, "N")
            
            Dim Sql9 As String
            Dim Rs9 As ADODB.Recordset
            Dim Precio As Currency
            
            Sql9 = "select precio1 from tmpinformes2 where fecha1 = " & DBSet(RS!Fecalbar, "F") & " and importe2 = " & DBSet(RS!codvarie, "N")
            Sql9 = Sql9 & " and importe3  = " & DBSet(RS!codcalid, "N") & " and codusu = " & vUsu.Codigo
            
            Set Rs9 = New ADODB.Recordset
            Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            If Not Rs9.EOF Then
                Precio = DBLet(Rs9.Fields(0).Value, "N")
                PorcBoni = 0
                
                ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                If Precio > 0 Then
                    PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                
                    '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                    PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                    If CCur(PorcComi) <> 0 Then
                        Precio = Precio - Round2(Precio * PorcComi / 100, 4)
                    End If
                
                End If
            
                '[Monica]25/01/2016: para el caso de Picassent si la calidad no tiene bonificacion PorcBoni = 0
                If Not EsCalidadConBonificacion(CStr(RS!codvarie), CStr(RS!codcalid)) Then PorcBoni = 0
            
            
                Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * Precio, 2)
                baseimpo = baseimpo + Round2(DBLet(RS!Kilos, "N") * Precio * (1 + (PorcBoni / 100)), 2)
                
            Else
                '[Monica]24/02/2014: añadida condicion
                ' los kilos que le he sumado se los quito
                KilosNet = KilosNet - DBLet(RS!Kilos, "N")
                
            End If
            
            Set Rs9 = Nothing
            
            HayReg = True
            
            RS.MoveNext
        Wend
        
        ' ultimo registro si ha entrado
        If HayReg Then
            '[Monica]15/04/2013: descontamos las facturas varias
            If Check1(14).Value = 1 Then                                                                                            'anticipos             no descontados   que no sean de ventacampo
                ImporteFVar = DevuelveValor("select sum(totalfac) from fvarcabfact where codsocio = " & DBSet(SocioAnt, "N") & " and enliquidacion = 2 and intliqui = 0 and envtacampo = 0")
                                                    'usuario, codsocio, importe facturas varias
                SqlVar = "insert into tmpsuperficies (codusu, codvarie, superficie1) values (" & vUsu.Codigo & ","
                SqlVar = SqlVar & DBSet(SocioAnt, "N") & ","
                SqlVar = SqlVar & DBSet(ImporteFVar, "N") & ")"
                conn.Execute SqlVar
            End If
            
            '[Monica]24/02/2014: añadida condicion
            If KilosNet <> 0 Then
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
                Sql1 = Sql1 & DBSet(TotalFac, "N") & "," & DBSet(Bonifica, "N") & "),"
            End If
            
            
            ' quitamos la ultima coma e insertamos
            If Len(Sql1) <> 0 Then
                Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
                conn.Execute Sql2 & Sql1
            End If
        End If
        
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        Set vSocio = Nothing
        
        CargarTemporalAnticiposPicassentNew = True
        Exit Function
    End If
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalAnticiposCalidadPicassentNew(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CalidAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bonifica As Currency
Dim Importe As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

Dim PorcBoni As Currency
Dim PrecioAnt As Currency
Dim PorcComi As Currency

Dim HayPrecio As Boolean


    On Error GoTo eCargarTemporal
    
    CargarTemporalAnticiposCalidadPicassentNew = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    '[Monica]15/04/2013: introducimos las facturas varias
    Sql2 = "delete from tmpsuperficies where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    If CargarTemporalPicassent(ctabla, cwhere, 0) Then

        SQL = "SELECT rhisfruta.codsocio,  rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
        SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rcalidad.nomcalid, rhisfruta.fecalbar, sum(coalesce(rhisfruta_clasif.kilosnet,0)) as kilos "
        SQL = SQL & " FROM  (" & ctabla & ") inner join rcalidad on rhisfruta_clasif.codcalid = rcalidad.codcalid and rhisfruta_clasif.codvarie = rcalidad.codvarie "
        
        If cwhere <> "" Then
            cwhere = QuitarCaracterACadena(cwhere, "{")
            cwhere = QuitarCaracterACadena(cwhere, "}")
            cwhere = QuitarCaracterACadena(cwhere, "_1")
            SQL = SQL & " WHERE " & cwhere
        End If
        SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "
        SQL = SQL & " having sum(coalesce(rhisfruta_clasif.kilosnet,0)) <> 0 "
        SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "
    
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                                        'codusu,  codvarie, nomvarie, calidad, Kneto,  Precio, importe, bonificacion,
        Sql2 = "insert into tmpinformes (codusu,  importe1, nombre1, campo1, importe2, precio1, importe3, importe4, "
                       'importetotal
        Sql2 = Sql2 & " importe5) values "
        
        Set vSeccion = New CSeccion
        
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If Not vSeccion.AbrirConta Then
                Exit Function
            End If
        End If
        
        HayReg = False
        If Not RS.EOF Then
            VarieAnt = RS!codvarie
            NVarieAnt = RS!nomvarie
            CalidAnt = RS!codcalid
        End If
        
        While Not RS.EOF
            '++monica:28/07/2009 añadida la segunda condicion
            If VarieAnt <> RS!codvarie Or CalidAnt <> RS!codcalid Then
                '[Monica]24/02/2014: añadida condicion
                If HayPrecio Then
                    Sql1 = Sql1 & "(" & vUsu.Codigo & ","
                    Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
                    Sql1 = Sql1 & DBSet(CalidAnt, "N") & ","
                    Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(PrecioAnt, "N") & ","
                    Sql1 = Sql1 & DBSet(Importe, "N") & "," & DBSet(Bonifica - Importe, "N") & ","
                    Sql1 = Sql1 & DBSet(Bonifica, "N") & "),"
                End If
                
                VarieAnt = RS!codvarie
                CalidAnt = RS!codcalid
                
                baseimpo = 0
                Bonifica = 0
                Importe = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                KilosNet = 0
            End If
            
            KilosNet = KilosNet + DBLet(RS!Kilos, "N")
            
            Recolect = DBLet(RS!Recolect, "N")
            
            '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
            Dim Sql9 As String
            Dim Rs9 As ADODB.Recordset
            Dim Precio As Currency
            
            Sql9 = "select precio1 from tmpinformes2 where fecha1 = " & DBSet(RS!Fecalbar, "F") & " and importe2 = " & DBSet(RS!codvarie, "N")
            Sql9 = Sql9 & " and importe3  = " & DBSet(RS!codcalid, "N") & " and codusu = " & vUsu.Codigo
            
            Set Rs9 = New ADODB.Recordset
            Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            If Not Rs9.EOF Then
                '[Monica]24/02/2014: añadida variable
                HayPrecio = True
            
                Precio = DBLet(Rs9.Fields(0).Value, "N")
                PorcBoni = 0
                PorcComi = 0
                
                ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                If Precio > 0 Then
                    PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                    
                    '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                    PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                    If CCur(PorcComi) <> 0 Then
                        Precio = Precio - Round2(Precio * PorcComi / 100, 4)
                    End If
                End If
                PrecioAnt = Precio
                Importe = Importe + Round2(DBLet(RS!Kilos, "N") * Precio, 2)
                
                '[Monica]25/01/2016: para el caso de Picassent si la calidad no tiene bonificacion PorcBoni = 0
                If Not EsCalidadConBonificacion(CStr(RS!codvarie), CStr(RS!codcalid)) Then PorcBoni = 0
                
                
                Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * Precio * (1 + (PorcBoni / 100)), 2)
            Else
                '[Monica]24/02/2014: añadida variable
                HayPrecio = False
            End If
            Set Rs9 = Nothing
            'hasta aqui
            HayReg = True
            
            RS.MoveNext
        Wend
        
        ' ultimo registro si ha entrado
        If HayReg Then
            '[Monica]24/02/2014: añadida condicion
            If HayPrecio Then
                Sql1 = Sql1 & "(" & vUsu.Codigo & ","
                Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
                Sql1 = Sql1 & DBSet(CalidAnt, "N") & ","
                Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(PrecioAnt, "N") & ","
                Sql1 = Sql1 & DBSet(Importe, "N") & "," & DBSet(Bonifica - Importe, "N") & ","
                Sql1 = Sql1 & DBSet(Bonifica, "N") & "),"
            End If
            ' quitamos la ultima coma e insertamos
            If Len(Sql1) <> 0 Then
                Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
                conn.Execute Sql2 & Sql1
            End If
        End If
        
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        Set vSocio = Nothing
        
        CargarTemporalAnticiposCalidadPicassentNew = True
        Exit Function
        
    End If
        
        
eCargarTemporal:
    MuestraError Err.Number, "Cargando temporal", Err.Description
End Function



Private Function CargarTemporalLiquidacionPicassentNew(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim SqlLiq As String

Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CampoAnt As Long
Dim AlbarAnt As Long
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
Dim Incremento As Currency
Dim TotalFac As Currency
Dim Bonifica As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim PorcBoni As Currency
Dim vPorcIva As String
Dim vPorcGasto As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean
Dim vGastos As Currency


Dim BaseImpoFactura As Currency
Dim ImpoIvaFactura As Currency
Dim ImpoAporFactura As Currency
Dim ImpoRetenFactura As Currency
Dim ImpoGastosFactura As Currency
Dim ImpoTotalFactura As Currency
Dim ImpoFrasVarias As Currency

Dim SqlFactura As String
Dim sqlLiquid As String
Dim ImpBonif As Currency
Dim ImpTot As Currency

Dim PorcComi As Currency

Dim Sql9 As String
Dim Rs9 As ADODB.Recordset
Dim Precio As Currency
Dim vConta As String
Dim vFecIni As String
Dim vFecFin As String
Dim vTipo As String

Dim HayPrecio As Boolean


    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionPicassentNew = False

    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql2 = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql2 = "delete from tmpfactura where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    If CargarTemporalPicassent(ctabla, cwhere, 1) Then
        
        SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo, rhisfruta.numalbar, rhisfruta.fecalbar, "
        SQL = SQL & "rhisfruta.recolect,  rhisfruta_clasif.codcalid, rcalidad.nomcalid, sum(coalesce(rhisfruta_clasif.kilosnet,0)) as kilos "
    ''[Monica]01/09/2010 : sustituida la siguiente linea por
    ''    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilos "
        SQL = SQL & " FROM  (" & ctabla & " ) inner join rcalidad on rhisfruta_clasif.codcalid = rcalidad.codcalid and rhisfruta_clasif.codvarie = rcalidad.codvarie "
    
        
        If cwhere <> "" Then
            cwhere = QuitarCaracterACadena(cwhere, "{")
            cwhere = QuitarCaracterACadena(cwhere, "}")
            cwhere = QuitarCaracterACadena(cwhere, "_1")
            SQL = SQL & " WHERE " & cwhere
        End If
        SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8, 9 "
        SQL = SQL & " having sum(coalesce(rhisfruta_clasif.kilosnet,0)) <> 0"
        SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8, 9 "
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                                        'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  bonificacion, gastos, incremento, anticipos, porceiva, imporiva,
        Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importeb6, importe4, porcen1, importe5, "
                       'porcerete, imporret, totalfac, max(contador),tipofact, rprecios.fecini, rprecios.fecfin
        Sql2 = Sql2 & " porcen2, importeb1, importeb2, campo1, campo2, fecha1, fecha2) values "
        
        'cargamos las bonificaciones para el informe de liquidacion
                                                                                    'albaran            %bonif  impbonif, total
        SqlLiq = "insert into tmpliquidacion (codusu, codsocio, codvarie, codcampo, kilosnet, codcalid, precio, importe, gastos) values "
        
        
        Set vSeccion = New CSeccion
        
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If Not vSeccion.AbrirConta Then
                Exit Function
            End If
        End If
        
        HayReg = False
        If Not RS.EOF Then
            SocioAnt = RS!Codsocio
            VarieAnt = RS!codvarie
            NVarieAnt = RS!nomvarie
            CampoAnt = RS!codcampo
            AlbarAnt = RS!numalbar
            
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
                TipoIRPF = vSocio.TipoIRPF
            End If
        End If
        Bonifica = 0
        baseimpo = 0
        KilosNet = 0
        ImpoGastos = 0
        
        BaseImpoFactura = 0
        ImpoIvaFactura = 0
        ImpoAporFactura = 0
        ImpoRetenFactura = 0
        ImpoTotalFactura = 0
        ImpoGastosFactura = 0
        
        
        sqlLiquid = ""
        
        While Not RS.EOF
            If AlbarAnt <> RS!numalbar Or VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
            
                '[Monica]24/02/2014: añadida condicion
                If KilosNet <> 0 Then
                
                    ' gastos de los albaranes
                    Sql4 = "select sum(rhisfruta_gastos.importe) "
                    Sql4 = Sql4 & " from rhisfruta_gastos "
                    Sql4 = Sql4 & " where rhisfruta_gastos.numalbar = " & DBSet(AlbarAnt, "N")
                    
                    ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
                    
                    '[Monica]23/07/2012: si es complementaria no hay gastos
                    If Check1(5).Value = 1 Then ' si es complementaria no hay gastos
                        ImpoGastos = 0
                    End If
                    
                    ImpoGastosFactura = ImpoGastosFactura + DevuelveValor(Sql4)
                
                End If
                
                AlbarAnt = RS!numalbar
            End If
        
            ' 23/07/2009: añadido el or con la segunda condicion
            If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
                '[Monica]24/02/2014: añadida condicion
                If KilosNet <> 0 Then
                    '[Monica]10/01/2014: cargamos los aumentos por variedad que tenga
                    Sql4 = "select sum(ringresos.importe) from ringresos where codsocio = " & DBSet(SocioAnt, "N")
                    Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
                    
                    Incremento = DevuelveValor(Sql4)
                
                    ' anticipos
                    Sql4 = "select sum(rfactsoc_variedad.imporvar) "
                    Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                    Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                    Sql4 = Sql4 & " where rfactsoc_variedad.codtipom in (" & DBSet(vSocio.CodTipomAnt, "T") & ",'FAT') " ' "FAA"
                    Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
                    Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
                    Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
                    
                    Anticipos = DevuelveValor(Sql4)
                    
                    Bruto = baseimpo - Bonifica
                    
                    ImpoBonif = Bonifica
                    'ImpoBonif = BaseImpo - Bonifica
                    
                    '[Monica]10/01/2014: añadimos el incremento
                    baseimpo = baseimpo - Anticipos + Incremento
                    
                    BaseImpoFactura = BaseImpoFactura + baseimpo
                    
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
                
                    If Check1(5).Value = 1 Then ' si es complementaria no hay importe de aportacion
                        ImpoAport = 0
                    Else
                        ImpoAport = DevuelveValor("select importe from raporreparto where codsocio = " & DBSet(SocioAnt, "N") & " and tipoentr = 0")
                    End If
                
                    TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
                    TotalFac = TotalFac - ImpoGastos
                    
                    Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
                    Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
                    Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
                    Sql1 = Sql1 & DBSet(Bruto, "N") & ","
                    Sql1 = Sql1 & DBSet(ImpoBonif, "N") & ","
                    Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
                    '[Monica]10/01/2014: añadimos el incremento
                    Sql1 = Sql1 & DBSet(Incremento, "N") & ","
                    Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
                    Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
                    Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
                    Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
                    Sql1 = Sql1 & DBSet(TotalFac, "N")
                    Sql1 = Sql1 & ","
                    Sql1 = Sql1 & DBSet(vConta, "N") & "," & DBSet(vTipo, "N") & "," & DBSet(vFecIni, "F") & "," & DBSet(vFecFin, "F") & "),"
                    
                End If
                
                VarieAnt = RS!codvarie
                
                baseimpo = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                KilosNet = 0
                
                ImpoGastos = 0
                ImpoBonif = 0
                Anticipos = 0
                '[Monica]10/01/2014: añadimos el incremento
                Incremento = 0
                Bonifica = 0
                
            End If
            
            If RS!Codsocio <> SocioAnt Then
            
                '[Monica]24/02/2014: añadida condicion
                If BaseImpoFactura <> 0 Then
            
                    ImpoIvaFactura = Round2((BaseImpoFactura) * ComprobarCero(vPorcIva) / 100, 2)
                
                    Select Case TipoIRPF
                        Case 0
                            ImpoRetenFactura = Round2((BaseImpoFactura + ImpoIvaFactura) * vParamAplic.PorcreteFacSoc / 100, 2)
                            PorcReten = vParamAplic.PorcreteFacSoc
                        Case 1
                            ImpoRetenFactura = Round2(BaseImpoFactura * vParamAplic.PorcreteFacSoc / 100, 2)
                            PorcReten = vParamAplic.PorcreteFacSoc
                        Case 2
                            ImpoRetenFactura = 0
                            PorcReten = 0
                    End Select
                
                    If Check1(5).Value = 1 Then ' si es complementaria no hay importe de aportacion
                        ImpoAporFactura = 0
                    Else
                        ImpoAporFactura = DevuelveValor("select importe from raporreparto where codsocio = " & DBSet(SocioAnt, "N") & " and tipoentr = 0")
                    End If
                    
                    '[Monica]15/04/2013: si hay importe de facturas varias a descontar del socio
                    ImpoFrasVarias = 0
                    If Check1(14).Value = 1 Then                                                                                      'en liquidacion       no descontada         que no sea vtacampo
                        ImpoFrasVarias = DevuelveValor("select sum(totalfac) from fvarcabfact where codsocio = " & DBSet(SocioAnt, "N") & " and enliquidacion = 1 and intliqui = 0 and envtacampo = 0")
                    End If
                    
                    ImpoTotalFactura = BaseImpoFactura + ImpoIvaFactura - ImpoRetenFactura - ImpoAporFactura - ImpoGastosFactura '- ImpoFrasVarias
                    
                    SqlFactura = "insert into tmpfactura(codusu,codsocio,baseimpo,imporiva,impreten,impapor,impgastos,totalfac, impfrasvar) values ( "
                    SqlFactura = SqlFactura & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(BaseImpoFactura, "N") & ","
                    SqlFactura = SqlFactura & DBSet(ImpoIvaFactura, "N") & "," & DBSet(ImpoRetenFactura, "N") & ","
                    SqlFactura = SqlFactura & DBSet(ImpoAporFactura, "N") & "," & DBSet(ImpoGastosFactura, "N") & ","
                    SqlFactura = SqlFactura & DBSet(ImpoTotalFactura, "N") & "," & DBSet(ImpoFrasVarias, "N") & ")"
                    
                    conn.Execute SqlFactura
                    
                End If
                
                BaseImpoFactura = 0
                ImpoIvaFactura = 0
                ImpoRetenFactura = 0
                ImpoAporFactura = 0
                ImpoGastosFactura = 0
                ImpoTotalFactura = 0
                ImpoFrasVarias = 0
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(RS!Codsocio) Then
                    If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        vPorcGasto = ""
                        vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                    End If
                    NSocioAnt = vSocio.Nombre
                End If
                SocioAnt = vSocio.Codigo
                TipoIRPF = vSocio.TipoIRPF
            End If
            
            KilosNet = KilosNet + DBLet(RS!Kilos, "N")
            
            '[Monica]01/09/2010: añadido precios
            
            Sql9 = "select precio1, fecha2, fecha3, campo1, campo2 from tmpinformes2 where fecha1 = " & DBSet(RS!Fecalbar, "F") & " and importe2 = " & DBSet(RS!codvarie, "N")
            Sql9 = Sql9 & " and importe3 = " & DBSet(RS!codcalid, "N") & " and codusu = " & vUsu.Codigo
            Sql9 = Sql9 & " and importe1 = " & DBSet(RS!numalbar, "N")
            
            Set Rs9 = New ADODB.Recordset
            Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            If Not Rs9.EOF Then
                '[Monica]24/02/2014: añadido
                HayPrecio = True
                
                Precio = DBLet(Rs9.Fields(0).Value, "N")
                PorcBoni = 0
                PorcComi = 0
                vConta = DBLet(Rs9!campo1, "N")
                ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                If Precio > 0 Then
                    PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                    
                    '[Monica]22/03/2012: Solo si le indicamos que no calcule comision no lo hace (solo prevision de liquidacion)
                    If Check1(13).Value Then
                        '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                        PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                        If CCur(PorcComi) <> 0 Then
                            Precio = Precio - Round2(Precio * PorcComi / 100, 4)
                        End If
                    End If
                End If
            
                '[Monica]25/01/2016: para el caso de Picassent si la calidad no tiene bonificacion PorcBoni = 0
                If Not EsCalidadConBonificacion(CStr(RS!codvarie), CStr(RS!codcalid)) Then PorcBoni = 0
            
            
                ImpBonif = Round2(DBLet(RS!Kilos, "N") * Precio * (PorcBoni / 100), 2)
                ImpTot = Round2(DBLet(RS!Kilos, "N") * Precio, 2) + ImpBonif
            
                Bonifica = Bonifica + ImpBonif
                baseimpo = baseimpo + ImpTot
                    
                vFecIni = DBLet(Rs9!fecha2, "F") ' fechaini
                vFecFin = DBLet(Rs9!fecha3, "F") ' fechafin
                vTipo = DBLet(Rs9!campo2, "N")  ' tipo de factura
            
            Else
                '[Monica]24/02/2014: añadida else
                HayPrecio = False
                KilosNet = KilosNet - DBLet(RS!Kilos, "N")
            End If
            
            Set Rs9 = Nothing
            
            '[Monica]24/02/2014: añadida condicion
            If HayPrecio Then
                ' insertamos en tmpliquidacion la linea de calidad
                sqlLiquid = sqlLiquid & "(" & vUsu.Codigo & ", " & DBSet(RS!Codsocio, "N") & "," & DBSet(RS!codvarie, "N") & ","
                sqlLiquid = sqlLiquid & DBSet(RS!codcampo, "N") & "," & DBSet(RS!numalbar, "N") & "," & DBSet(RS!codcalid, "N") & "," & DBSet(PorcBoni, "N") & ","
                sqlLiquid = sqlLiquid & DBSet(ImpBonif, "N") & "," & DBSet(ImpTot, "N") & "),"
            End If
            
            
            'hasta aqui
                
            HayReg = True
            
            RS.MoveNext
        Wend
        
        ' Metemos las bonificaciones
        If sqlLiquid <> "" Then
            conn.Execute SqlLiq & Mid(sqlLiquid, 1, Len(sqlLiquid) - 1)
        End If
        
        ' ultimo registro si ha entrado
        If HayReg Then
        
            '[Monica]24/02/2014: añadida condicion
            If KilosNet <> 0 Then
            
                ' gastos de los albaranes
                Sql4 = "select sum(rhisfruta_gastos.importe) "
                Sql4 = Sql4 & " from rhisfruta_gastos "
                Sql4 = Sql4 & " where rhisfruta_gastos.numalbar = " & DBSet(AlbarAnt, "N")
                
                ImpoGastos = ImpoGastos + DevuelveValor(Sql4)
                
                '[Monica]23/07/2012: si es complementaria no hay gastos
                If Check1(5).Value = 1 Then ' si es complementaria no hay gastos
                    ImpoGastos = 0
                    ImpoGastosFactura = 0
                Else
                    ImpoGastosFactura = ImpoGastosFactura + DevuelveValor(Sql4)
                End If
                
                '[Monica]10/01/2014: cargamos los aumentos por variedad que tenga
                Sql4 = "select sum(ringresos.importe) from ringresos where codsocio = " & DBSet(SocioAnt, "N")
                Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
                
                Incremento = DevuelveValor(Sql4)
                
                ' anticipos
                Sql4 = "select sum(rfactsoc_variedad.imporvar) "
                Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql4 = Sql4 & " where rfactsoc_variedad.codtipom in (" & DBSet(vSocio.CodTipomAnt, "T") & ",'FAT')" ' "FAA"
                Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
                Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
                Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"
                
                Anticipos = DevuelveValor(Sql4)
                
                Bruto = baseimpo - Bonifica
                
                ImpoBonif = Bonifica
                
                '[Monica]10/01/2014: cargamos los aumentos por variedad que tenga
                baseimpo = baseimpo - Anticipos + Incremento
                
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
            
                If Check1(5).Value = 1 Then
                    ImpoAport = 0
                Else
                    ImpoAport = DevuelveValor("select importe from raporreparto where codsocio = " & DBSet(SocioAnt, "N") & " and tipoentr = 0")
                End If
            
                TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
                TotalFac = TotalFac - ImpoGastos
                
                Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
                Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
                Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
                Sql1 = Sql1 & DBSet(Bruto, "N") & ","
                Sql1 = Sql1 & DBSet(ImpoBonif, "N") & ","
                Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
                Sql1 = Sql1 & DBSet(Incremento, "N") & ","
                Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
        '            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
                Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
                Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
                Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
                Sql1 = Sql1 & DBSet(TotalFac, "N")
        '02/09/2010
        '            Sql1 = Sql1 & "),"
                Sql1 = Sql1 & ","
                Sql1 = Sql1 & DBSet(vConta, "N") & "," & DBSet(vTipo, "N") & "," & DBSet(vFecIni, "F") & "," & DBSet(vFecFin, "F") & "),"
                
            End If
            
            ' quitamos la ultima coma e insertamos
            Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
            conn.Execute Sql2 & Sql1
            
            '[Monica]24/02/2014: añadida condicion
            If baseimpo <> 0 Then
                BaseImpoFactura = BaseImpoFactura + baseimpo
                ImpoIvaFactura = Round2((BaseImpoFactura) * ComprobarCero(vPorcIva) / 100, 2)
            
                Select Case TipoIRPF
                    Case 0
                        ImpoRetenFactura = Round2((BaseImpoFactura + ImpoIvaFactura) * vParamAplic.PorcreteFacSoc / 100, 2)
                        PorcReten = vParamAplic.PorcreteFacSoc
                    Case 1
                        ImpoRetenFactura = Round2(BaseImpoFactura * vParamAplic.PorcreteFacSoc / 100, 2)
                        PorcReten = vParamAplic.PorcreteFacSoc
                    Case 2
                        ImpoRetenFactura = 0
                        PorcReten = 0
                End Select
            
                If Check1(5).Value = 1 Then
                    ImpoAporFactura = 0
                Else
                    ImpoAporFactura = DevuelveValor("select importe from raporreparto where codsocio = " & DBSet(SocioAnt, "N") & " and tipoentr = 0")
                End If
                
                '[Monica]15/04/2013: si hay importe de facturas varias a descontar del socio
                ImpoFrasVarias = 0
                If Check1(14).Value = 1 Then                                                                                          'liquidacion          no descontada    que no sea vtacampo
                   ImpoFrasVarias = DevuelveValor("select sum(totalfac) from fvarcabfact where codsocio = " & DBSet(SocioAnt, "N") & " and enliquidacion = 1 and intliqui = 0 and envtacampo = 0")
                End If
                
                ImpoTotalFactura = BaseImpoFactura + ImpoIvaFactura - ImpoRetenFactura - ImpoAporFactura - ImpoGastosFactura ' - ImpoFrasVarias
                
                SqlFactura = "insert into tmpfactura(codusu,codsocio,baseimpo,imporiva,impreten,impapor,impgastos,totalfac,impfrasvar) values ( "
                SqlFactura = SqlFactura & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(BaseImpoFactura, "N") & ","
                SqlFactura = SqlFactura & DBSet(ImpoIvaFactura, "N") & "," & DBSet(ImpoRetenFactura, "N") & ","
                SqlFactura = SqlFactura & DBSet(ImpoAporFactura, "N") & "," & DBSet(ImpoGastosFactura, "N") & ","
                SqlFactura = SqlFactura & DBSet(ImpoTotalFactura, "N") & "," & DBSet(ImpoFrasVarias, "N") & ")"
                
                conn.Execute SqlFactura
            End If
                
        End If
        
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        Set vSocio = Nothing
        
        CargarTemporalLiquidacionPicassentNew = True
        Exit Function
        
    End If
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalLiquidacionesCalidadPicassentNew(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim SocioAnt As Long
Dim VarieAnt As Long
Dim CalidAnt As Long
Dim NSocioAnt As String
Dim NVarieAnt As String
Dim Recolect As Integer
            
Dim Neto As Currency
Dim baseimpo As Currency
Dim Bonifica As Currency
Dim Importe As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim TotalFac As Currency
Dim KilosNet As Currency
Dim PorcReten As Currency
Dim vPorcIva As String
Dim PorcIva As Currency
Dim TipoIRPF As Currency

    
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Cad As String
Dim HayReg As Boolean

Dim PorcBoni As Currency
Dim PrecioAnt As Currency
Dim PorcComi As Currency

Dim HayPrecio As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionesCalidadPicassentNew = False

    
    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql2 = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql2 = "delete from tmpfactura where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    ctabla = QuitarCaracterACadena(ctabla, "{")
    ctabla = QuitarCaracterACadena(ctabla, "}")
    
    If CargarTemporalPicassent(ctabla, cwhere, 1) Then
    
        SQL = "SELECT rhisfruta.codsocio,  rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codcampo,"
        SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, "
        SQL = SQL & "rhisfruta.numalbar, sum(coalesce(rhisfruta_clasif.kilosnet,0)) as kilos "
        
        SQL = SQL & " FROM  " & ctabla
        
        If cwhere <> "" Then
            cwhere = QuitarCaracterACadena(cwhere, "{")
            cwhere = QuitarCaracterACadena(cwhere, "}")
            cwhere = QuitarCaracterACadena(cwhere, "_1")
            SQL = SQL & " WHERE " & cwhere
        End If
        SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect, rhisfruta.numalbar "
        SQL = SQL & " having sum(coalesce(rhisfruta_clasif.kilosnet,0)) <> 0 "
        SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect, rhisfruta.numalbar "
    
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                                        'codusu,  codvarie, nomvarie, calidad, Kneto,  Precio, importe, bonificacion,
        Sql2 = "insert into tmpinformes (codusu,  importe1, nombre1, campo1, importe2, precio1, importe3, importe4, "
                       'importetotal
        Sql2 = Sql2 & " importe5) values "
        
        Set vSeccion = New CSeccion
        
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If Not vSeccion.AbrirConta Then
                Exit Function
            End If
        End If
        
        HayReg = False
        If Not RS.EOF Then
            VarieAnt = RS!codvarie
            NVarieAnt = RS!nomvarie
            CalidAnt = RS!codcalid
        End If
        
        While Not RS.EOF
            '++monica:28/07/2009 añadida la segunda condicion
            If VarieAnt <> RS!codvarie Or CalidAnt <> RS!codcalid Then
                '[Monica]24/02/2014: añadida condicion
                If HayPrecio Then
                    Sql1 = Sql1 & "(" & vUsu.Codigo & ","
                    Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
                    Sql1 = Sql1 & DBSet(CalidAnt, "N") & ","
                    Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(PrecioAnt, "N") & ","
                    Sql1 = Sql1 & DBSet(Importe, "N") & "," & DBSet(Bonifica - Importe, "N") & ","
                    Sql1 = Sql1 & DBSet(Bonifica, "N") & "),"
                End If
                
                VarieAnt = RS!codvarie
                CalidAnt = RS!codcalid
                
                baseimpo = 0
                Bonifica = 0
                Importe = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                KilosNet = 0
            End If
            
            KilosNet = KilosNet + DBLet(RS!Kilos, "N")
            
            Recolect = DBLet(RS!Recolect, "N")
            
            
            Dim Sql9 As String
            Dim Rs9 As ADODB.Recordset
            Dim Precio As Currency
            Dim vConta As String
            Dim vFecIni As String
            Dim vFecFin As String
            Dim vTipo As String
                
            Sql9 = "select precio1, fecha2, fecha3, campo1, campo2 from tmpinformes2 where fecha1 = " & DBSet(RS!Fecalbar, "F") & " and importe2 = " & DBSet(RS!codvarie, "N")
            Sql9 = Sql9 & " and importe3 = " & DBSet(RS!codcalid, "N") & " and codusu = " & vUsu.Codigo
            Sql9 = Sql9 & " and importe1 = " & DBSet(RS!numalbar, "N")
            
            Set Rs9 = New ADODB.Recordset
            Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            If Not Rs9.EOF Then
                '[Monica]24/02/2014: añadida variable
                HayPrecio = True
            
                Precio = DBLet(Rs9.Fields(0).Value, "N")
                PorcBoni = 0
                PorcComi = 0
                vConta = DBLet(Rs9!campo1, "N")
                ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                If Precio > 0 Then
                    PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                    
                    '[Monica]22/03/2012: Solo si le indicamos que no calcule comision no lo hace (solo prevision de liquidacion)
                    If Check1(13).Value Then
                        '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                        PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                        If CCur(PorcComi) <> 0 Then
                            Precio = Precio - Round2(Precio * PorcComi / 100, 4)
                        End If
                    End If
                End If
            
                '[Monica]25/01/2016: para el caso de Picassent si la calidad no tiene bonificacion PorcBoni = 0
                If Not EsCalidadConBonificacion(CStr(RS!codvarie), CStr(RS!codcalid)) Then PorcBoni = 0
            
                PrecioAnt = Precio
                Importe = Importe + Round2(DBLet(RS!Kilos, "N") * Precio, 2)
                Bonifica = Bonifica + Round2(DBLet(RS!Kilos, "N") * Precio, 2) + Round2(DBLet(RS!Kilos, "N") * Precio * PorcBoni / 100, 2)
            
                vFecIni = DBLet(Rs9!fecha2, "F") ' fechaini
                vFecFin = DBLet(Rs9!fecha3, "F") ' fechafin
                vTipo = DBLet(Rs9!campo2, "N")  ' tipo de factura
            Else
                '[Monica]24/02/2014: añadida condicion
                HayPrecio = False
                
            End If
            Set Rs9 = Nothing
            
            HayReg = True
            
            RS.MoveNext
        Wend
        
        ' ultimo registro si ha entrado
        If HayReg Then
            '[Monica]24/02/2014: añadida condicion
            If HayPrecio Then
                Sql1 = Sql1 & "(" & vUsu.Codigo & ","
                Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
                Sql1 = Sql1 & DBSet(CalidAnt, "N") & ","
                Sql1 = Sql1 & DBSet(KilosNet, "N") & "," & DBSet(PrecioAnt, "N") & ","
                Sql1 = Sql1 & DBSet(Importe, "N") & "," & DBSet(Bonifica - Importe, "N") & ","
                Sql1 = Sql1 & DBSet(Bonifica, "N") & "),"
            End If
            
            If Len(Sql1) > 0 Then
                ' quitamos la ultima coma e insertamos
                Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
                conn.Execute Sql2 & Sql1
            End If
        End If
        
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        Set vSocio = Nothing
        
    End If
    
    CargarTemporalLiquidacionesCalidadPicassentNew = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargando temporal", Err.Description
End Function


Private Function CargarTemporalLiquidacionAlziraNew(ctabla As String, cwhere As String) As Boolean
Dim RS As ADODB.Recordset
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
Dim Cad As String
Dim HayReg As Boolean
Dim vAnticipos As Currency

    On Error GoTo eCargarTemporal
    
    CargarTemporalLiquidacionAlziraNew = False

    If CargarTemporalCatadau(ctabla, cwhere, 1) Then
        '[Monica]24/04/2013: pq en la anterior funcion se graba la tmpinformes
        Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
        conn.Execute Sql2

    
        SQL = "SELECT tmpliquidacion.codsocio, tmpliquidacion.codvarie, variedades.nomvarie, tmpliquidacion.codcampo,"
        SQL = SQL & " sum(tmpliquidacion.kilosnet) as kilos , sum(tmpliquidacion.importe) as importe "
        SQL = SQL & " FROM tmpliquidacion, variedades where codusu = " & vUsu.Codigo
        SQL = SQL & " and tmpliquidacion.codvarie = variedades.codvarie "
        SQL = SQL & " group by 1, 2, 3, 4 "
        SQL = SQL & " order by 1, 2, 3, 4 "
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                                        'codusu, codsocio, nomsocio, codvarie, nomvarie, neto,  impbruto,  gastos,    anticipos, baseimpo, porceiva, imporiva,
        Sql2 = "insert into tmpinformes (codusu, importe1, nombre1, importe2, nombre2, importe3, importeb3, importeb4, importeb5, importe4, porcen1, importe5, "
                       'porcerete, imporret, totalfac
        Sql2 = Sql2 & " porcen2, importeb1, importeb2) values "
        
        Set vSeccion = New CSeccion
        
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If Not vSeccion.AbrirConta Then
                Exit Function
            End If
        End If
        
        HayReg = False
        If Not RS.EOF Then
            SocioAnt = RS!Codsocio
            VarieAnt = RS!codvarie
            NVarieAnt = RS!nomvarie
            CampoAnt = RS!codcampo
            
            Set vSocio = Nothing
            Set vSocio = New cSocio
            If vSocio.LeerDatos(RS!Codsocio) Then
                If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                    vPorcIva = ""
                    '[Monica]05/03/2014:
                    If vParamAplic.Cooperativa = 4 Then
                        '[Monica]29/04/2011: INTERNAS
                        If vSocio.EsFactADVInt Then
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                        Else
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                    
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                End If
                NSocioAnt = vSocio.Nombre
                TipoIRPF = vSocio.TipoIRPF
            End If
        End If
        
        While Not RS.EOF
           
            ' 23/07/2009: añadido el or con la segunda condicion
            If VarieAnt <> RS!codvarie Or SocioAnt <> RS!Codsocio Then
                
                Bruto = baseimpo
                
                ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
                
                baseimpo = baseimpo - ImpoGastos - Anticipos
                
                ImpoIva = Round2((baseimpo) * ImporteSinFormato(vPorcIva) / 100, 2)
                
                
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
            
                ImpoAport = Round2((Bruto - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
            
                TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
                
                Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
                Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
                Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
                Sql1 = Sql1 & DBSet(Bruto, "N") & ","
                Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
                Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
                Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
                Sql1 = Sql1 & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
                Sql1 = Sql1 & DBSet(PorcReten, "N", "S") & ","
                Sql1 = Sql1 & DBSet(ImpoReten, "N", "S") & ","
                Sql1 = Sql1 & DBSet(TotalFac, "N") & "),"
                
                VarieAnt = RS!codvarie
                baseimpo = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                KilosNet = 0
                
                ImpoGastos = 0
                Anticipos = 0
                
            End If
            
            If RS!Codsocio <> SocioAnt Then
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(RS!Codsocio) Then
                    If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        If vParamAplic.Cooperativa = 4 Then
                            '[Monica]29/04/2011: INTERNAS
                            If vSocio.EsFactADVInt Then
                                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                            Else
                                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                            End If
                        Else
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                        
                        vPorcGasto = ""
                        vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                    End If
                    NSocioAnt = vSocio.Nombre
                End If
                SocioAnt = vSocio.Codigo
                TipoIRPF = vSocio.TipoIRPF
            End If
            
            KilosNet = KilosNet + DBLet(RS!Kilos, "N")
            
            
            ' anticipos
            Sql4 = "select sum(rfactsoc_variedad.imporvar) "
            Sql4 = Sql4 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql4 = Sql4 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql4 = Sql4 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' "FAA"
            Sql4 = Sql4 & " and rfactsoc.codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and rfactsoc.esanticipogasto = 0 "
            Sql4 = Sql4 & " and rfactsoc_variedad.codvarie = " & DBSet(VarieAnt, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.codcampo = " & DBSet(RS!codcampo, "N")
            Sql4 = Sql4 & " and rfactsoc_variedad.descontado = 0"


            vAnticipos = DevuelveValor(Sql4)

            baseimpo = baseimpo + DBLet(RS!Importe, "N")
                
                
            '[Monica]10/03/2014: esto solo seria para el caso de alzira
            '                    si no permitimos facturas negativas el valor de anticipos es mayor que la base imponible
            If Check1(21).Value = 1 And baseimpo < vAnticipos Then
                ' si no queremos que sea negativa no descuento los anticipos
                vAnticipos = 0
            Else
'                baseimpo = baseimpo - vAnticipos
            End If

            Anticipos = Anticipos + vAnticipos


            ' gastos
            Sql4 = "select sum(gastos) from tmpliquidacion1 "
            Sql4 = Sql4 & " where codsocio = " & DBSet(SocioAnt, "N")
            Sql4 = Sql4 & " and codusu = " & vUsu.Codigo
            Sql4 = Sql4 & " and codvarie = " & DBSet(VarieAnt, "N")
            
            ImpoGastos = DevuelveValor(Sql4)
                
            HayReg = True
            
            RS.MoveNext
        Wend
            
        ' ultimo registro si ha entrado
        If HayReg Then
            
            Bruto = baseimpo
            
            ImpoGastos = ImpoGastos + Round2(Bruto * ImporteSinFormato(vPorcGasto) / 100, 2)
            
            baseimpo = baseimpo - ImpoGastos - Anticipos
            
            ImpoIva = Round2((baseimpo) * ImporteSinFormato(vPorcIva) / 100, 2)
            
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
        
            ImpoAport = Round2((Bruto - ImpoGastos) * vParamAplic.PorcenAFO / 100, 2)
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAport
            
            Sql1 = Sql1 & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NSocioAnt, "T") & ","
            Sql1 = Sql1 & DBSet(VarieAnt, "N") & "," & DBSet(NVarieAnt, "T") & ","
            Sql1 = Sql1 & DBSet(KilosNet, "N") & ","
            Sql1 = Sql1 & DBSet(Bruto, "N") & ","
            Sql1 = Sql1 & DBSet(ImpoGastos, "N") & ","
            Sql1 = Sql1 & DBSet(Anticipos, "N") & ","
            Sql1 = Sql1 & DBSet(baseimpo, "N") & ","
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
        
        CargarTemporalLiquidacionAlziraNew = True
        Exit Function
    End If
        
    
eCargarTemporal:
    MuestraError "Cargando temporal", Err.Description
End Function






Public Function InformeAnticiposPicassentNew(ctabla As String, cwhere As String, FecFac As String, Pb1 As ProgressBar, DescontarFVarias As Boolean, EsTerceros As Boolean) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency
Dim Bonifica As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim vBonifica As Currency
Dim PorcBoni As Currency
Dim PorcComi As Currency

Dim HayPrecio As Boolean


Dim baseimpo As Currency
Dim BaseReten As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim PorcIva As Currency
Dim PorcReten As Currency
Dim ImpoAFO As Currency
Dim PorcAFO As Currency
Dim BaseAFO As Currency

Dim Gastos As Currency

Dim Anticipos As Currency

Dim TotalFac As Currency

Dim vSocio As cSocio
Dim vTipoMov As CTiposMov

Dim numfactu As Long


    On Error GoTo eFacturacion

    InformeAnticiposPicassentNew = False
    
    If EsTerceros Then
        tipoMov = "FAT" ' facturas de anticipos de terceros
    Else
        tipoMov = "FAA"
    End If
    
    BorrarTMPs
    b = CrearTMPs()
    If Not b Then
         Exit Function
    End If
    
    conn.BeginTrans
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie,"
    SQL = SQL & "rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.numalbar, rhisfruta.fecalbar, sum(coalesce(rhisfruta_clasif.kilosnet,0)) as kilosnet "
    SQL = SQL & " FROM  " & ctabla

    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.numalbar, rhisfruta.fecalbar "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.numalbar, rhisfruta.fecalbar "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        AntSocio = CStr(DBLet(RS!Codsocio, "N"))
        AntVarie = CStr(DBLet(RS!codvarie, "N"))
        AntCampo = CStr(DBLet(RS!codcampo, "N"))
        AntCalid = CStr(DBLet(RS!codcalid, "N"))
        
        ActSocio = CStr(DBLet(RS!Codsocio, "N"))
        ActVarie = CStr(DBLet(RS!codvarie, "N"))
        actCampo = CStr(DBLet(RS!codcampo, "N"))
        ActCalid = CStr(DBLet(RS!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0
                Importe = 0
                Bonifica = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                If EsTerceros Then
                    tipoMov = "FAT"
                Else
                    tipoMov = vSocio.CodTipomAnt
                End If
                
'                Set vTipoMov = New CTiposMov
'
'                numfactu = vTipoMov.ConseguirContador(tipoMov)
'                Do
'                    numfactu = vTipoMov.ConseguirContador(tipoMov)
'                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
'                    If devuelve <> "" Then
'                        'Ya existe el contador incrementarlo
'                        Existe = True
'                        vTipoMov.IncrementarContador (tipoMov)
'                        numfactu = vTipoMov.ConseguirContador(tipoMov)
'                    Else
'                        Existe = False
'                    End If
'                Loop Until Not Existe
'
'                vParamAplic.PrimFactAnt = numfactu
                
                
                numfactu = 1
                
            End If
        End If
    End If
    
    While Not RS.EOF And b
        ActCalid = DBLet(RS!codcalid, "N")
        ActVarie = DBLet(RS!codvarie, "N")
        actCampo = DBSet(RS!codcampo, "N")
        ActSocio = DBSet(RS!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            '[Monica]24/02/2014: añadida condicion
            If HayPrecio Then
                Kilos = Kilos + KilosCal
                Importe = Importe + vImporte
                Bonifica = Bonifica + vBonifica
                
                baseimpo = baseimpo + vImporte
                
                b = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte), CStr(vBonifica))
            End If
            KilosCal = 0
            vImporte = 0
            vBonifica = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            ' insertar linea de variedad, campo
            '[Monica]24/02/2014: añadida condicion
            If Kilos <> 0 Then
                b = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), "0", CStr(Bonifica))
            End If
            
            If b Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
                Bonifica = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            
           '[Monica]24/02/2014: añadida condicion
            If baseimpo <> 0 Then
            
                ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
            
                Select Case DBLet(vSocio.TipoIRPF, "N")
                    Case 0
                        ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                        BaseReten = (baseimpo + ImpoIva)
                        PorcReten = vParamAplic.PorcreteFacSoc
                    Case 1
                        ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                        BaseReten = baseimpo
                        PorcReten = vParamAplic.PorcreteFacSoc
                    Case 2
                        ImpoReten = 0
                        BaseReten = 0
                        PorcReten = 0
                End Select
            
                TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
                
                
                'insertar cabecera de factura
'                b = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
                
                SQL = "insert into tmpfact_fvarias (codtipom,numfactu,fecfactu,numfactufvar) select "
                SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & vSocio.Codigo & ")"
                conn.Execute SQL
                
                
'                '[Monica]24/12/2013: si es tercero he de marcarla como contabilizada
'                If EsTerceros Then
'                    If b Then b = MarcarFactura(tipoMov, CStr(numfactu), FecFac)
'                End If
                
                
'                If b Then b = InsertResumen(tipoMov, CStr(numfactu))
                
'                '[Monica]15/04/2013: Introducimos las facturas varias a descontar
'                If DescontarFVarias Then
'                    If b Then b = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, 0, 0)
'                End If
'
'                If b Then b = vTipoMov.IncrementarContador(tipoMov)
                numfactu = numfactu + 1

            Else
                b = True
                
            End If
                
            IncrementarProgresNew Pb1, 1
            
            
            If b Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    If EsTerceros Then
                        tipoMov = "FAT"
                    Else
                        tipoMov = vSocio.CodTipomAnt
                    End If
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
'                numfactu = vTipoMov.ConseguirContador(tipoMov)
'                Do
'                    numfactu = vTipoMov.ConseguirContador(tipoMov)
'                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
'                    If devuelve <> "" Then
'                        'Ya existe el contador incrementarlo
'                        Existe = True
'                        vTipoMov.IncrementarContador (tipoMov)
'                        numfactu = vTipoMov.ConseguirContador(tipoMov)
'                    Else
'                        Existe = False
'                    End If
'                Loop Until Not Existe
                numfactu = numfactu + 1
           End If
        End If
        
        
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim Precio As Currency
        
        Sql9 = "select precio1 from tmpinformes2 where importe1 = " & DBSet(RS!numalbar, "N") & " and importe2 = " & DBSet(RS!codvarie, "N")
        Sql9 = Sql9 & " and importe3  = " & DBSet(RS!codcalid, "N") & " and codusu = " & vUsu.Codigo
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            '[Monica]24/02/2014: añadida variable
            HayPrecio = True
            
            Precio = DBLet(Rs9.Fields(0).Value, "N")
            PorcBoni = 0
            PorcComi = 0
            ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
            If Precio > 0 Then
                PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(RS!codvarie, "N") & " and fechaent = " & DBSet(RS!Fecalbar, "F"))
                
                '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(RS!codcampo, "N"))
                If CCur(PorcComi) <> 0 Then
                    Precio = Precio - Round2(Precio * PorcComi / 100, 4)
                End If
            End If
        
            vPrecio = DBLet(Precio, "N")
            vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * Precio * (1 + (PorcBoni / 100)), 2)
            vBonifica = vBonifica + Round2(DBLet(RS!KilosNet, "N") * Precio * (1 + (PorcBoni / 100)), 2) - Round2(DBLet(RS!KilosNet, "N") * Precio, 2)
            
            KilosCal = KilosCal + DBLet(RS!KilosNet, "N")
            
        Else
            HayPrecio = False
        End If
        
        Set Rs9 = Nothing
        
        '[Monica]20/03/2014: miramos si hay precio para la calidad
        Sql9 = "select count(*) from tmpinformes2 where importe5 = " & DBSet(RS!codcampo, "N") & " and importe2 = " & DBSet(RS!codvarie, "N") & " and importeb1 = " & DBSet(RS!Codsocio, "N")
        Sql9 = Sql9 & " and importe3  = " & DBSet(RS!codcalid, "N") & " and codusu = " & vUsu.Codigo
        HayPrecio = (TotalRegistros(Sql9) <> 0)
        
        
        'hasta aqui
        HayReg = True
        
        RS.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If b And HayReg Then
        ' insertar linea de calidad
        '[Monica]24/02/2014: añadida condicion
        If HayPrecio Then
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            Bonifica = Bonifica + vBonifica
            
            baseimpo = baseimpo + vImporte
            
            If b Then b = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte), CStr(vBonifica))
        End If
        
        '[Monica]24/02/2014: añadida condicion
        If Kilos <> 0 Then
            ' insertar linea de variedad
            If b Then b = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), "0", CStr(Bonifica))
        End If
        
        '[Monica]24/02/2014: añadida condicion
        If baseimpo <> 0 Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
    
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
    '        BaseAFO = baseimpo
    '        PorcAFO = vParamAplic.PorcenAFO
    '        ImpoAFO = Round2(BaseAFO * PorcAFO / 100, 2)
    
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            
'            'insertar cabecera de factura
'            b = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
                SQL = "insert into tmpfact_fvarias (codtipom,numfactu,fecfactu,numfactufvar) select "
                SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & vSocio.Codigo & ")"
                conn.Execute SQL
            
            
'            '[Monica]24/12/2013: si es tercero he de marcarla como contabilizada
'            If EsTerceros Then
'                If b Then b = MarcarFactura(tipoMov, CStr(numfactu), FecFac)
'            End If
'
'
'            '[Monica]15/04/2013: Introducimos las facturas varias a descontar
'            If DescontarFVarias Then
'                If b Then b = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, 0, 0)
'            End If
'
'            If b Then b = InsertResumen(tipoMov, CStr(numfactu))
            
'            If b Then b = vTipoMov.IncrementarContador(tipoMov)
            numfactu = numfactu + 1
        
        Else
            b = True
        End If
        
        IncrementarProgresNew Pb1, 1
        
'        vParamAplic.UltFactAnt = numfactu
        
        'pasamos las temporales a las tablas
'        If b Then b = PasarTemporales()

        If b Then
            
            SQL = "update tmpliquidacion tt, tmpfact_fvarias ss,  tmpfact_variedad vv, tmpfact_calidad cc "
            SQL = SQL & " set tt.precio = cc.precio, tt.importe = cc.imporcal  "
            SQL = SQL & " where vv.codtipom = cc.codtipom and vv.numfactu = cc.numfactu and vv.fecfactu = cc.fecfactu "
            SQL = SQL & " and ss.codtipom = cc.codtipom and ss.numfactu = cc.numfactu and ss.fecfactu = cc.fecfactu"
            SQL = SQL & " and tt.codusu = " & vUsu.Codigo & " and  vv.codvarie = tt.codvarie and cc.codcalid = tt.codcalid"
            SQL = SQL & " and tt.codcampo = vv.codcampo and tt.codsocio = ss.numfactufvar "
        
            conn.Execute SQL
        End If
         
        
'        If b Then b = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        InformeAnticiposPicassentNew = False
    Else
        conn.CommitTrans
        InformeAnticiposPicassentNew = True
    End If
End Function


