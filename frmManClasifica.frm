VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManClasifica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clasificación de Campos"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13935
   Icon            =   "frmManClasifica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   4860
      TabIndex        =   95
      Text            =   "Text3"
      Top             =   7635
      Width           =   1300
   End
   Begin VB.Frame Frame2 
      Height          =   3525
      Index           =   0
      Left            =   135
      TabIndex        =   32
      Top             =   495
      Width           =   13635
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   9
         Tag             =   "Contrato|T|S|||rclasifica|contrato|||"
         Top             =   3120
         Width           =   795
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   4650
         TabIndex        =   94
         Text            =   "12345678901234567890"
         Top             =   1410
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   6390
         MaxLength       =   7
         TabIndex        =   17
         Tag             =   "Kilos Trans|N|N|||rclasifica|kilostra|###,##0||"
         Top             =   2400
         Width           =   1380
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Transportado por|N|N|0|1|rclasifica|transportadopor||N|"
         Top             =   840
         Width           =   1380
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   6390
         MaxLength       =   12
         TabIndex        =   16
         Tag             =   "Precio Estimado|N|S|||rclasifica|prestimado|###,##0.0000||"
         Top             =   2100
         Width           =   1380
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1170
         TabIndex        =   89
         Text            =   "12345678901234567890"
         Top             =   1755
         Width           =   4140
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   88
         Text            =   "12345678901234567890"
         Top             =   1410
         Width           =   2520
      End
      Begin VB.Frame Frame4 
         Caption         =   "Gastos Recolección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1500
         Left            =   8010
         TabIndex        =   78
         Top             =   135
         Width           =   5550
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   4140
            MaxLength       =   8
            TabIndex        =   26
            Tag             =   "Importe Penalización|N|S|||rclasifica|imppenal|#,##0.00||"
            Text            =   "123"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   2790
            MaxLength       =   8
            TabIndex        =   25
            Tag             =   "Importe Recolección|N|S|||rclasifica|imprecol|#,##0.00||"
            Text            =   "123"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   24
            Tag             =   "Importe Acarreo|N|S|||rclasifica|impacarr|#,##0.00||"
            Text            =   "123"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   90
            MaxLength       =   8
            TabIndex        =   23
            Tag             =   "Importe Transporte|N|S|||rclasifica|imptrans|#,##0.00||"
            Text            =   "123"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Tag             =   "Tipo Recolección|N|N|0|1|rclasifica|tiporecol||N|"
            Top             =   540
            Width           =   1350
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   4140
            MaxLength       =   3
            TabIndex        =   22
            Tag             =   "Nro.Trabajadores|N|S|||rclasifica|numtraba|##0||"
            Text            =   "123"
            Top             =   540
            Width           =   1305
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   18
            Left            =   2790
            MaxLength       =   8
            TabIndex        =   21
            Tag             =   "Horas Trabajadas|N|S|||rclasifica|horastra|#,##0.00||"
            Text            =   "123"
            Top             =   540
            Width           =   1305
         End
         Begin VB.Label Label5 
            Caption         =   "Imp.Penalización"
            Height          =   255
            Left            =   4140
            TabIndex        =   85
            Top             =   855
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Imp.Recolección"
            Height          =   255
            Left            =   2790
            TabIndex        =   84
            Top             =   855
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Imp.Acarreo"
            Height          =   255
            Left            =   1485
            TabIndex        =   83
            Top             =   855
            Width           =   1140
         End
         Begin VB.Label Label18 
            Caption         =   "Imp.Transporte"
            Height          =   255
            Left            =   90
            TabIndex        =   82
            Top             =   855
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Recolección"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   81
            Top             =   315
            Width           =   1260
         End
         Begin VB.Label Label8 
            Caption         =   "Nro.Trabajadores"
            Height          =   255
            Left            =   4185
            TabIndex        =   80
            Top             =   315
            Width           =   1230
         End
         Begin VB.Label Label7 
            Caption         =   "Horas Trabajadas"
            Height          =   255
            Left            =   2790
            TabIndex        =   79
            Top             =   315
            Width           =   1365
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   6390
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Fecha Albarán|F|S|||rclasifica|fecalbar|dd/mm/yyyy||"
         Top             =   3000
         Width           =   1380
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2730
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Entrada|F|N|||rclasifica|fechaent|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   240
         Width           =   1020
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   4500
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   840
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Tag             =   "Tipo Entrada|N|N|0|3|rclasifica|tipoentr||N|"
         Top             =   195
         Width           =   1380
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "Recolectado|N|N|0|1|rclasifica|recolect||N|"
         Top             =   525
         Width           =   1380
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1770
         TabIndex        =   69
         Top             =   2100
         Width           =   3540
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1170
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "Código Capataz|N|S|0|9999|rclasifica|codcapat|0000||"
         Top             =   2100
         Width           =   555
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2205
         TabIndex        =   68
         Top             =   2445
         Width           =   3090
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Código Transporte|T|S|||rclasifica|codtrans|||"
         Top             =   2445
         Width           =   1005
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   1755
         TabIndex        =   67
         Top             =   2790
         Width           =   3540
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1170
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "Código Tarifa|N|S|0|99|rclasifica|codtarif|00||"
         Top             =   2790
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1170
         MaxLength       =   8
         TabIndex        =   5
         Tag             =   "Nombre|N|N|||rclasifica|codcampo|00000000||"
         Text            =   "12345678"
         Top             =   1395
         Width           =   825
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   65
         Text            =   "12345678901234567890"
         Top             =   1020
         Width           =   3390
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   6390
         MaxLength       =   7
         TabIndex        =   18
         Tag             =   "Albarán|N|S|||rclasifica|numalbar|0000000||"
         Top             =   2700
         Width           =   1380
      End
      Begin VB.TextBox Text1 
         Height          =   1080
         Index           =   20
         Left            =   8040
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Tag             =   "Observaciones|T|S|||rclasifica|observac|||"
         Top             =   2040
         Width           =   5445
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   6390
         MaxLength       =   7
         TabIndex        =   15
         Tag             =   "Nro.Cajas|N|N|||rclasifica|numcajon|###,##0||"
         Top             =   1800
         Width           =   1380
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   51
         Text            =   "12345678901234567890"
         Top             =   660
         Width           =   3390
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Variedad|N|N|0|999999|rclasifica|codvarie|000000||"
         Text            =   "123456"
         Top             =   675
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   6390
         MaxLength       =   7
         TabIndex        =   14
         Tag             =   "Peso Neto|N|N|||rclasifica|kilosnet|###,##0||"
         Top             =   1500
         Width           =   1380
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   6390
         MaxLength       =   7
         TabIndex        =   13
         Tag             =   "Peso Bruto|N|N|||rclasifica|kilosbru|###,##0||"
         Top             =   1200
         Width           =   1380
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Nombre|N|N|||rclasifica|codsocio|000000||"
         Text            =   "123456"
         Top             =   1020
         Width           =   690
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nro.Nota|N|S|||rclasifica|numnotac|0000000|S|"
         Text            =   "1234567"
         Top             =   240
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   4500
         MaxLength       =   20
         TabIndex        =   87
         Tag             =   "Hora|FH|N|||rclasifica|horaentr|yyyy-mm-dd hh:mm:ss||"
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label9 
         Caption         =   "Contrato"
         Height          =   255
         Left            =   150
         TabIndex        =   96
         Top             =   3150
         Width           =   705
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   7770
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Trans."
         Height          =   255
         Index           =   6
         Left            =   5400
         TabIndex        =   93
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Transportado"
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   92
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Pr.Estimado"
         Height          =   255
         Index           =   5
         Left            =   5400
         TabIndex        =   91
         Top             =   2130
         Width           =   1185
      End
      Begin VB.Label Label6 
         Caption         =   "Población"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   90
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label15 
         Caption         =   "F.Albarán"
         Height          =   255
         Left            =   5400
         TabIndex        =   77
         Top             =   3000
         Width           =   705
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   6120
         Picture         =   "frmManClasifica.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   1980
         TabIndex        =   76
         Top             =   270
         Width           =   480
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2475
         Picture         =   "frmManClasifica.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label26 
         Caption         =   "Hora"
         Height          =   255
         Left            =   4050
         TabIndex        =   75
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Entrada"
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   74
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label11 
         Caption         =   "Recolectado"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   73
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label23 
         Caption         =   "Capataz"
         Height          =   255
         Left            =   150
         TabIndex        =   72
         Top             =   2175
         Width           =   705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   900
         ToolTipText     =   "Buscar Capataz"
         Top             =   2100
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Transp."
         Height          =   255
         Left            =   150
         TabIndex        =   71
         Top             =   2505
         Width           =   705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   900
         ToolTipText     =   "Buscar Transportista"
         Top             =   2460
         Width           =   240
      End
      Begin VB.Label Label12 
         Caption         =   "Tarifa"
         Height          =   255
         Left            =   150
         TabIndex        =   70
         Top             =   2850
         Width           =   705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   900
         ToolTipText     =   "Buscar Tarifa"
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   900
         ToolTipText     =   "Buscar Campo"
         Top             =   1410
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Campo"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   66
         Top             =   1425
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   900
         ToolTipText     =   "Buscar Socio"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Albarán"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   62
         Top             =   2730
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Cajas"
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   60
         Top             =   1815
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Variedad"
         Height          =   255
         Left            =   150
         TabIndex        =   52
         Top             =   705
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   900
         ToolTipText     =   "Buscar Variedad"
         Top             =   690
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Neto"
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   50
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Bruto"
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   49
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label Label6 
         Caption         =   "Socio"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   48
         Top             =   1050
         Width           =   690
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   8070
         TabIndex        =   37
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   9195
         ToolTipText     =   "Zoom descripción"
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Nota"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   33
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6210
      TabIndex        =   58
      Text            =   "Text3"
      Top             =   7635
      Width           =   1300
   End
   Begin VB.Frame FrameAux1 
      Caption         =   "Incidencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3360
      Left            =   8055
      TabIndex        =   43
      Top             =   4140
      Width           =   5700
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   290
         Index           =   1
         Left            =   1755
         MaskColor       =   &H00000000&
         TabIndex        =   57
         ToolTipText     =   "Buscar Incidencia"
         Top             =   2610
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   1980
         TabIndex        =   56
         Text            =   "nombre"
         Top             =   2610
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   9
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   45
         Tag             =   "Incidencia|N|N|||rclasifica_incidencia|codincid|0000|S|"
         Text            =   "inci"
         Top             =   2610
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   8
         Left            =   360
         MaxLength       =   16
         TabIndex        =   44
         Tag             =   "Nro.Nota|N|N|||rclasifica_incidencia|numnotac|0000000|S|"
         Text            =   "codfor"
         Top             =   2610
         Visible         =   0   'False
         Width           =   555
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   90
         TabIndex        =   46
         Top             =   270
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmManClasifica.frx":0122
         Height          =   2610
         Index           =   1
         Left            =   90
         TabIndex        =   47
         Top             =   630
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   4604
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   1
         Left            =   1350
         Top             =   180
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "AdoAux(1)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Clasificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3360
      Left            =   150
      TabIndex        =   38
      Top             =   4140
      Width           =   7815
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   5805
         MaxLength       =   7
         TabIndex        =   64
         Tag             =   "Kilos Neto|N|S|||rclasifica_clasif|kilosnet|###,##0||"
         Text            =   "neto"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   2
         Left            =   3420
         MaskColor       =   &H00000000&
         TabIndex        =   86
         ToolTipText     =   "Buscar Calidad"
         Top             =   2565
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   3060
         MaxLength       =   2
         TabIndex        =   61
         Tag             =   "Calidad|N|N|||rclasifica_clasif|codcalid|00|S|"
         Text            =   "Ca"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3645
         TabIndex        =   55
         Text            =   "Calidad"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   945
         MaskColor       =   &H00000000&
         TabIndex        =   54
         ToolTipText     =   "Buscar Envase"
         Top             =   2565
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1155
         TabIndex        =   53
         Top             =   2565
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   63
         Tag             =   "Muestra|N|S|||rclasifica_clasif|muestra|###,##0.00||"
         Text            =   "muestra"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   495
         MaxLength       =   6
         TabIndex        =   40
         Tag             =   "Variedad|N|N|||rclasifica_clasif|codvarie|000000|S|"
         Text            =   "Var"
         Top             =   2565
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   45
         MaxLength       =   16
         TabIndex        =   39
         Tag             =   "Nro.Nota|N|N|||rclasifica_clasif|numnotac|0000000|S|"
         Text            =   "nota"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   41
         Top             =   225
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   3720
         Top             =   225
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "AdoAux(0)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmManClasifica.frx":013A
         Height          =   2615
         Index           =   0
         Left            =   135
         TabIndex        =   42
         Top             =   630
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   4604
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   135
      TabIndex        =   30
      Top             =   7545
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12705
      TabIndex        =   29
      Top             =   7590
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11610
      TabIndex        =   28
      Top             =   7605
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1980
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cálculo Gastos"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Importar Excel"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   8520
         TabIndex        =   36
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12720
      TabIndex        =   34
      Top             =   7620
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label11 
      Caption         =   "TOTAL NETO: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   3120
      TabIndex        =   59
      Top             =   7665
      Width           =   1365
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnGastos 
         Caption         =   "&Cálculo Gastos"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnExporImpor 
         Caption         =   "Importar"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManClasifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: CLIENTES                  -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes ' campos del socio
Attribute frmMens.VB_VarHelpID = -1
'Private WithEvents frmExp As frmExpImpExcel ' Exportacion o importacion a pagina excel

'Private WithEvents frmArt As frmManArtic 'articulos
Private WithEvents frmVar As frmComVar 'variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCap As frmManCapataz 'capataces
Attribute frmCap.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTranspor 'tranportistas
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarTra 'tarifas de transporte
Attribute frmTar.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden 'incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCampos 'campos
Attribute frmCam.VB_VarHelpID = -1
'
'*****************************************************
Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de llínies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim Gastos As Boolean

Dim CodTipoMov As String
Dim NotaExistente As Boolean

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim VarieAnt As String
Dim MuestraAnt As String


Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
'        Case 0 'Variedades
'            Set frmVar = New frmComVar
'            frmVar.DatosADevolverBusqueda = "0|1|"
'            frmVar.CodigoActual = txtAux(1).Text
'            frmVar.Show vbModal
'            Set frmVar = Nothing
'            PonerFoco txtAux(1)
        Case 1 'Incidencia
            indice = 9
            Set frmInc = New frmManInciden
            frmInc.DatosADevolverBusqueda = "0|1|"
            frmInc.CodigoActual = txtAux(9).Text
            frmInc.Show vbModal
            Set frmInc = Nothing
            PonerFoco txtAux(9)
        Case 2 'calidades
            indice = Index
            Set frmCal = New frmManCalidades
            frmCal.DatosADevolverBusqueda = "2|3|"
            frmCal.CodigoActual = txtAux(2).Text
            frmCal.Show vbModal
            Set frmCal = Nothing
            PonerFoco txtAux(2)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                NotaExistente = False
                InsertarCabecera
            
            
'                If InsertarDesdeForm2(Me, 1) Then
'                    text2(9).Text = PonerNombreCuenta(text1(9), Modo, text1(0).Text)
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
'                If Gastos Then CalcularGastos

                If ModificaCabecera Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
                    PonerCampos
                End If
'[Monica]24/06/2010
'                CalcularGastos
'
'                If ModificaDesdeFormulario2(Me, 1) Then
'                    TerminaBloquear
'                    PosicionarData
'                End If
'24/06/2010
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                    
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        '050505
                          CalcularKilosNetos
                          CalcularGastos

'                        PosicionarData
'                        PasarSigReg
                    Else
'                        PonerFoco txtAux(12)
                    End If
            End Select
            'nuevo calculamos los totales de lineas
'            CalcularTotales
                    

        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If PrimeraVez Then PrimeraVez = False
    If PrimeraVez Then
        PrimeraVez = False
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                BotonAnyadir
            Else
                PonerModo 1 'búsqueda
                ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
                Text1(0).BackColor = vbYellow 'nro nota
                ' ****************************************************************************
            End If
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
 
    ' ICONETS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        .Buttons(11).Image = 31   'Expandir Añadir, Borrar y Modificar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 35  'Importar clasificacion a excel
        .Buttons(14).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
    ' ***********************************
    
    'cargar IMAGES .Image =de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
'    Me.imgBuscar(8).Picture = frmPpal.imgListComun.ListImages(21).Picture
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i
   
    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    CodTipoMov = "NOC"
    
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rclasifica"
    Ordenacion = " ORDER BY numnotac"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where numnotac=-1"
    Data1.Refresh
       
    CargaGrid 0, False
    CargaGrid 1, False
       
    ModoLineas = 0
    '[Monica]04/10/2016: nuevo contrato de Coopic
    Text1(24).Enabled = (vParamAplic.Cooperativa = 16)
    Text1(24).visible = (vParamAplic.Cooperativa = 16)
    Label9.visible = (vParamAplic.Cooperativa = 16)
    
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'búsqueda
'        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
'        Text1(0).BackColor = vbYellow 'codforfait
'        ' ****************************************************************************
'    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la capçalera ***
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    Me.Combo1(2).ListIndex = -1
    Me.Combo1(3).ListIndex = -1
    ' *****************************************

    Text3(0).Text = ""
    Text3(1).Text = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    Text1(5).Enabled = True
    Combo1(1).Enabled = True
    
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    '---------------------------------------------
    
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    b = (Modo <> 1) And (Modo <> 3)
    'Campos Nº entrada bloqueado y en azul
    BloquearTxt Text1(0), b, True
    
    
    '*** si n'hi han combos a la capçalera ***
    BloquearCombo Me, Modo
    '**************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    BloquearImgFec Me, 0, Modo, ModoLineas
    
    imgFec(1).Enabled = (Modo = 1)
    imgFec(1).visible = (Modo = 1)
    BloquearTxt Text1(2), Not (Modo = 3 Or Modo = 4 Or Modo = 1)
    BloquearTxt Text1(12), Not (Modo = 1)
    BloquearTxt Text1(13), Not (Modo = 1)
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    
    PonerLongCampos
    
'    Frame4.Enabled = (Modo = 1)
'  cambiado por esto
    For i = 15 To 19
        Text1(i).Enabled = (Modo = 1)
    Next i
    Combo1(2).Enabled = (Modo = 1)
    Text1(14).Enabled = (Modo = 1) Or (Modo = 4)
'
    
    Text1(3).Enabled = (Modo = 1) Or (Modo = 3) Or (Modo = 4)
    imgBuscar(0).Enabled = (Modo = 1) Or (Modo = 3) Or (Modo = 4)
    imgBuscar(0).visible = (Modo = 1) Or (Modo = 3) Or (Modo = 4)

    
    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = b
      
    ' ****** si n'hi han combos a la capçalera ***********************
     If (Modo = 0) Or (Modo = 2) Or (Modo = 5) Then
        Combo1(0).Enabled = False
        Combo1(0).BackColor = &H80000018 'groc
        Combo1(1).Enabled = False
        Combo1(1).BackColor = &H80000018 'groc
        Combo1(2).Enabled = False
        Combo1(2).BackColor = &H80000018 'groc
        Combo1(3).Enabled = False
        Combo1(3).BackColor = &H80000018 'groc
    ElseIf (Modo = 1) Or (Modo = 3) Or (Modo = 4) Then
        Combo1(0).Enabled = True
        Combo1(0).BackColor = &H80000005 'blanc
        Combo1(1).Enabled = True
        Combo1(1).BackColor = &H80000005 'blanc
        Combo1(2).Enabled = True
        Combo1(2).BackColor = &H80000005 'blanc
        Combo1(3).Enabled = True
        Combo1(3).BackColor = &H80000005 'blanc
    End If
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    'Expandir operaciones
    Toolbar1.Buttons(11).Enabled = b
    Me.mnGastos.Enabled = b
    
    'Exportar / Importar a pagina excel
    Toolbar1.Buttons(13).Enabled = (Modo = 0 Or Modo = 2) And vParamAplic.Cooperativa = 4
    Me.mnExporImpor.Enabled = (Modo = 0 Or Modo = 2) And vParamAplic.Cooperativa = 4
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 2) And Not DeConsulta
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Adoaux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'CLASIFICACION
            Sql = "SELECT rclasifica_clasif.numnotac, rclasifica_clasif.codvarie, variedades.nomvarie, rclasifica_clasif.codcalid,"
            Sql = Sql & " rcalidad.nomcalid, rclasifica_clasif.muestra, rclasifica_clasif.kilosnet "
            Sql = Sql & " from rclasifica_clasif, variedades, rcalidad "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE rclasifica_clasif.numnotac = -1"
            End If
            Sql = Sql & " and rclasifica_clasif.codvarie = variedades.codvarie "
            Sql = Sql & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
            Sql = Sql & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
            Sql = Sql & " ORDER BY rclasifica_clasif.codvarie"
               
        Case 1 'INCIDENCIAS
            Sql = "SELECT rclasifica_incidencia.numnotac, rclasifica_incidencia.codincid, rincidencia.nomincid "
            Sql = Sql & " FROM rclasifica_incidencia, rincidencia "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE rclasifica_incidencia.numnotac = -1"
            End If
            Sql = Sql & " and rclasifica_incidencia.codincid = rincidencia.codincid"
            Sql = Sql & " ORDER BY rclasifica_incidencia.codincid"
            
    End Select
    
    MontaSQLCarga = Sql
End Function


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Calidad
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codcalid
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
'Campos
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codcampo
End Sub

Private Sub frmCap_DatoSeleccionado(CadenaSeleccion As String)
'Capataces
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1) 'codcapat
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Incidencias
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codincid
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    PonerDatosCampo Text1(5)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Socios
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
'Tarifas
    Text1(8).Text = RecuperaValor(CadenaSeleccion, 1) 'codtarifa
    Text2(8).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Transportistas
    Text1(7).Text = RecuperaValor(CadenaSeleccion, 1) 'codtranspor
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(20).Text = vCampo
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           CargaCadenaAyuda vCadena, Index
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
End Sub

Private Sub imgFec_Click(Index As Integer)
   
   Screen.MousePointer = vbHourglass
   
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

   
   frmC.NovaData = Now
   Select Case Index
        Case 0
            indice = 1
        Case 1
            indice = 13
   End Select
   
   Me.imgFec(0).Tag = indice
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmC.NovaData = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmC.Show vbModal
   Set frmC = Nothing
   PonerFoco Text1(indice)

End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 20
        frmZ.pTitulo = "Observaciones de la Clasificación"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub

Private Sub mnBuscar_Click()
Dim i As Integer
    BotonBuscar
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1 'quan busque, per defecte no seleccione cap tipo de client
    Next i
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnExporImpor_Click()
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'        Set frmExp = New frmExpImpExcel
'
'        frmExp.Show vbModal
'
'        Set frmExp = Nothing
'    End If
    If vParamAplic.Cooperativa = 4 Then
    
        MsgBox "Se va a proceder a la importación de la clasificación. " & vbCrLf & vbCrLf & "Cuando termine el proceso recuerde realizar el recálculo de gastos." & vbCrLf, vbExclamation
    
        Shell App.Path & "\clasificacion.exe /B|" & vUsu.CadenaConexion & "|", vbNormalFocus
    End If

End Sub

Private Sub mnGastos_Click()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean


    If vParamAplic.Cooperativa = 4 Then
        If MsgBox("¿Quiere realizar el cálculo de gastos masivo?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            Sql = "select * from rclasifica order by numnotac"
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF And b
                b = ActualizarGastos(Rs, "")
                Rs.MoveNext
            Wend
            Set Rs = Nothing
            If b Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                mnVerTodos_Click
            End If
        Else
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonGastos
        End If
    
    Else

        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonGastos
        
    End If
End Sub

Private Sub mnImprimir_Click()
'    Screen.MousePointer = vbHourglass
'    frmListConfeccion.Show vbModal
'    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 3  'Búscar
           mnBuscar_Click
        Case 4  'Tots
            mnVerTodos_Click
        Case 7  'Nou
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
        Case 11  'Expandir operaciones
            mnGastos_Click
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13 'Exportar/importar
            mnExporImpor_Click
        Case 14   'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
        For i = 0 To Combo1.Count - 1
            Combo1(i).ListIndex = -1
        Next i
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()
    
    If Text1(2).Text <> "" Then
        Text1(21).Text = Text1(2).Text
        Text1(21).Tag = Replace(Text1(21).Tag, "FH", "FHH")
    End If

    CadB = ObtenerBusqueda2(Me, 1)
    
    Text1(21).Tag = Replace(Text1(21).Tag, "FHH", "FH")
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim Cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    Cad = ""
    Cad = Cad & ParaGrid(Text1(0), 12, "Nro.Nota")
    Cad = Cad & ParaGrid(Text1(1), 15, "Fecha")
    Cad = Cad & "Socio|nomsocio|T||30·"
'    cad = cad & ParaGrid(text1(2), 60, "Descripción")
    Cad = Cad & "Variedad|nomvarie|T||30·"
    Cad = Cad & ParaGrid(Text1(5), 15, "Campo")
    
    If Cad <> "" Then
        
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        Cad = "(" & NombreTabla & " inner join variedades on rclasifica.codvarie = variedades.codvarie) inner join rsocios on rclasifica.codsocio = rsocios.codsocio "
        frmB.vtabla = Cad 'NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Clasificación" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 1

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonVerTodos()
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = ""
    End If
    '********************************************************************


    '[Monica]24/10/2013: cuando me estan dando de alta miro si
'    If vParamAplic.NroNotaManual Then
'        PonerFoco Text1(0)
'    Else
'        PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
'    End If
    Text1(1).Text = Now
    Text1(2).Text = Mid(Now, 12, 8)
    ' ***********************************************************
       
    PosicionarCombo Combo1(0), 0
    PosicionarCombo Combo1(1), 0
    PosicionarCombo Combo1(3), 0
            
    Text1(0) = NumF
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()

    Gastos = False


    VarieAnt = Text1(3).Text

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
End Sub

Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar la Clasificación?"
    Cad = Cad & vbCrLf & "Número: " & Data1.Recordset.Fields(0)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Proveedor", Err.Description
End Sub

Private Sub BotonGastos()
Dim i As Integer

    Gastos = True

    '[Monica]20/07/2010 inicializada variable variedad anterior pq le damos a modificar
    VarieAnt = Text1(3).Text
    '
    
    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    For i = 0 To 13
        BloquearTxt Text1(i), True
    Next i
    BloquearTxt Text1(20), True
    imgFec(0).Enabled = False
    imgFec(1).Enabled = False
    For i = 0 To 5
        BloquearImage imgBuscar(i), True
    Next i
    BloquearCmb Combo1(0), True
    BloquearCmb Combo1(1), True
    BloquearCmb Combo1(3), True
    
    For i = 14 To 16
        BloquearTxt Text1(i), True
        Text1(i).Enabled = False
    Next i
    
    
    ' desbloqueamos el frame de gastos
    Frame4.Enabled = True
    
    For i = 17 To 19
        BloquearTxt Text1(i), False
        Text1(i).Enabled = True
    Next i
    
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    Combo1(2).SetFocus
End Sub




Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    
    Text1(2).Text = Mid(Text1(21).Text, 12, 8)
    
    
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    VisualizarDatosCampo Data1.Recordset!codcampo
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For i = 0 To 1
        CargaGrid i, True
        If Not Adoaux(i).Recordset.EOF Then _
            PonerCamposForma2 Me, Adoaux(i), 2, "FrameAux" & i
    Next i

    Text1(2).Text = Mid(Text1(21).Text, 12, 8)
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    Text2(3).Text = PonerNombreDeCod(Text1(3), "variedades", "nomvarie")
    Text2(4).Text = PonerNombreDeCod(Text1(4), "rsocios", "nomsocio")
'    Text2(5).Text = PonerNombreDeCod(Text1(8), "rcampos", "nomcapac")
    Text2(6).Text = PonerNombreDeCod(Text1(6), "rcapataz", "nomcapat")
    Text2(7).Text = PonerNombreDeCod(Text1(7), "rtransporte", "nomtrans")
    Text2(8).Text = PonerNombreDeCod(Text1(8), "rtarifatra", "nomtarif")
    ' ********************************************************************************
    
'    CalcularTotales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco Text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'txtAux2(2).text = ""

                    End If

'                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)

                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        Adoaux(NumTabMto).Recordset.MoveFirst
                    End If
                    
                    CalcularKilosNetos
                    CalcularGastos

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = Adoaux(NumTabMto).Recordset.Fields(3) 'el 2 es el nº de llinia
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(3).Name & " =" & V)
                        ' ***************************************************************
                    End If
                    CalcularKilosNetos
                    CalcularGastos
            End Select
            
            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not Adoaux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    Text1(21).Text = Format(Text1(1).Text, "dd/mm/yyyy") & " " & Format(Text1(2).Text, "HH:MM:SS")
    
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
'    If (Modo = 3) Then 'insertar
'        'comprobar si existe ya el cod. del campo clave primaria
'        If ExisteCP(Text1(0)) Then b = False
'    End If
    
    If Modo = 3 Or Modo = 4 Then
        Select Case Combo1(0).ListIndex
            Case 0  'caja
                Select Case Combo1(1).ListIndex
                    Case 0  'unidad
                        If CCur(Text1(5).Text) = 0 Then
                            MsgBox "El campo Kilos/Unidad debe tener un valor superior a cero", vbExclamation
                            b = False
                        End If
                    Case 1  'kilo
                        If CCur(Text1(4).Text) = 0 Then
                            MsgBox "El campo Kilos/Caja debe tener un valor superior a cero", vbExclamation
                            b = False
                        End If
                End Select
            Case 1  'kilo
                If CCur(Text1(4).Text) = 0 Then
                    MsgBox "El campo Kilos/Caja debe tener un valor superior a cero", vbExclamation
                    b = False
                End If
            
        End Select
        
        If b Then
            If Not EsCampoSocioVariedad(Text1(5).Text, Text1(4).Text, Text1(3).Text) Then
                MsgBox "El campo no es del socio o de la variedad introducida. Revise.", vbExclamation
                b = False
            End If
        End If
    End If
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(numnotac=" & DBSet(Text1(0).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE numnotac=" & Data1.Recordset!Numnotac
    
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rclasifica_clasif " & vWhere
        
    conn.Execute "DELETE FROM rclasifica_incidencia " & vWhere
        
    'Eliminar la CAPÇALERA
    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
  
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim b As Boolean

    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Sql As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'numero de nota
            If PonerFormatoEntero(Text1(Index)) Then
            
                '[Monica]24/10/2013: comprobamos si existe el nro de nota
                If Text1(Index).Text <> "" And Modo = 3 Then
                    If ExisteNota(Text1(Index).Text) Then
                        MsgBox "Número de Nota ya existe. Reintroduzca.", vbExclamation
                        PonerFoco Text1(Index)
                    End If
                End If
                
            End If
        Case 3 'Variedad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "variedades", "nomvarie")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmComVar
                        frmVar.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    If (Modo = 3 Or Modo = 4) And EsVariedadGrupo6(Text1(Index).Text) Then
                        MsgBox "Esta variedad es del Grupo de Bodega. Revise.", vbExclamation
                        PonerFoco Text1(Index)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Socio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
'                        frmSoc.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmSoc.Show vbModal
                        Set frmSoc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    If EstaSocioDeAlta(Text1(Index)) Then
                        PonerCamposSocioVariedad
                    Else
                        MsgBox "El socio está dado de baja. Reintroduzca.", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
                
                
        Case 5 'campo
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then Exit Sub
                Sql = ""
                Sql = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", Text1(Index).Text, "N")
                If Sql = "" Then
                    cadMen = "No existe el Campo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCam = New frmManCampos
                        frmCam.DatosADevolverBusqueda = "0|1|"
'                        frmCamp.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCam.Show vbModal
                        Set frmCam = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    If Not EstaCampoDeAlta(Text1(Index).Text) Then
                        MsgBox "El campo no está dado de alta. Reintroduzca.", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    Else
                        VisualizarDatosCampo (Text1(Index))
                    End If
                End If
            End If
        
        
        
        Case 6 'Capataz
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rcapataz", "nomcapat")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Capataz: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCap = New frmManCapataz
                        frmCap.DatosADevolverBusqueda = "0|1|"
                        frmCap.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCap.Show vbModal
                        Set frmCap = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 7 'Transportista
            If Text1(Index).Text <> "" Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rtransporte", "nomtrans")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Transportista: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra = New frmManTranspor
                        frmTra.DatosADevolverBusqueda = "0|1|"
                        frmTra.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTra.Show vbModal
                        Set frmTra = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
            
        Case 8 'Tarifa de transporte
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rtarifatra", "nomtarif")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Tarifa de Transporte: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTar = New frmManTarTra
                        frmTar.DatosADevolverBusqueda = "0|1|"
                        frmTar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTar.Show vbModal
                        Set frmTar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 14, 15, 16, 17 'importes de gastos
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(Index), 3
            
        '[Monica]24/10/2013: desdoblado
        Case 10, 11, 23 'kilos bruot, cajas , kilosneto
            PonerFormatoEntero Text1(Index)
        Case 9 'kilos bruot, cajas , kilosneto
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 And vParamAplic.Cooperativa = 14 Then
                    Text1(10).Text = Text1(Index).Text
                    Text1(11).Text = Text1(Index).Text
                    Text1(23).Text = Text1(Index).Text
                End If
            End If
        
        Case 12 ' numero de albaran
            PonerFormatoEntero Text1(Index)
        
        Case 18 ' horas trabajadas
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(Index), 3
            
        Case 19 ' nro de trabajadores
            PonerFormatoEntero Text1(Index)
        
        Case 2 'formato hora
            If Modo = 1 Then Exit Sub
            PonerFormatoHora Text1(Index)
       
        Case 1, 13 ' formato fecha
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            PonerFormatoFecha Text1(Index), True
       
        Case 22 ' precio estimado
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(Index), 11
            
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 5) And (Modo = 3 Or Modo = 4) Then
        If Not EsCampoSocioVariedad(Text1(5).Text, Text1(4).Text, Text1(3).Text) Then
            MsgBox "El campo no es del socio o de la variedad introducida. Revise.", vbExclamation
        End If
    End If
    
    
    
    ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 7: KEYBusqueda KeyAscii, 0 'envase
                Case 8: KEYBusqueda KeyAscii, 1 'capacidad
                Case 9: KEYBusqueda KeyAscii, 2 'medida
                Case 10: KEYBusqueda KeyAscii, 3 'confeccion
                Case 11: KEYBusqueda KeyAscii, 4 'presentacion
                Case 12: KEYBusqueda KeyAscii, 5 'marca
                Case 13: KEYBusqueda KeyAscii, 6 'palet
                Case 2: KEYBusqueda KeyAscii, 7 'variedad
            End Select
        End If
    Else
'        If Index <> 3 Or (Index = 3 And Text1(3).Text = "") Then KEYpress KeyAscii
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub



'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
    'End If
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If Adoaux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'calidades
            Sql = "¿Seguro que desea eliminar la Calidad?"
            Sql = Sql & vbCrLf & "Calidad: " & Adoaux(Index).Recordset!codcalid
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rclasifica_clasif "
                Sql = Sql & vWhere & " AND codvarie= " & Adoaux(Index).Recordset!codvarie
                Sql = Sql & " and codcalid= " & Adoaux(Index).Recordset!codcalid
            End If
            
        Case 1 'incidencias
            Sql = "¿Seguro que desea eliminar la Incidencia?"
            Sql = Sql & vbCrLf & "Nombre: " & Adoaux(Index).Recordset!nomincid
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rclasifica_incidencia "
                Sql = Sql & vWhere & " AND codincid= " & Adoaux(Index).Recordset!codincid
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        If Index = 0 Then
            CalcularKilosNetos
            CalcularGastos
        End If
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
'        SituarTab (NumTabMto + 1)
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "rclasifica_clasif"
        Case 1: vtabla = "rclasifica_incidencia"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
'            If Index = 0 Then NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'calidades
                    txtAux(0).Text = Text1(0).Text 'numnotac
                    txtAux(1).Text = Text1(3).Text 'codvarie
                    txtAux(2).Text = ""
                    txtAux2(2).Text = ""
                    txtAux(3).Text = ""
                    txtAux(4).Text = "0"
                    BloquearTxt txtAux(2), False
                    BloquearTxt txtAux(3), False
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    Me.btnBuscar(0).Enabled = False
                    Me.btnBuscar(0).visible = False
                    PonerFoco txtAux(2)
                Case 1 'incidencias
                    txtAux(8).Text = Text1(0).Text 'numnotac
                    txtAux(9).Text = "" 'NumF 'codcoste
                    txtAux2(9).Text = ""
                    For i = 9 To 9
                        BloquearTxt txtAux(i), False
                    Next i
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
                    PonerFoco txtAux(9)
            End Select
            
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
    If Adoaux(Index).Recordset.EOF Then Exit Sub
    If Adoaux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
  
    Select Case Index
        Case 0, 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 0 ' muestra
        
            For J = 0 To 1
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            txtAux2(0).Text = DataGridAux(Index).Columns(2).Text
            txtAux(2).Text = DataGridAux(Index).Columns(3).Text
            txtAux2(2).Text = DataGridAux(Index).Columns(4).Text
            txtAux(3).Text = DataGridAux(Index).Columns(5).Text
            txtAux(4).Text = DataGridAux(Index).Columns(6).Text
            BloquearTxt txtAux(0), True
            BloquearTxt txtAux(1), True
            BloquearTxt txtAux(2), True
            BloquearTxt txtAux(4), True
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
            
        Case 1 'incidencias
            For J = 8 To 9
                txtAux(J).Text = DataGridAux(Index).Columns(J - 8).Text
            Next J
            txtAux2(9).Text = DataGridAux(Index).Columns(2).Text
            For i = 9 To 9
                BloquearTxt txtAux(i), True
            Next i
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'muestras
            PonerFoco txtAux(3)
            
            MuestraAnt = txtAux(3).Text
            
        Case 1 'incidencias
            PonerFoco txtAux(9)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'muestras
             For jj = 2 To 4
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            For jj = 2 To 2
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            btnBuscar(2).visible = b
            btnBuscar(2).Top = alto
            
        Case 1 'incidencias
            For jj = 9 To 9
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            txtAux2(9).visible = b
            txtAux2(9).Top = alto
            btnBuscar(1).visible = b
            btnBuscar(1).Top = alto
            
    End Select
End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de entrada
    Combo1(0).AddItem "Normal"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "V.Campo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "P.Integrado"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Ind.Directo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "Retirada"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    Combo1(0).AddItem "Venta Comercio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 5
    
    
    'recolectado por
    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    'tipo de recoleccion
    Combo1(2).AddItem "Horas"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Destajo"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    
    'transportado por
    Combo1(3).AddItem "Cooperativa"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 0
    Combo1(3).AddItem "Socio"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 1
    
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2 ' codigo de calidad
            If txtAux(Index) <> "" Then
                txtAux2(2).Text = DevuelveDesdeBDNew(cAgro, "rcalidad", "nomcalid", "codvarie", txtAux(1).Text, "N", , "codcalid", txtAux(2).Text, "N")
                If txtAux2(2).Text = "" Then
                    cadMen = "No existe la Calidad: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCal = New frmManCalidades
                        frmCal.DatosADevolverBusqueda = "0|1|"
                        frmCal.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCal.Show vbModal
                        Set frmCal = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(2).Text = ""
            End If
        
        Case 9 ' codigo de incidencia
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rincidencia", "nomincid")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Código de Incidencia: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmInc = New frmManInciden
                        frmInc.DatosADevolverBusqueda = "0|1|"
                        frmInc.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmInc.Show vbModal
                        Set frmInc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    cmdAceptar.SetFocus
                End If
            Else
                txtAux2(Index).Text = ""
            End If
        
'        Case 2 ' codigo de calidad
'            PonerFormatoEntero txtAux(Index)
            
        Case 3  ' muestra
            PonerFormatoDecimal txtAux(Index), 3
            
        Case 4 ' kilosnetos
            PonerFormatoEntero txtAux(Index)
            
            cmdAceptar.SetFocus
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    ' si no estamos en muestra salimos
    If Index <> 3 Then Exit Sub
    
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
'050509
'            cmdAceptar_Click
            ModificarLinea
            
'            If Me.DataGridAux(0).Bookmark > 0 Then
'                DataGridAux(0).Bookmark = DataGridAux(0).Bookmark - 1
'            End If
            PasarAntReg
        Case 40 'Desplazamiento Flecha Hacia Abajo
            'ModificarExistencia
'050509
'            cmdAceptar_Click
            ModificarLinea
            
            PasarSigReg
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
                Select Case Index
                    Case 2: 'calidad
                        KeyAscii = 0
                        btnBuscar_Click (2)
                    Case 9: 'incidencia
                        KeyAscii = 0
                        btnBuscar_Click (1)
                End Select
            End If
        Else
            If Index = 3 Then ' estoy introduciendo la muestra
               If KeyAscii = 13 Then 'ENTER
                    PonerFormatoDecimal txtAux(Index), 3
                    If ModoLineas = 2 Then
                        '050509 cmdAceptar_Click 'ModificarExistencia
                        ModificarLinea

                        PasarSigReg
                    End If
                    If ModoLineas = 1 Then
                        cmdAceptar.SetFocus
                    End If
                    
                    '050509
'                    If ModoLineas = 1 Then
'                        cmdAceptar.SetFocus
'                    End If
               ElseIf KeyAscii = 27 Then
                    cmdCancelar_Click 'ESC
               End If
            Else
                KEYpress KeyAscii
            End If
        End If
    End If
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    ' ******************************************************************************
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    indice = Index + 3
     Select Case Index
        Case 0 'variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(3).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(3)
        Case 1 'socios
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(4).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(4)
        Case 2 'campos
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|"
'            frmCam.CodigoActual = Text1(5).Text
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(5)
        Case 3 'Capataces
            Set frmCap = New frmManCapataz
            frmCap.DatosADevolverBusqueda = "0|1|"
            frmCap.CodigoActual = Text1(6).Text
            frmCap.Show vbModal
            Set frmCap = Nothing
            PonerFoco Text1(6)
        Case 4 'Transportista
            Set frmTra = New frmManTranspor
            frmTra.DatosADevolverBusqueda = "0|1|"
            frmTra.CodigoActual = Text1(7).Text
            frmTra.Show vbModal
            Set frmTra = Nothing
            PonerFoco Text1(7)
        Case 5 'Tarifa
            Set frmTar = New frmManTarTra
            frmTar.DatosADevolverBusqueda = "0|1|"
            frmTar.CodigoActual = Text1(8).Text
            frmTar.Show vbModal
            Set frmTar = Nothing
            PonerFoco Text1(8)
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'cuentas bancarias
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
'                    txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
                End If
                
            Case 1 'departamentos
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(21).Text = DataGridAux(Index).Columns(5).Text
'                    txtAux(22).Text = DataGridAux(Index).Columns(6).Text
'                    txtAux(23).Text = DataGridAux(Index).Columns(8).Text
'                    txtAux(24).Text = DataGridAux(Index).Columns(15).Text
'                    txtAux2(22).Text = DataGridAux(Index).Columns(7).Text
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'cuentas bancarias
'                txtAux(11).Text = ""
'                txtAux(12).Text = ""
            Case 1 'departamentos
                For i = 21 To 24
'                   txtAux(i).Text = ""
                Next i
'               txtAux2(22).Text = ""
            Case 2 'Tarjetas
'               txtAux(50).Text = ""
'               txtAux(51).Text = ""
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
'Private Sub SituarTab(numTab As Integer)
'    On Error Resume Next
'
'    SSTab1.Tab = numTab
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub
' **********************************************

Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
Dim tip As Integer
Dim i As Byte

    Adoaux(Index).ConnectionString = conn
    Adoaux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    Adoaux(Index).CursorType = adOpenDynamic
    Adoaux(Index).LockType = adLockPessimistic
    Adoaux(Index).Refresh
    
    If Not Adoaux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, Adoaux(Index), 2, "FrameAux" & Index
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
    End If
End Sub

' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'clasificacion
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;" 'numnotac
            tots = tots & "N|txtAux(1)|T|Variedad|800|;N|btnBuscar(0)|B|||;N|txtAux2(0)|T|Nombre|2000|;"
            tots = tots & "S|txtAux(2)|T|Calidad|1000|;S|btnBuscar(2)|B|||;S|txtAux2(2)|T|Nombre|3200|;"
            tots = tots & "S|txtAux(3)|T|Muestra|1400|;S|txtAux(4)|T|Peso Neto|1400|;"
            
            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(0).Columns(3).Alignment = dbgLeft
            DataGridAux(0).Columns(5).NumberFormat = "###,##0.00"
            DataGridAux(0).Columns(5).Alignment = dbgRight
            DataGridAux(0).Columns(6).NumberFormat = "###,##0"
            DataGridAux(0).Columns(6).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 1 'incidencias
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtAux(9)|T|Incidencia|1000|;S|btnBuscar(1)|B||195|;"
            tots = tots & "S|txtAux2(9)|T|Denominación|3700|;"

            arregla tots, DataGridAux(Index), Me
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
''        LimpiarCamposFrame Index
'    End If
    CalcularTotales
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'envases
        Case 1: nomframe = "FrameAux1" 'costes
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            Select Case NumTabMto
                Case 0 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    
                    CalcularKilosNetos
                    CalcularGastos

                    
                    b = BLOQUEADesdeFormulario2(Me, Data1, 1)
                    If b Then BotonAnyadirLinea NumTabMto
                Case 1
                    b = BLOQUEADesdeFormulario2(Me, Data1, 1)
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
           
'            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'envases
        Case 1: nomframe = "FrameAux1" 'costes
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
    
        If NumTabMto = 0 Then
            If MuestraAnt <> "" And ComprobarCero(txtAux(3).Text) <> 0 Then
                If MsgBox("¿Desea acumular muestras?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    txtAux(3).Text = CCur(ImporteSinFormato(txtAux(3).Text)) + CCur(ImporteSinFormato(MuestraAnt))
                    txtAux(3).Text = Format(txtAux(3).Text, "###,##0.00")
                End If
            End If
        End If
    
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            Select Case NumTabMto
                Case 0
                    V = Adoaux(NumTabMto).Recordset.Fields(3) 'el 2 es el nº de llinia
                Case 1
                    V = Adoaux(NumTabMto).Recordset.Fields(2) 'el 2 es el nº de llinia
            End Select
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(3).Name & " =" & V)
            
            LLamaLineas NumTabMto, 0
            ModificarLinea = True
        End If
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " numnotac=" & Me.Data1.Recordset!Numnotac
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
'    Select Case Index
'        Case 0 'Cuentas Bancarias
'            txtAux(11).Text = ""
'            txtAux(12).Text = ""
'        Case 1 'Departamentos
'            txtAux(21).Text = ""
'            txtAux(22).Text = ""
'            txtAux2(22).Text = ""
'            txtAux(23).Text = ""
'            txtAux(24).Text = ""
'        Case 2 'Tarjetas
'            txtAux(50).Text = ""
'            txtAux(51).Text = ""
'        Case 4 'comisiones
'            txtAux2(2).Text = ""
'    End Select
'
    If Err.Number <> 0 Then Err.Clear
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub


Private Sub CalcularTotales()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency

    On Error Resume Next

    If Data1.Recordset.EOF Or Modo = 1 Then
        Text3(0).Text = ""
        Text3(1).Text = ""
        Exit Sub
    End If

    'total kilosnetos
    Sql = "select sum(kilosnet) from rclasifica_clasif "
    Sql = Sql & " where numnotac = " & Data1.Recordset!Numnotac
    
    
    Text3(0).Text = TotalRegistros(Sql)
    Valor = CCur(TransformaPuntosComas(Text3(0).Text))
    If Valor <> 0 Then
        Text3(0).Text = Format(Valor, "###,###,##0")
    Else
        Text3(0).Text = ""
    End If
    
    'total muestra
    Sql = "select sum(muestra) from rclasifica_clasif "
    Sql = Sql & " where numnotac = " & Data1.Recordset!Numnotac
    
    
    Text3(1).Text = TotalRegistros(Sql)
    Valor = CCur(TransformaPuntosComas(Text3(1).Text))
    If Valor <> 0 Then
        Text3(1).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(1).Text = ""
    End If
    
    
    
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

Private Sub CalcularGastos()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency
Dim GasRecol As Currency
Dim GasAcarreo As Currency
Dim KilosTria As Long
Dim KilosNet As Long
Dim EurDesta As Currency
Dim EurRecol As Currency
Dim PrecAcarreo As Currency
Dim i As Integer

    On Error Resume Next
    
    GasRecol = 0
    GasAcarreo = 0
    
    If Combo1(0).ListIndex = 1 Then
        For i = 14 To 19
            Text1(i).Text = ""
        Next i
        Exit Sub
    End If
    
    
    Sql = "select eurdesta, eurecole from variedades where codvarie = " & DBSet(Text1(3).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        EurDesta = DBLet(Rs.Fields(0).Value, "N")
        EurRecol = DBLet(Rs.Fields(1).Value, "N")
    End If

    Set Rs = Nothing

'    Sql = "select sum(kilosnet) from rclasifica_clasif where numnotac = " & DBSet(Text1(0).Text, "N")
'    KilosNet = TotalRegistros(Sql)

    KilosNet = CLng(ImporteSinFormato(Text1(10).Text))

    '[Monica]14/10/2010: para picassent los kilos son los de transporte
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then KilosNet = CLng(ImporteSinFormato(ComprobarCero(Text1(23).Text)))


    'recolecta socio
    If Combo1(1).ListIndex = 1 Then
        Sql = "select sum(kilosnet) from rclasifica_clasif, rcalidad  where numnotac = " & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
        Sql = Sql & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
        Sql = Sql & " and rcalidad.gastosrec = 1"
        
        KilosTria = TotalRegistros(Sql)
        
        GasRecol = Round2(KilosTria * EurRecol, 2)
        
        
    Else
    'recolecta cooperativa
        If Combo1(2).ListIndex = 0 Then
            'horas
            'gastosrecol = horas * personas * rparam.(costeshora + costesegso)
            GasRecol = Round2(HorasDecimal(Text1(18).Text) * CCur(Text1(19).Text) * (vParamAplic.CosteHora + vParamAplic.CosteSegSo), 2)
        Else
            'destajo
            GasRecol = Round2(KilosNet * EurDesta, 2)
        End If
    End If
    
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then GasRecol = Round2(KilosNet * EurDesta, 2)
    
    
'12/05/2009
'    If Text1(8).Text <> "" Then
'        sql = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", Text1(8).Text, "N")
'        PrecAcarreo = CCur(sql)
'    Else
'        PrecAcarreo = 0
'    End If
'12/05/2009 cambiado por esto pq si que hay tarifa 0

    PrecAcarreo = 0
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", Text1(8).Text, "N")
    If Sql <> "" Then
        PrecAcarreo = CCur(Sql)
    End If
    
    If vParamAplic.Cooperativa = 4 Then
        Sql = "select sum(kilosnet) from rclasifica_clasif, rcalidad  where numnotac = " & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
        Sql = Sql & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
        Sql = Sql & " and rcalidad.gastosrec = 1"
        
        KilosTria = TotalRegistros(Sql)
        
        If Combo1(3).ListIndex = 1 Then ' transportado por socio
            GasAcarreo = Round2(PrecAcarreo * KilosTria, 2)
        Else
            GasAcarreo = Round2(PrecAcarreo * KilosNet, 2)
        End If
        ' cargamos los kilos de transporte
        Text1(23).Text = Format(KilosTria, "###,##0")
    Else
        GasAcarreo = Round2(PrecAcarreo * KilosNet, 2)
    End If
    
    Text1(16).Text = Format(GasRecol, "#,##0.00")
    Text1(15).Text = Format(GasAcarreo, "#,##0.00")
    

End Sub

'Private Function HorasDecimal(cantidad As String) As Currency
'Dim Entero As Long
'Dim vCantidad As String
'Dim vDecimal As String
'Dim vEntero As String
'Dim vHoras As Currency
'Dim J As Integer
'    HorasDecimal = 0
'
'    vCantidad = ImporteSinFormato(cantidad)
'
'    J = InStr(1, vCantidad, ",")
'
'    If J > 0 Then
'        vEntero = Mid(vCantidad, 1, J - 1)
'        vDecimal = Mid(vCantidad, J + 1, Len(vCantidad))
'    Else
'        vEntero = vCantidad
'        vDecimal = ""
'    End If
'
'    vHoras = (CLng(vEntero) * 60) + CLng(vDecimal)
'
'    HorasDecimal = Round2(vHoras / 60, 2)
'
'End Function


Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If Me.DataGridAux(0).Bookmark < Me.Adoaux(0).Recordset.RecordCount Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGridAux(0).Bookmark = DataGridAux(0).Bookmark + 1
        BotonModificarLinea 0
    ElseIf DataGridAux(0).Bookmark = Adoaux(0).Recordset.RecordCount Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificarLinea 0
    End If
End Sub


Private Sub PasarAntReg()
'Nos situamos en el siguiente registro
    If Me.DataGridAux(0).Bookmark > 1 Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGridAux(0).Bookmark = DataGridAux(0).Bookmark - 1
        BotonModificarLinea 0
    ElseIf DataGridAux(0).Bookmark = 1 Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificarLinea 0
    End If
End Sub


Private Sub CalcularKilosNetos()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim TotalKilos As Currency
Dim TotalMuestra As Currency
Dim Calidad As Integer
Dim vSQL As String

Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency
Dim GasRecol As Currency
Dim GasAcarreo As Currency
Dim KilosTria As Long
Dim KilosNet As Long
Dim EurDesta As Currency
Dim EurRecol As Currency
Dim PrecAcarreo As Currency
Dim i As Integer
Dim KilosNetos As Long


    On Error GoTo eCalcularKilosNetos
    
    Sql = "select sum(muestra) from rclasifica_clasif where numnotac = " & Me.Data1.Recordset!Numnotac
'[Monica]14/10/2011: lo dejo en la clasificacion automatica
'    If vParamAplic.Cooperativa = 0 Then
'        SQL = SQL & " and codcalid not in (select codcalid from rcalidad where codvarie = " & Me.Data1.Recordset!CodVarie
'        SQL = SQL & " and tipcalid in (1,3)) " ' muestras que no sean de destrio ni de merma
'    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        TotalMuestra = DBLet(Rs.Fields(0).Value, "N")
    End If

    Set Rs = Nothing
    
    If TotalMuestra = 0 Then
        Sql = "update rclasifica_clasif set kilosnet = " & ValorNulo & " where numnotac = " & Me.Data1.Recordset!Numnotac
        conn.Execute Sql
        
        CargaGrid 0, True
        BotonGastos
        cmdAceptar_Click
        PosicionarData
        
        Exit Sub
    End If

    Sql = "select rclasifica_clasif.* from rclasifica_clasif where numnotac = " & Me.Data1.Recordset!Numnotac
    '050509
'    sql = sql & " and muestra <> 0"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalKilos = 0
    Calidad = 0
    
'[Monica]14/10/2011: lo dejamos como estaba
'    If vParamAplic.Cooperativa = 0 Then
'        KilosNetos = Me.Data1.Recordset!KilosNet - DevuelveValor("select sum(muestra) from rclasifica_clasif where numnotac = " & Me.Data1.Recordset!numnotac & " and codcalid in (select codcalid from rcalidad where codvarie = " & Data1.Recordset!CodVarie & " and tipcalid in (1,3))")
'    End If
    
    While Not Rs.EOF
'[Monica]14/10/2011: se queda como estaba
'        '[Monica]11/10/2011: si es Catadau quiere que los kilos que pongo en la muestra si es destrio o merma me coincidan
'        If vParamAplic.Cooperativa = 0 Then ' Catadau
'            If EsCalidadDestrio(CStr(Me.Data1.Recordset!CodVarie), CStr(DBLet(RS.Fields!codcalid, "N"))) Or _
'               EsCalidadMerma(CStr(Me.Data1.Recordset!CodVarie), CStr(DBLet(RS.Fields!codcalid, "N"))) Then
'
'               KilosNet = DBLet(RS!Muestra, "N")
'            Else
'               KilosNet = Round2(DBLet(RS!Muestra, "N") * KilosNetos / TotalMuestra, 0)
'            End If
'        Else
            ' como estaba para todos
            KilosNet = Round2(DBLet(Rs!Muestra, "N") * Me.Data1.Recordset!KilosNet / TotalMuestra, 0)
            
'        End If
        
        TotalKilos = TotalKilos + KilosNet
        Calidad = DBLet(Rs!codcalid, "N")
        
        vSQL = "update rclasifica_clasif set kilosnet = " & DBSet(KilosNet, "N", "S")
        vSQL = vSQL & ", muestra = " & DBSet(Rs!Muestra, "N", "S")
        vSQL = vSQL & " where numnotac = " & DBSet(Rs!Numnotac, "N")
        vSQL = vSQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
        vSQL = vSQL & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        conn.Execute vSQL
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing

    
    'redondeamos en la ultima calidad o en la destrio
    If TotalKilos <> Me.Data1.Recordset!KilosNet Then
        '[Monica]28/06/2011: si es Quatretonda la calidad de redondeo es la de maxima muestra no la de destrio
        If vParamAplic.Cooperativa = 7 Then
            vSQL = CalidadMaximaMuestraenClasificacion(Me.Data1.Recordset!codvarie, Me.Data1.Recordset!Numnotac, True)
        Else
            vSQL = CalidadDestrioenClasificacion(Me.Data1.Recordset!codvarie, Me.Data1.Recordset!Numnotac, True)
        End If
        
        If vSQL <> "" Then Calidad = CInt(vSQL)
    
        Sql = "update rclasifica_clasif set kilosnet = kilosnet + (" & (Me.Data1.Recordset!KilosNet - TotalKilos) & ") "
        Sql = Sql & " where numnotac = " & Data1.Recordset!Numnotac
        Sql = Sql & " and codvarie = " & Data1.Recordset!codvarie
        Sql = Sql & " and codcalid = " & DBSet(Calidad, "N")
    
        conn.Execute Sql
    End If
    
    
    CargaGrid 0, True
    BotonGastos
    cmdAceptar_Click
    PosicionarData
    Exit Sub
    
eCalcularKilosNetos:
    MuestraError Err.Number, "Calcular Kilos Netos", Err.Description
End Sub

Private Sub VisualizarDatosCampo(campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codcampo = " & DBSet(campo, "N")
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla, rcampos.nrocampo from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    If Not Rs.EOF Then
        Text1(5).Text = campo
        PonerFormatoEntero Text1(5)
        Text2(0).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text2(1).Text = DBLet(Rs!desPobla, "T")        ' nombre de la poblacion
        Text2(2).Text = DBLet(Rs!NroCampo, "T")        ' nombre de la poblacion
    End If
    
    Set Rs = Nothing
    
End Sub


Private Sub PonerCamposSocioVariedad()
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If Text1(3).Text = "" Or Text1(4).Text = "" Then Exit Sub
    
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codsocio = " & DBSet(Text1(4).Text, "N") & " and rcampos.fecbajas is null"
    Cad = Cad & " and rcampos.codvarie = " & DBSet(Text1(3).Text, "N")
     
    Cad1 = "select count(*) from rcampos where " & Cad
     
    NumRegis = TotalRegistros(Cad1)
    
    If NumRegis = 0 Then Exit Sub
    If NumRegis = 1 Then
        Cad1 = "select codcampo from rcampos where " & Cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text1(5).Text = DBLet(Rs.Fields(0).Value)
            PonerDatosCampo Text1(5).Text
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadWHERE = " and " & Cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens.campo = Text1(5).Text
        frmMens.OpcionMensaje = 6
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
    
End Sub

Private Sub PonerDatosCampo(campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla, rcampos.nrocampo from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text1(5).Text = ""
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    If Not Rs.EOF Then
        Text1(5).Text = campo
        PonerFormatoEntero Text1(5)
        Text2(0).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text2(1).Text = DBLet(Rs.Fields(4).Value, "T") ' descripcion de poblacion
        Text2(2).Text = DBLet(Rs.Fields(5).Value, "T") ' nro de campo
    End If
    
    Set Rs = Nothing
    
End Sub

Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String
Dim actualiza As Boolean
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        If Text1(0).Text = "" Then
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            actualiza = True
        Else
            actualiza = False
        End If
        
        Sql = CadenaInsertarDesdeForm(Me)
        If InsertarOferta(Sql, vTipoMov, actualiza) Then
            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            PonerCadenaBusqueda
            PonerModo 2
        
            If Not NotaExistente Then
                Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                PosicionarData
                BotonModificar
                cmdAceptar_Click
            End If
        
        End If
    End If
    Text1(0).Text = Format(Text1(0).Text, "0000000")
End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov, ActualizarContador As Boolean) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
Dim Sql2 As String

Dim Rs As ADODB.Recordset
Dim Sql3 As String
Dim cadMen As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Factura
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        If ExisteNota(Text1(0).Text) Then
            devuelve = Text1(0).Text
        Else
            devuelve = ""
        End If
        
'        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numnotac", "numnotac", Text1(0).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            If ActualizarContador Then
                vTipoMov.IncrementarContador (CodTipoMov)
                Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            Else
                MsgBox "Número de Nota ya existe. Reintroduzca.", vbExclamation
                PonerFoco Text1(0)
                NotaExistente = True
                InsertarOferta = False
                Exit Function
            End If
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Factura (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    cadMen = ""
    Sql3 = "select * from rclasifica where numnotac = " & DBSet(Text1(0).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        bol = InsertarClasificacion(Rs, cadMen, "")
        cadMen = "Insertando Clasificacion: " & cadMen
'26-05-2009: Santi no quiere que se calcule el transporte quiere meterlo él
'           cuando se da de alta.
'        If bol Then
'            bol = ActualizarTransporte(Rs, cadMen)
'            cadMen = "Actualizando Gastos de Transporte" & cadMen
'        End If
'26-05-2009
    End If
    
    Set Rs = Nothing
    
    If bol Then
        If ActualizarContador Then
            MenError = "Error al actualizar el contador de la Factura."
            vTipoMov.IncrementarContador (CodTipoMov)
        End If
    End If
    
    MenError = MenError & cadMen
    
EInsertarOferta:
    If Err.Number <> 0 Or Not bol Then
        MenError = "Insertando Entrada." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        InsertarOferta = True
    Else
        conn.RollbackTrans
        InsertarOferta = False
    End If
End Function

Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = " numnotac= " & Text1(0).Text
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function

Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String
Dim Sql As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    b = True
    
    If CLng(VarieAnt) <> CLng(Text1(3).Text) Then
        Sql = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Text1(0).Text, "N")
        Sql = Sql & "  and kilosnet <> 0 "
        
        If TotalRegistros(Sql) <> 0 Then
            MsgBox "La entrada está clasificada, no se pueden modificar la variedad.", vbExclamation
            conn.RollbackTrans
            ModificaCabecera = False
            Exit Function
        Else
            MenError = "Eliminando lineas"
            Sql = "delete from rclasifica_clasif where numnotac = " & DBSet(Text1(0).Text, "N")
            conn.Execute Sql
            
            MenError = "Insertando nueva clasificación"
            Sql = "insert into rclasifica_clasif (numnotac,codvarie,codcalid,muestra,kilosnet)"
            Sql = Sql & " select " & DBSet(Text1(0).Text, "N") & ",codvarie, codcalid, null, null from rcalidad "
            Sql = Sql & " where codvarie = " & DBSet(Text1(3).Text, "N")
            
            conn.Execute Sql
            
        End If
    End If
    
    '[Monica]08/02/2012: si modifican variedad o socio o campo o fecha u hora y tienen traza
    If b And (CLng(Data1.Recordset!codvarie) <> CLng(Text1(3).Text) Or CLng(Data1.Recordset!Codsocio) <> CLng(Text1(4).Text) Or CLng(Data1.Recordset!codcampo) <> CLng(Text1(5).Text) Or _
             CStr(Data1.Recordset!FechaEnt) <> Text1(1).Text Or CStr(Data1.Recordset!horaentr) <> Text1(21).Text) Then
          MenError = "Actualizar Traza: "
          b = ActualizarTraza(Text1(0).Text, Text1(3).Text, Text1(4).Text, Text1(5).Text, Text1(1).Text, Text1(21).Text, MenError)
    End If
    
    If b Then CalcularGastos
        
    If b Then b = ModificaDesdeFormulario1(Me, 1) 'ModificaDesdeFormulario2(Me, 2, "Frame2")

EModificarCab:
    If Err.Number <> 0 Or Not b Then
        MenError = "Modificando Entrada." & vbCrLf & "----------------------------" & vbCrLf & MenError
        If Err.Number <> 0 Then
            MuestraError Err.Number, MenError, Err.Description
        Else
            MsgBox MenError, vbExclamation
        End If
        b = False
    End If
    If b Then
        ModificaCabecera = True
        conn.CommitTrans
    Else
        ModificaCabecera = False
        conn.RollbackTrans
    End If
End Function


