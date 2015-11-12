VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPOZHidrantesIndefa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hidrantes Indefa"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15540
   Icon            =   "frmPOZHidrantesIndefa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   15540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   210
      TabIndex        =   23
      Top             =   480
      Width           =   15195
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   13590
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Fecha Lectura Actual|F|S|||rpozos|fech_act|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   900
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   10290
         MaxLength       =   7
         TabIndex        =   16
         Tag             =   "Contador Actual|N|S|||rpozos|lect_act|######0||"
         Text            =   "1234567"
         Top             =   870
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   10290
         MaxLength       =   7
         TabIndex        =   14
         Tag             =   "Lectura Anterior|N|N|||rpozos|lect_ant|######0||"
         Text            =   "1234567"
         Top             =   540
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   13590
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Fecha lectura anterior|F|S|||rpozos|fech_ant|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   540
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Index           =   11
         Left            =   9300
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "Observaciones|T|S|||rpozos|observac|||"
         Top             =   1950
         Width           =   5745
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   8
         Tag             =   "Parcelas|T|S|||rpozos|parcelas||N|"
         Text            =   "1234567890123456789012345"
         Top             =   2250
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Polígono|T|S|||rpozos|poligono||N|"
         Text            =   "1234567890"
         Top             =   1935
         Width           =   1035
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3390
         MaxLength       =   40
         TabIndex        =   64
         Top             =   1935
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "Partida|N|N|1|9999|rpozos|codparti|0000||"
         Top             =   1620
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   62
         Top             =   1620
         Width           =   4305
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2220
         MaxLength       =   30
         TabIndex        =   60
         Top             =   1290
         Width           =   4305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Pozo|N|N|0|999|rpozos|codpozo|000||"
         Top             =   1290
         Width           =   840
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   4
         Tag             =   "Campo|N|S|1|99999999|rpozos|codcampo|00000000||"
         Top             =   975
         Width           =   840
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   8100
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Hanegadas|N|S|0|9999.99|rpozos|hanegada|###0.0000||"
         Text            =   "1234567890"
         Top             =   1290
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   8100
         MaxLength       =   4
         TabIndex        =   12
         Tag             =   "Calibre|N|S|||rpozos|calibre|###0|N|"
         Text            =   "1234"
         Top             =   1620
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   8100
         MaxLength       =   7
         TabIndex        =   13
         Tag             =   "Acciones|N|S|||rpozos|nroacciones|#,###,##0|N|"
         Text            =   "1234567"
         Top             =   1950
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   8100
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Fecha Alta|F|N|||rpozos|fechaalta|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   390
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   8100
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Fecha Alta|F|S|||rpozos|fechabaja|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   750
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Socio|N|N|1|999999|rpozos|codsocio|000000||"
         Top             =   660
         Width           =   840
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2220
         MaxLength       =   30
         TabIndex        =   52
         Top             =   660
         Width           =   4305
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   6270
         MaxLength       =   1
         TabIndex        =   2
         Tag             =   "Digito Control|T|N|||rpozos|digcontrol|||"
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   3810
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Toma|N|S|0|999999|rpozos|nroorden|000000||"
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Contador TCH|T|N|||rpozos|hidrante||S|"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   10290
         MaxLength       =   9
         TabIndex        =   111
         Tag             =   "Consumo|N|S|||rpozos|consumo|########0||"
         Text            =   "1234567"
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   10170
         X2              =   11490
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label14 
         Caption         =   "Consumo"
         Height          =   255
         Left            =   9330
         TabIndex        =   110
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   13290
         Picture         =   "frmPOZHidrantesIndefa.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha Lectura"
         Height          =   255
         Left            =   12060
         TabIndex        =   109
         Top             =   900
         Width           =   1065
      End
      Begin VB.Label Label9 
         Caption         =   "Actual"
         Height          =   255
         Left            =   9330
         TabIndex        =   108
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label23 
         Caption         =   "Anterior"
         Height          =   255
         Left            =   9330
         TabIndex        =   107
         Top             =   570
         Width           =   1125
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   13290
         Picture         =   "frmPOZHidrantesIndefa.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha Lectura"
         Height          =   255
         Left            =   12060
         TabIndex        =   106
         Top             =   570
         Width           =   1200
      End
      Begin VB.Label Label180 
         Caption         =   "Lecturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9330
         TabIndex        =   105
         Top             =   180
         Width           =   810
      End
      Begin VB.Line Line3 
         X1              =   9330
         X2              =   14970
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Line Line2 
         X1              =   9330
         X2              =   14970
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   9300
         TabIndex        =   68
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   10500
         Tag             =   "-1"
         ToolTipText     =   "Zoom descripción"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Parcelas"
         Height          =   255
         Left            =   210
         TabIndex        =   67
         Top             =   2250
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "Poligono"
         Height          =   255
         Left            =   210
         TabIndex        =   66
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label Label15 
         Caption         =   "Población"
         Height          =   255
         Left            =   2550
         TabIndex        =   65
         Top             =   1935
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1020
         ToolTipText     =   "Buscar Partida"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Partida"
         Height          =   255
         Left            =   210
         TabIndex        =   63
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Pozo"
         Height          =   255
         Left            =   210
         TabIndex        =   61
         Top             =   1305
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1020
         ToolTipText     =   "Buscar Pozo"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Campo"
         Height          =   255
         Left            =   210
         TabIndex        =   59
         Top             =   960
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1020
         ToolTipText     =   "Buscar Campo"
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label41 
         Caption         =   "Hanegadas"
         Height          =   255
         Left            =   6900
         TabIndex        =   58
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Calibre"
         Height          =   255
         Left            =   6900
         TabIndex        =   57
         Top             =   1665
         Width           =   810
      End
      Begin VB.Label Label8 
         Caption         =   "Acciones"
         Height          =   255
         Left            =   6900
         TabIndex        =   56
         Top             =   1950
         Width           =   930
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   7800
         Picture         =   "frmPOZHidrantesIndefa.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   390
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Alta"
         Height          =   255
         Left            =   6900
         TabIndex        =   55
         Top             =   420
         Width           =   870
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   7800
         Picture         =   "frmPOZHidrantesIndefa.frx":01AD
         ToolTipText     =   "Buscar fecha"
         Top             =   750
         Width           =   240
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha Baja"
         Height          =   255
         Left            =   6900
         TabIndex        =   54
         Top             =   780
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1020
         ToolTipText     =   "Buscar Socio"
         Top             =   675
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Socio"
         Height          =   255
         Left            =   210
         TabIndex        =   53
         Top             =   645
         Width           =   690
      End
      Begin VB.Label Label16 
         Caption         =   "Dígito Control"
         Height          =   255
         Left            =   5190
         TabIndex        =   39
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label26 
         Caption         =   "Toma"
         Height          =   255
         Left            =   3015
         TabIndex        =   29
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Contador TCH"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   270
      TabIndex        =   21
      Top             =   8910
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
         TabIndex        =   22
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   14370
      TabIndex        =   20
      Top             =   9090
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   13260
      TabIndex        =   19
      Top             =   9090
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   15540
      _ExtentX        =   27411
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
            Object.ToolTipText     =   "Buscar Diferencias Indefa"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar Registro"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
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
         Left            =   11160
         TabIndex        =   27
         Top             =   60
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   14370
      TabIndex        =   25
      Top             =   9090
      Visible         =   0   'False
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5700
      Left            =   270
      TabIndex        =   28
      Top             =   3150
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   10054
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Indefa"
      TabPicture(0)   =   "frmPOZHidrantesIndefa.frx":0238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Coopropietarios"
      TabPicture(1)   =   "frmPOZHidrantesIndefa.frx":0254
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Campos"
      TabPicture(2)   =   "frmPOZHidrantesIndefa.frx":0270
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux1"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   5205
         Left            =   120
         TabIndex        =   70
         Top             =   420
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   9181
         _Version        =   393216
         Tabs            =   5
         Tab             =   1
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Contadores"
         TabPicture(0)   =   "frmPOZHidrantesIndefa.frx":028C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "FrameAux2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Hidrantes"
         TabPicture(1)   =   "frmPOZHidrantesIndefa.frx":02A8
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "FrameAux4"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Válvulas"
         TabPicture(2)   =   "frmPOZHidrantesIndefa.frx":02C4
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FrameAux5"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Ventosas"
         TabPicture(3)   =   "frmPOZHidrantesIndefa.frx":02E0
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "FrameAux6"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Desagües"
         TabPicture(4)   =   "frmPOZHidrantesIndefa.frx":02FC
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "FrameAux3"
         Tab(4).ControlCount=   1
         Begin VB.Frame FrameAux3 
            BorderStyle     =   0  'None
            Height          =   4695
            Left            =   -74880
            TabIndex        =   202
            Top             =   420
            Width           =   14280
            Begin VB.TextBox txtaux8 
               BackColor       =   &H80000018&
               Height          =   555
               Index           =   8
               Left            =   2010
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   213
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   4080
               Width           =   12135
            End
            Begin VB.TextBox txtaux8 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   7
               Left            =   2010
               MaxLength       =   250
               ScrollBars      =   2  'Vertical
               TabIndex        =   212
               Tag             =   "the_geom|T|S|||rae_visitas_hidtomas|the_geom||N|"
               Top             =   3180
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   5
               Left            =   2010
               MaxLength       =   50
               TabIndex        =   211
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|observaciones||N|"
               Top             =   2250
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   4
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   210
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_emisor_impulsos||N|"
               Top             =   1785
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   3
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   209
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_contador||N|"
               Top             =   1350
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   2
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   208
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_vavulas_3vias||N|"
               Top             =   915
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   1
               Left            =   2010
               MaxLength       =   10
               TabIndex        =   207
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|dn_valvula_esfera||N|"
               Top             =   555
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   0
               Left            =   2010
               MaxLength       =   10
               TabIndex        =   206
               Tag             =   "dn_contador|T|S|||rae_visitas_hidtomas|dn_contador||N|"
               Top             =   210
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   6
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   205
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   2715
               Width           =   2865
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   3
               Left            =   2550
               Top             =   4110
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
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   2
               Left            =   5310
               TabIndex        =   203
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   3
               Left            =   9750
               TabIndex        =   204
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label51 
               Caption         =   "Opertividad de la vávula"
               Height          =   315
               Left            =   30
               TabIndex        =   224
               Top             =   2295
               Width           =   1905
            End
            Begin VB.Label Label49 
               Caption         =   "Tapa arqueta"
               Height          =   315
               Left            =   30
               TabIndex        =   223
               Top             =   3195
               Width           =   1815
            End
            Begin VB.Label Label48 
               Caption         =   "Punto de desagüe"
               Height          =   315
               Left            =   30
               TabIndex        =   222
               Top             =   1845
               Width           =   1815
            End
            Begin VB.Label Label47 
               Caption         =   "Tipo de vávula"
               Height          =   315
               Left            =   30
               TabIndex        =   221
               Top             =   1395
               Width           =   1815
            End
            Begin VB.Label Label46 
               Caption         =   "Observaciones"
               Height          =   315
               Left            =   30
               TabIndex        =   220
               Top             =   4170
               Width           =   1635
            End
            Begin VB.Label Label44 
               Caption         =   "Tipo de arqueta"
               Height          =   315
               Left            =   30
               TabIndex        =   219
               Top             =   2745
               Width           =   2175
            End
            Begin VB.Image Image9 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   10740
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
            Begin VB.Image Image8 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   6210
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
            Begin VB.Label Label39 
               Caption         =   "INTERIOR"
               Height          =   315
               Left            =   9750
               TabIndex        =   218
               Top             =   60
               Width           =   855
            End
            Begin VB.Label Label37 
               Caption         =   "EXTERIOR"
               Height          =   315
               Left            =   5310
               TabIndex        =   217
               Top             =   60
               Width           =   825
            End
            Begin VB.Label Label34 
               Caption         =   "NUDO +"
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
               Height          =   315
               Left            =   30
               TabIndex        =   216
               Top             =   915
               Width           =   1875
            End
            Begin VB.Label Label33 
               Caption         =   "SECTOR"
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
               Height          =   315
               Left            =   30
               TabIndex        =   215
               Top             =   240
               Width           =   1875
            End
            Begin VB.Label Label32 
               Caption         =   "NUDO -"
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
               Height          =   315
               Left            =   30
               TabIndex        =   214
               Top             =   555
               Width           =   1875
            End
         End
         Begin VB.Frame FrameAux6 
            BorderStyle     =   0  'None
            Height          =   4710
            Left            =   -74880
            TabIndex        =   175
            Top             =   420
            Width           =   14280
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   555
               Index           =   10
               Left            =   2010
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   186
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   4080
               Width           =   12135
            End
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   7
               Left            =   2010
               MaxLength       =   250
               ScrollBars      =   2  'Vertical
               TabIndex        =   185
               Tag             =   "the_geom|T|S|||rae_visitas_hidtomas|the_geom||N|"
               Top             =   2880
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   5
               Left            =   2010
               MaxLength       =   50
               TabIndex        =   184
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|observaciones||N|"
               Top             =   2130
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   4
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   183
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_emisor_impulsos||N|"
               Top             =   1725
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   3
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   182
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_contador||N|"
               Top             =   1350
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   2
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   181
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_vavulas_3vias||N|"
               Top             =   915
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   1
               Left            =   2010
               MaxLength       =   10
               TabIndex        =   180
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|dn_valvula_esfera||N|"
               Top             =   555
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   0
               Left            =   2010
               MaxLength       =   10
               TabIndex        =   179
               Tag             =   "dn_contador|T|S|||rae_visitas_hidtomas|dn_contador||N|"
               Top             =   210
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   8
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   178
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   3270
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   6
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   177
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   2505
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   9
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   176
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   3705
               Width           =   2865
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   6
               Left            =   2340
               Top             =   4140
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
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   6
               Left            =   9750
               TabIndex        =   187
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   7
               Left            =   5310
               TabIndex        =   188
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label164 
               Caption         =   "Válvula de aislamiento"
               Height          =   315
               Left            =   30
               TabIndex        =   201
               Top             =   2175
               Width           =   1575
            End
            Begin VB.Label Label165 
               Caption         =   "Tipo Arqueta"
               Height          =   315
               Left            =   30
               TabIndex        =   200
               Top             =   2955
               Width           =   1815
            End
            Begin VB.Label Label168 
               Caption         =   "DN tubería"
               Height          =   315
               Left            =   30
               TabIndex        =   199
               Top             =   1785
               Width           =   1815
            End
            Begin VB.Label Label170 
               Caption         =   "DN Ventosa"
               Height          =   315
               Left            =   30
               TabIndex        =   198
               Top             =   1395
               Width           =   1815
            End
            Begin VB.Label Label176 
               Caption         =   "Observaciones"
               Height          =   315
               Left            =   30
               TabIndex        =   197
               Top             =   4170
               Width           =   1635
            End
            Begin VB.Label Label177 
               Caption         =   "Tapa Arqueta"
               Height          =   315
               Left            =   30
               TabIndex        =   196
               Top             =   3360
               Width           =   1635
            End
            Begin VB.Label Label178 
               Caption         =   "Operatividad de la válvula"
               Height          =   315
               Left            =   30
               TabIndex        =   195
               Top             =   2565
               Width           =   2175
            End
            Begin VB.Image Image6 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   10740
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
            Begin VB.Image Image7 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   6210
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
            Begin VB.Label Label24 
               Caption         =   "INTERIOR"
               Height          =   315
               Left            =   9750
               TabIndex        =   194
               Top             =   60
               Width           =   855
            End
            Begin VB.Label Label25 
               Caption         =   "EXTERIOR"
               Height          =   315
               Left            =   5310
               TabIndex        =   193
               Top             =   60
               Width           =   825
            End
            Begin VB.Label Label27 
               Caption         =   "NUDO +"
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
               Height          =   315
               Left            =   30
               TabIndex        =   192
               Top             =   915
               Width           =   1875
            End
            Begin VB.Label Label28 
               Caption         =   "SECTOR"
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
               Height          =   315
               Left            =   30
               TabIndex        =   191
               Top             =   240
               Width           =   1875
            End
            Begin VB.Label Label30 
               Caption         =   "NUDO -"
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
               Height          =   315
               Left            =   30
               TabIndex        =   190
               Top             =   555
               Width           =   1875
            End
            Begin VB.Label Label31 
               Caption         =   "Situación"
               Height          =   315
               Left            =   30
               TabIndex        =   189
               Top             =   3750
               Width           =   1635
            End
         End
         Begin VB.Frame FrameAux5 
            BorderStyle     =   0  'None
            Height          =   4680
            Left            =   -74880
            TabIndex        =   146
            Top             =   420
            Width           =   14280
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   10
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   158
               Tag             =   "x|T|S|||rae_visitas_hidtomas|x||N|"
               Top             =   3720
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   7
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   157
               Tag             =   "x|T|S|||rae_visitas_hidtomas|x||N|"
               Top             =   2670
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   555
               Index           =   11
               Left            =   2010
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   156
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   4080
               Width           =   12135
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   0
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   155
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_solenoide||N|"
               Top             =   210
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   1
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   154
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_contador||N|"
               Top             =   555
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   2
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   153
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_emisor_impulsos||N|"
               Top             =   915
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   3
               Left            =   2010
               MaxLength       =   50
               TabIndex        =   152
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|observaciones||N|"
               Top             =   1260
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   4
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   151
               Tag             =   "x|T|S|||rae_visitas_hidtomas|x||N|"
               Top             =   1620
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   5
               Left            =   2010
               MaxLength       =   250
               ScrollBars      =   2  'Vertical
               TabIndex        =   150
               Tag             =   "the_geom|T|S|||rae_visitas_hidtomas|the_geom||N|"
               Top             =   1965
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   6
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   149
               Tag             =   "Parcelas|T|S|||rae_visitas_hidtomas|parcelas||N|"
               Top             =   2310
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   9
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   148
               Tag             =   "Superficie|N|S|||rae_visitas_hidtomas|superficie||N|"
               Top             =   3375
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   8
               Left            =   2010
               MaxLength       =   255
               TabIndex        =   147
               Tag             =   "Poligono|N|S|||rae_visitas_hidtomas|poligono||N|"
               Top             =   3015
               Width           =   2865
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   5
               Left            =   2610
               Top             =   120
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
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   4
               Left            =   9750
               TabIndex        =   159
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   5
               Left            =   5310
               TabIndex        =   160
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label135 
               Caption         =   "Situacion Arqueta"
               Height          =   315
               Left            =   30
               TabIndex        =   174
               Top             =   3720
               Width           =   1875
            End
            Begin VB.Label Label136 
               Caption         =   "EXTERIOR"
               Height          =   315
               Left            =   5310
               TabIndex        =   173
               Top             =   60
               Width           =   885
            End
            Begin VB.Label Label137 
               Caption         =   "Operatividad de la válvula"
               Height          =   315
               Left            =   30
               TabIndex        =   172
               Top             =   2670
               Width           =   1875
            End
            Begin VB.Label Label140 
               Caption         =   "Observaciones"
               Height          =   315
               Left            =   30
               TabIndex        =   171
               Top             =   4080
               Width           =   1635
            End
            Begin VB.Label Label147 
               Caption         =   "NUDO -"
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
               Height          =   315
               Left            =   30
               TabIndex        =   170
               Top             =   555
               Width           =   1875
            End
            Begin VB.Label Label148 
               Caption         =   "SECTOR"
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
               Height          =   315
               Left            =   30
               TabIndex        =   169
               Top             =   240
               Width           =   1875
            End
            Begin VB.Label Label149 
               Caption         =   "NUDO +"
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
               Height          =   315
               Left            =   30
               TabIndex        =   168
               Top             =   915
               Width           =   1875
            End
            Begin VB.Label Label150 
               Caption         =   "INTERIOR"
               Height          =   315
               Left            =   9750
               TabIndex        =   167
               Top             =   60
               Width           =   945
            End
            Begin VB.Label Label152 
               Caption         =   "Conexiones"
               Height          =   315
               Left            =   30
               TabIndex        =   166
               Top             =   1965
               Width           =   1875
            End
            Begin VB.Label Label153 
               Caption         =   "DN Válvula"
               Height          =   315
               Left            =   30
               TabIndex        =   165
               Top             =   1260
               Width           =   1875
            End
            Begin VB.Label Label154 
               Caption         =   "Tipo de Válvula"
               Height          =   315
               Left            =   30
               TabIndex        =   164
               Top             =   1620
               Width           =   1875
            End
            Begin VB.Label Label158 
               Caption         =   "Tipologia Arqueta"
               Height          =   315
               Left            =   30
               TabIndex        =   163
               Top             =   3015
               Width           =   1875
            End
            Begin VB.Label Label159 
               Caption         =   "Tapa arqueta"
               Height          =   315
               Left            =   30
               TabIndex        =   162
               Top             =   3375
               Width           =   1875
            End
            Begin VB.Label Label160 
               Caption         =   "Posicion del eje (mariposa)"
               Height          =   315
               Left            =   30
               TabIndex        =   161
               Top             =   2310
               Width           =   1875
            End
            Begin VB.Image Image4 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   6210
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
            Begin VB.Image Image5 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   10740
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
         End
         Begin VB.Frame FrameAux4 
            BorderStyle     =   0  'None
            Height          =   4560
            Left            =   120
            TabIndex        =   121
            Top             =   420
            Width           =   14280
            Begin VB.TextBox txtaux5 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   7
               Left            =   1770
               MaxLength       =   250
               TabIndex        =   131
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   2790
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   5
               Left            =   1770
               MaxLength       =   250
               TabIndex        =   130
               Tag             =   "Instaladora|T|S|||rae_visitas_hidtomas|instaladora||N|"
               Top             =   2070
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   4
               Left            =   1770
               MaxLength       =   250
               TabIndex        =   129
               Tag             =   "x|T|S|||rae_visitas_hidtomas|x||N|"
               Top             =   1710
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   3
               Left            =   1770
               MaxLength       =   250
               TabIndex        =   128
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_estanqueidad||N|"
               Top             =   1350
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   6
               Left            =   1770
               MaxLength       =   250
               TabIndex        =   127
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_solenoide||N|"
               Top             =   2400
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   0
               Left            =   1770
               MaxLength       =   10
               TabIndex        =   126
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|dn_toma||N|"
               Top             =   150
               Width           =   1185
            End
            Begin VB.TextBox txtaux5 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   2
               Left            =   1770
               MaxLength       =   10
               TabIndex        =   125
               Tag             =   "dn_contador|T|S|||rae_visitas_hidtomas|dn_contador||N|"
               Top             =   990
               Width           =   1185
            End
            Begin VB.TextBox txtaux5 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   1
               Left            =   1770
               MaxLength       =   40
               TabIndex        =   124
               Tag             =   "Falta Bypass|T|N|||rae_visitas_hidtomas|falta_bypass||S|"
               Top             =   630
               Width           =   1185
            End
            Begin VB.TextBox txtaux5 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   8
               Left            =   1770
               MaxLength       =   250
               TabIndex        =   123
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   3150
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
               BackColor       =   &H80000018&
               Height          =   1005
               Index           =   9
               Left            =   1770
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   122
               Tag             =   "x|T|S|||rae_visitas_hidtomas|x||N|"
               Top             =   3510
               Width           =   3945
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   4
               Left            =   2610
               Top             =   120
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
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   0
               Left            =   7080
               TabIndex        =   132
               Top             =   30
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   1
               Left            =   11340
               TabIndex        =   133
               Top             =   30
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label87 
               Caption         =   "Tch Fijacion Colector"
               Height          =   315
               Left            =   30
               TabIndex        =   145
               Top             =   2430
               Width           =   1635
            End
            Begin VB.Label Label88 
               Caption         =   "Caja Empalmes Correcta"
               Height          =   315
               Left            =   30
               TabIndex        =   144
               Top             =   2820
               Width           =   1755
            End
            Begin VB.Label Label89 
               Caption         =   "Tch Tipo"
               Height          =   315
               Left            =   30
               TabIndex        =   143
               Top             =   2130
               Width           =   1755
            End
            Begin VB.Label Label95 
               Caption         =   "INTERIOR"
               Height          =   315
               Left            =   10380
               TabIndex        =   142
               Top             =   180
               Width           =   855
            End
            Begin VB.Label Label96 
               Caption         =   "EXTERIOR"
               Height          =   315
               Left            =   6120
               TabIndex        =   141
               Top             =   150
               Width           =   1185
            End
            Begin VB.Label Label99 
               Caption         =   "Estado Colector"
               Height          =   315
               Left            =   30
               TabIndex        =   140
               Top             =   1770
               Width           =   1545
            End
            Begin VB.Label Label100 
               Caption         =   "HIDRANTE"
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
               Height          =   255
               Left            =   30
               TabIndex        =   139
               Top             =   240
               Width           =   1035
            End
            Begin VB.Label Label101 
               Caption         =   "Válvula Compuerta"
               Height          =   255
               Left            =   30
               TabIndex        =   138
               Top             =   1380
               Width           =   1545
            End
            Begin VB.Label Label102 
               Caption         =   "Fecha 1ª revisión"
               Height          =   255
               Left            =   30
               TabIndex        =   137
               Top             =   990
               Width           =   1455
            End
            Begin VB.Label Label103 
               Caption         =   "Constructora"
               Height          =   255
               Left            =   30
               TabIndex        =   136
               Top             =   630
               Width           =   1035
            End
            Begin VB.Label Label105 
               Caption         =   "Nivelacion Arquetar"
               Height          =   315
               Left            =   30
               TabIndex        =   135
               Top             =   3210
               Width           =   1815
            End
            Begin VB.Label Label119 
               Caption         =   "Observaciones"
               Height          =   315
               Left            =   30
               TabIndex        =   134
               Top             =   3540
               Width           =   1575
            End
            Begin VB.Image Image2 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   6150
               Stretch         =   -1  'True
               Top             =   480
               Width           =   3405
            End
            Begin VB.Image Image3 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   10350
               Stretch         =   -1  'True
               Top             =   480
               Width           =   3405
            End
         End
         Begin VB.Frame FrameAux2 
            BorderStyle     =   0  'None
            Height          =   4410
            Left            =   -74880
            TabIndex        =   71
            Top             =   420
            Width           =   14370
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   2
               Left            =   1410
               MaxLength       =   255
               TabIndex        =   120
               Top             =   1230
               Width           =   3375
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   3
               Left            =   1410
               MaxLength       =   5
               TabIndex        =   117
               Tag             =   "Poligono|N|S|||rae_visitas_hidtomas|poligono||N|"
               Top             =   1680
               Width           =   3375
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   11
               Left            =   6900
               MaxLength       =   20
               TabIndex        =   116
               Tag             =   "Fecha Entrada|F|S|||rae_visitas_hidtomas|fecha_entrada||N|"
               Top             =   1313
               Width           =   1605
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   9
               Left            =   6900
               MaxLength       =   10
               TabIndex        =   113
               Tag             =   "Toma|T|S|||rae_visitas_hidtomas|toma||N|"
               Top             =   531
               Width           =   1575
            End
            Begin MSComctlLib.Toolbar Toolbar2 
               Height          =   330
               Left            =   13470
               TabIndex        =   104
               Top             =   30
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Sigpac"
                  EndProperty
               EndProperty
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   0
               Left            =   1410
               MaxLength       =   40
               TabIndex        =   87
               Tag             =   "Contador|T|N|||rae_visitas_hidtomas|contador||S|"
               Top             =   270
               Width           =   1245
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   6
               Left            =   1410
               MaxLength       =   10
               TabIndex        =   86
               Tag             =   "dn_contador|T|S|||rae_visitas_hidtomas|dn_contador||N|"
               Top             =   2880
               Width           =   3345
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   7
               Left            =   1410
               MaxLength       =   125
               TabIndex        =   85
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|dn_valvula_esfera||N|"
               Top             =   3270
               Width           =   3315
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   8
               Left            =   1410
               MaxLength       =   125
               TabIndex        =   84
               Tag             =   "dn_toma|T|S|||rae_visitas_hidtomas|dn_toma||N|"
               Top             =   3660
               Width           =   3285
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   4
               Left            =   1410
               MaxLength       =   125
               TabIndex        =   83
               Tag             =   "Parcelas|T|S|||rae_visitas_hidtomas|parcelas||N|"
               Top             =   2070
               Width           =   3375
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   12
               Left            =   6900
               MaxLength       =   20
               TabIndex        =   82
               Tag             =   "Instaladora|T|S|||rae_visitas_hidtomas|instaladora||N|"
               Top             =   1704
               Width           =   1605
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   5
               Left            =   1410
               MaxLength       =   20
               TabIndex        =   81
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   2430
               Width           =   1275
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   10
               Left            =   6900
               MaxLength       =   20
               TabIndex        =   80
               Tag             =   "Recibido|T|S|||rae_visitas_hidtomas|Recibido||N|"
               Top             =   922
               Width           =   1605
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   13
               Left            =   6900
               MaxLength       =   20
               TabIndex        =   79
               Tag             =   "Fecha Revision|F|S|||rae_visitas_hidtomas|fecha_revision||N|"
               Top             =   2095
               Width           =   1605
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   1275
               Index           =   18
               Left            =   9540
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   78
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|observaciones_RAE||N|"
               Top             =   510
               Width           =   4665
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   14
               Left            =   6900
               MaxLength       =   255
               TabIndex        =   77
               Tag             =   "Fecha Turno|T|S|||rae_visitas_hidtomas|fecha_turno||N|"
               Top             =   2486
               Width           =   1635
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   17
               Left            =   6900
               MaxLength       =   255
               TabIndex        =   76
               Tag             =   "Verificacion|T|S|||rae_visitas_hidtomas|verificacion||N|"
               Top             =   3660
               Width           =   1635
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   1275
               Index           =   19
               Left            =   9570
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   75
               Tag             =   "Comentarios Indefa|T|S|||rae_visitas_hidtomas|Comentarios_INDEFA||N|"
               Top             =   2610
               Width           =   4665
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   15
               Left            =   6900
               MaxLength       =   255
               TabIndex        =   74
               Tag             =   "Turno|T|S|||rae_visitas_hidtomas|turno||N|"
               Top             =   2877
               Width           =   1635
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   16
               Left            =   6900
               MaxLength       =   255
               TabIndex        =   73
               Tag             =   "q_instantaneo|T|S|||rae_visitas_hidtomas|q_instantaneo||N|"
               Top             =   3268
               Width           =   1635
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   290
               Index           =   1
               Left            =   1410
               MaxLength       =   255
               TabIndex        =   72
               Tag             =   "Nro socio|N|S|||rae_visitas_hidtomas|socio_revisado||N|"
               Top             =   825
               Width           =   1245
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   2
               Left            =   6360
               Top             =   -90
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
            Begin VB.Label Label65 
               Caption         =   "Polígono"
               Height          =   285
               Left            =   30
               TabIndex        =   119
               Top             =   1665
               Width           =   1815
            End
            Begin VB.Label Label40 
               Caption         =   "F.entrada ficha"
               Height          =   290
               Left            =   4980
               TabIndex        =   118
               Top             =   1308
               Width           =   1125
            End
            Begin VB.Label Label186 
               Caption         =   "DN Contador"
               Height          =   290
               Left            =   30
               TabIndex        =   115
               Top             =   2872
               Width           =   1035
            End
            Begin VB.Label Label185 
               Caption         =   "Toma"
               Height          =   285
               Left            =   4980
               TabIndex        =   114
               Top             =   536
               Width           =   1035
            End
            Begin VB.Label Label17 
               Caption         =   "CONTADOR"
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
               Height          =   255
               Left            =   30
               TabIndex        =   103
               Top             =   240
               Width           =   1035
            End
            Begin VB.Label Label19 
               Caption         =   "Nombre Socio"
               Height          =   285
               Left            =   30
               TabIndex        =   102
               Top             =   1215
               Width           =   1035
            End
            Begin VB.Label Label20 
               Caption         =   "DN Valvula"
               Height          =   290
               Left            =   30
               TabIndex        =   101
               Top             =   3263
               Width           =   1035
            End
            Begin VB.Label Label21 
               Caption         =   "DN Toma"
               Height          =   290
               Left            =   30
               TabIndex        =   100
               Top             =   3660
               Width           =   1035
            End
            Begin VB.Label Label22 
               Caption         =   "Nº Socio"
               Height          =   285
               Left            =   30
               TabIndex        =   99
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label Label35 
               Caption         =   "Parcela/s"
               Height          =   285
               Left            =   30
               TabIndex        =   98
               Top             =   2055
               Width           =   705
            End
            Begin VB.Label Label36 
               Caption         =   "Instaladora"
               Height          =   290
               Left            =   4980
               TabIndex        =   97
               Top             =   1699
               Width           =   795
            End
            Begin VB.Label Label38 
               Caption         =   "Superficie total(hg)"
               Height          =   285
               Left            =   30
               TabIndex        =   96
               Top             =   2445
               Width           =   1455
            End
            Begin VB.Label Label42 
               Caption         =   "Alta Recibida"
               Height          =   285
               Left            =   5010
               TabIndex        =   95
               Top             =   922
               Width           =   1125
            End
            Begin VB.Label Label43 
               Caption         =   "Fecha revisión instalación"
               Height          =   290
               Left            =   4980
               TabIndex        =   94
               Top             =   2090
               Width           =   1845
            End
            Begin VB.Label Label45 
               Caption         =   "Acciones Requeridas RAE"
               Height          =   315
               Left            =   9540
               TabIndex        =   93
               Top             =   210
               Width           =   5115
            End
            Begin VB.Label Label52 
               Caption         =   "Observaciones INDEFA"
               Height          =   315
               Left            =   9540
               TabIndex        =   92
               Top             =   2220
               Width           =   1815
            End
            Begin VB.Label Label54 
               Caption         =   "Turno Asignado"
               Height          =   290
               Left            =   4980
               TabIndex        =   91
               Top             =   2872
               Width           =   1815
            End
            Begin VB.Label Label57 
               Caption         =   "Caudal instantáneo"
               Height          =   290
               Left            =   4980
               TabIndex        =   90
               Top             =   3263
               Width           =   1815
            End
            Begin VB.Label Label63 
               Caption         =   "Fecha puesta en turno"
               Height          =   290
               Left            =   4980
               TabIndex        =   89
               Top             =   2481
               Width           =   2535
            End
            Begin VB.Label Label64 
               Caption         =   "Tipo verificación (C/T)"
               Height          =   290
               Left            =   4980
               TabIndex        =   88
               Top             =   3660
               Width           =   1815
            End
         End
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   4350
         Left            =   -74910
         TabIndex        =   40
         Top             =   450
         Width           =   6780
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   5
            Left            =   6570
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   51
            Text            =   "Par"
            Top             =   2940
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   4
            Left            =   6180
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   50
            Text            =   "Pol"
            Top             =   2940
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   5700
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   49
            Text            =   "Hdas"
            Top             =   2940
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   4350
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   48
            Text            =   "Poblacion"
            Top             =   2940
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   2565
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   45
            Text            =   "Partida"
            Top             =   2925
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   1
            Left            =   2385
            TabIndex        =   44
            ToolTipText     =   "Buscar campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   1710
            MaxLength       =   8
            TabIndex        =   43
            Tag             =   "Campo|N|N|||rpozos_campos|codcampo|00000000|N|"
            Text            =   "campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   225
            MaxLength       =   10
            TabIndex        =   42
            Tag             =   "Hidrante|T|N|||rpozos_campos|hidrante||S|"
            Text            =   "Hidrante"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   945
            MaxLength       =   6
            TabIndex        =   41
            Tag             =   "Linea|N|N|||rpozos_campos|numlinea|000|N|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   45
            TabIndex        =   46
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   688
            ButtonWidth     =   609
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
            Index           =   1
            Left            =   4590
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
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmPOZHidrantesIndefa.frx":0318
            Height          =   3810
            Index           =   1
            Left            =   30
            TabIndex        =   47
            Top             =   480
            Width           =   6660
            _ExtentX        =   11748
            _ExtentY        =   6720
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
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   3930
         Left            =   -74910
         TabIndex        =   30
         Top             =   450
         Width           =   6780
         Begin VB.TextBox txtaux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   5940
            MaxLength       =   6
            TabIndex        =   36
            Tag             =   "Porcentaje|N|N|0|100|rpozos_cooprop|porcentaje|##0.00||"
            Text            =   "porc"
            Top             =   2940
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   945
            MaxLength       =   6
            TabIndex        =   35
            Tag             =   "Linea|N|N|||rpozos_cooprop|numlinea|000|S|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   225
            MaxLength       =   10
            TabIndex        =   34
            Tag             =   "Hidrante|T|N|||rpozos_cooprop|hidrante||S|"
            Text            =   "Hidrante"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   1710
            MaxLength       =   6
            TabIndex        =   33
            Tag             =   "Socio|N|N|||rpozos_cooprop|codsocio|000000|N|"
            Text            =   "socio"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   0
            Left            =   2385
            TabIndex        =   32
            ToolTipText     =   "Buscar socio"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   2565
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   31
            Text            =   "Nombre socio"
            Top             =   2925
            Visible         =   0   'False
            Width           =   3285
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   45
            TabIndex        =   37
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   688
            ButtonWidth     =   609
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
            Left            =   4590
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
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmPOZHidrantesIndefa.frx":0330
            Height          =   3195
            Index           =   0
            Left            =   30
            TabIndex        =   38
            Top             =   450
            Width           =   6450
            _ExtentX        =   11377
            _ExtentY        =   5636
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
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   780
      Top             =   6300
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
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1710
      MaxLength       =   40
      TabIndex        =   112
      Top             =   990
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   7170
      Top             =   4770
      Width           =   1245
   End
   Begin VB.Label Label50 
      Caption         =   "Buscando diferencias"
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   3510
      TabIndex        =   69
      Top             =   9030
      Visible         =   0   'False
      Width           =   4215
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
      Begin VB.Menu mnDiferencias 
         Caption         =   "Buscar &Diferencias"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnActualizar 
         Caption         =   "Actualizar Registro"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
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
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmPOZHidrantesIndefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA                   -+-+
' +-+- Menú: Hidrantes de Pozos        -+-+
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

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmPar As frmManPartidas 'partidas
Attribute frmPar.VB_VarHelpID = -1
Private WithEvents frmPoz As frmPOZPozos 'tipos de Pozos
Attribute frmPoz.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmSoc1 As frmManSocios 'socios
Attribute frmSoc1.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCampos 'campos
Attribute frmCam.VB_VarHelpID = -1
Private WithEvents frmCam1 As frmManCampos 'campos
Attribute frmCam1.VB_VarHelpID = -1
Private WithEvents frmMen2 As frmMensajes ' orden de printnou
Attribute frmMen2.VB_VarHelpID = -1
Private WithEvents frmMen3 As frmMensajes ' busqueda de diferencias
Attribute frmMen3.VB_VarHelpID = -1
Private frmMensImg As frmMensajes ' visualizacion de la imagen

' *****************************************************
Dim CodTipoMov As String

Dim Orden As String

Dim ConexionIndefa As Boolean
Dim Continuar As Boolean

Dim SocioAnt As String

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

Dim ListOrigen As Collection
Dim ListDestino As Collection



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

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim vSeccion As CSeccion
Dim b As Boolean

Private BuscaChekc As String
Private NumCajas As Currency
Private NumCajasAnt As Currency
Private NumKilosAnt As Currency

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadparam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Public ImpresoraDefecto As String

Dim Lineas As Collection
Dim NF As Integer


Dim MostradoAviso As Boolean

Private Sub cmdAceptar_Click()
Dim Diferencias As String
Dim i As Integer

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 2, "Frame2") Then
                
                    '------------------------------------------------------------------------------
                    '  LOG de acciones
                    Set LOG = New cLOG
                    LOG.Insertar 10, vUsu, "Nuevo contador: " & Text1(0).Text
                    Set LOG = Nothing
                    '-----------------------------------------------------------------------------
                
                    ImprimirComunicacionIndefa True

                    ' *** canviar o llevar el WHERE, repasar codEmpre ****
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    'Data1.RecordSource = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
                    ' ***************************************************************
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 2, "Frame2") Then
                
                    Diferencias = ""
                    For i = 0 To Text1.Count - 1
                        If Text1(i).Text <> ListOrigen.item(i + 1) Then
                            Diferencias = Diferencias & Mid(Text1(i).Tag, 1, 8) & ":" & ListOrigen.item(i + 1) & "-" & Text1(i).Text & "·"
                        End If
                    Next i
                    Set ListOrigen = Nothing
                    
                    '------------------------------------------------------------------------------
                    '  LOG de acciones
                    Set LOG = New cLOG
                    LOG.Insertar 11, vUsu, "Contador: " & Text1(0).Text & " " & Diferencias
                    Set LOG = Nothing
                    '-----------------------------------------------------------------------------
                    
                    ImprimirComunicacionIndefa False
                
                
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    ModificarLinea
                    PosicionarData
                    SumaTotalPorcentajes NumTabMto
            End Select
        
    End Select
    Screen.MousePointer = vbDefault
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub ImprimirComunicacionIndefa(EsAlta As Boolean)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Contador As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

    If Text1(0).Text = "" Or Len(Text1(0).Text) <> 6 Then Exit Sub
    
    Contador = Text1(0).Text
    
    If EsAlta Then
        If MsgBox(" Se ha dado de alta un nuevo contador." & vbCrLf & vbCrLf & "¿ Desea imprimir un informe de comunicación a Indefa ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            
            indRPT = 88 ' informe de comunicacion de cambios a indefa
            
            If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
                       
            InicializarVbles
            
            'Añadir el parametro de Empresa
            cadparam = cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
            cadparam = cadparam & "pContador=""" & Text1(0).Text & """|"
            numParam = numParam + 1
            cadparam = cadparam & "pHanegadas=" & DBSet(Text1(6).Text, "N") & "|"
            numParam = numParam + 1
            cadparam = cadparam & "pPoligono=" & DBSet(Text1(4).Text, "T") & "|"
            numParam = numParam + 1
            cadparam = cadparam & "pParcela=" & DBSet(Text1(5).Text, "T") & "|"
            numParam = numParam + 1
            cadparam = cadparam & "pSocio=" & Text1(2).Text & "|"
            numParam = numParam + 1
            cadparam = cadparam & "pToma=" & CLng(ComprobarCero(Text1(1).Text)) Mod 100 & "|"
            numParam = numParam + 1
            

            cadTitulo = "Carta de Comunicación a Indefa"
            cadNombreRPT = nomDocu
            cadFormula = "{rpozos.hidrante} = " & DBSet(Contador, "T")
            
            LlamarImprimir
        
        End If
        
        Exit Sub
    End If
    
    
    
    If Not ConexionIndefa Then Exit Sub
    
    
    Sql = "select poligono, parcelas, hanegadas, toma, socio_revisado from rae_visitas_hidtomas where sector = " & DBSet(CInt(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante = " & DBSet(CInt(Mid(Contador, 3, 2)), "T")
    '[Monica]18/07/2013:cambio
                                    '[Monica]27/01/2014: lo cambio a numerico
    Sql = Sql & " and salida_tch = " & DBSet(CInt(Mid(Contador, 5, 2)), "N")

    Set RS = New ADODB.Recordset
    RS.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
      '[Monica]19/11/2013: si no hay socio revisado no hacemos nada
      If DBLet(RS!socio_revisado, "N") <> 0 Then
        
        If (DBLet(RS!Poligono, "T") <> Text1(4).Text) Or (Mid(DBLet(RS!parcelas, "T"), 1, 25) <> Mid(Text1(5).Text, 1, 25)) Or (Int(ComprobarCero(DBLet(RS!Hanegadas, "N"))) <> Int(ComprobarCero(Text1(6).Text))) Or CInt(DBLet(RS!socio_revisado, "N") <> Text1(2).Text And DBLet(RS!socio_revisado, "N") <> 0) Or _
           (CLng(ComprobarCero(DBLet(RS!toma, "N"))) <> CLng(ComprobarCero(Text1(1).Text)) Mod 100) Then
            If MsgBox("Se han producido diferencias con los datos de Indefa." & vbCrLf & vbCrLf & " ¿ Desea imprimir un informe de comunicación ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                
                indRPT = 88 ' informe de comunicacion de cambios a indefa
                
                If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
                           
                InicializarVbles
                
                'Añadir el parametro de Empresa
                cadparam = cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
                numParam = numParam + 1
                cadparam = cadparam & "pIndPol=""" & DBLet(RS!Poligono, "T") & """|"
                numParam = numParam + 1
                cadparam = cadparam & "pIndPar=""" & DBLet(RS!parcelas, "T") & """|"
                numParam = numParam + 1
                cadparam = cadparam & "pIndHda=" & DBSet(RS!Hanegadas, "N") & "|"
                numParam = numParam + 1
                cadparam = cadparam & "pIndToma=" & CLng(ComprobarCero(DBSet(RS!toma, "N"))) & "|"
                numParam = numParam + 1
                cadparam = cadparam & "pContador=""" & Text1(0).Text & """|"
                numParam = numParam + 1
                cadparam = cadparam & "pSocioAnt=" & SocioAnt & "|"
                numParam = numParam + 1
                cadparam = cadparam & "pHanegadas=" & DBSet(Text1(6).Text, "N") & "|"
                numParam = numParam + 1
                cadparam = cadparam & "pPoligono=" & DBSet(Text1(4).Text, "T") & "|"
                numParam = numParam + 1
                cadparam = cadparam & "pParcela=" & DBSet(Text1(5).Text, "T") & "|"
                numParam = numParam + 1
                cadparam = cadparam & "pSocio=" & Text1(2).Text & "|"
                numParam = numParam + 1
                cadparam = cadparam & "pToma=" & CLng(ComprobarCero(Text1(1).Text)) Mod 100 & "|"
                numParam = numParam + 1
                

                cadTitulo = "Carta de Comunicación a Indefa"
                cadNombreRPT = nomDocu
                cadFormula = "{rpozos.hidrante} = " & DBSet(Contador, "T")
                
                LlamarImprimir
                
            End If
        End If
      End If
    End If
        
End Sub


Private Sub cmdAux_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 ' Socios coopropietarios
            Set frmSoc1 = New frmManSocios
            frmSoc1.DatosADevolverBusqueda = "0|1|"
            frmSoc1.Show vbModal
            Set frmSoc1 = Nothing
            PonerFoco txtaux3(2)
            
        Case 1 ' campos
            Set frmCam1 = New frmManCampos
            frmCam1.DatosADevolverBusqueda = "0|"
'            frmCam.CodigoActual = Text1(5).Text
            frmCam1.Show vbModal
            Set frmCam1 = Nothing
            PonerFoco txtaux4(2)
        
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub


Private Sub Form_Activate()
Dim RS As ADODB.Recordset
Dim Sql As String

    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
        DoEvents
        If Not Continuar Then Unload Me
        
        If Not MostradoAviso Then
            If ConexionIndefa Then
'                SQL = "select * from rae_visitas_hidtomas "
'
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Not RS.EOF Then
'                    If RS.Fields.Count > 52 Then
'                        If MsgBox("Han cambiado la estructura de Contadores Indefa, hay datos que no se van a mostrar." & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'                            Set RS = Nothing
'                            Exit Sub
'                        End If
'                    End If
'                End If
'
'                SQL = "select * from rae_visitas_hidrantes "
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Not RS.EOF Then
'                    If RS.Fields.Count > 38 Then
'                        If MsgBox("Han cambiado la estructura de Hidrantes Indefa, hay datos que no se van a mostrar." & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'                            Set RS = Nothing
'                            Exit Sub
'                        End If
'                    End If
'                End If
'
'                SQL = "select * from rae_visitas_desagues "
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Not RS.EOF Then
'                    If RS.Fields.Count > 20 Then
'                        If MsgBox("Han cambiado la estructura de Desagües Indefa, hay datos que no se van a mostrar." & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'                            Set RS = Nothing
'                            Exit Sub
'                        End If
'                    End If
'                End If
'
            End If
        End If
        
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
    CerrarConexionIndefa
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    Continuar = True
    
    '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
    Me.Caption = "Contadores"
    Me.Label1(0).Caption = "Contador TCH"
    
    ConexionIndefa = False
    If AbrirConexionIndefa() = False Then
        If MsgBox("No se ha podido acceder a los datos de Indefa. " & vbCrLf & "¿ Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Continuar = False
            Exit Sub
        End If
    Else
        ConexionIndefa = True
    End If
    
    ' ICONETS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Tots
        .Buttons(5).Image = 21  'Buscar diferencias
        .Buttons(6).Image = 26  'Actualizar desde datos de indefa
        'el 5 i el 6 son separadors
        .Buttons(8).Image = 3   'Insertar
        .Buttons(9).Image = 4   'Modificar
        .Buttons(10).Image = 5   'Borrar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Eixir
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
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i


'    Me.cmdSigpac.Picture = frmPpal.imgListComun16.ListImages(29).Picture
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 29   'Insertar
    End With
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    '[Monica]08/02/2013: cargamos todos los toolbar de camara de fotos
    For i = 0 To 7
        With Me.Toolbar3(i)
            .HotImageList = frmPpal.imgListComun_OM
            .DisabledImageList = frmPpal.imgListComun_BN
            .ImageList = frmPpal.imgListComun
            .Buttons(1).Image = 36   'camara
        End With
    Next i
    
'    With Me.Toolbar4
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 36   'camara
'    End With
    
    LimpiarCampos   'Neteja els camps TextBox
    
    CodTipoMov = "NOC"

    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rpozos"
    Ordenacion = " ORDER BY hidrante "
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where hidrante is null"
    Data1.Refresh
       
    ModoLineas = 0
    
    Me.SSTab2.Tab = 0
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        Text1(0).BackColor = vbYellow 'codclien
        ' ****************************************************************************
    End If
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
'    Me.chkAbonos(0).Value = 0
    
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************

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
Dim i As Integer, Numreg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    BuscaChekc = ""
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    
' cambio la siguiente expresion por la de abajo
'   BloquearText1 Me, Modo
    For i = 0 To Text1.Count - 1
        BloquearTxt Text1(i), Not (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next i
    BloquearCombo Me, Modo
    
    b = (Modo <> 1)
    BloquearTxt Text1(19), b
    
    FrameAux2.Enabled = (Modo = 2)
    FrameAux3.Enabled = (Modo = 2)
    FrameAux4.Enabled = (Modo = 2)
    FrameAux5.Enabled = (Modo = 2)
    FrameAux6.Enabled = (Modo = 2)
    For i = 0 To txtaux1.Count - 1
        txtaux1(i).Locked = True
    Next i
    For i = 0 To txtaux5.Count - 1
        txtaux5(i).Locked = True
    Next i
    For i = 0 To txtaux6.Count - 1
        txtaux6(i).Locked = True
    Next i
    For i = 0 To txtaux7.Count - 1
        txtaux7(i).Locked = True
    Next i
    'Campos Nº entrada bloqueado y en azul
'    BloquearTxt Text1(0), Modo = 4
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    For i = 0 To imgFec.Count - 1
        BloquearImgFec Me, i, Modo
    Next i
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    ' *** si n'hi han llínies i imagens de buscar que no estiguen als grids ******
    'Llínies Departaments
'    b = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
    ' ****************************************************************************
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    

    DataGridAux(0).Enabled = b
    
    Me.Toolbar2.Enabled = (Modo = 2)
    Me.Toolbar2.visible = (Modo = 2)
    ' las camaras
    For i = 0 To 7
        Me.Toolbar3(i).Enabled = (Modo = 2)
        Me.Toolbar3(i).visible = (Modo = 2)
    Next i
        
'        Me.Toolbar3.Enabled = (Modo = 2)
'        Me.Toolbar3.visible = (Modo = 2)
'        Me.Toolbar4.Enabled = (Modo = 2)
'        Me.Toolbar4.visible = (Modo = 2)

'     '-----------------------------
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
    'Buscar diferencias con indefa
    Toolbar1.Buttons(5).Enabled = b And ConexionIndefa
    Me.mnDiferencias.Enabled = b And ConexionIndefa
    
    'Insertar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
    
    'actualizar registro con datos de indefa         '[Monica]27/01/2014: solo para el caso de escalona
    Toolbar1.Buttons(6).Enabled = b And ConexionIndefa And (vParamAplic.Cooperativa = 10)
    Me.mnActualizar.Enabled = b And ConexionIndefa And (vParamAplic.Cooperativa = 10)
    
    'Modificar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(10).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(12).Enabled = b
       
       
       
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    ' ****************************************
       
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub






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
        ' *** canviar o llevar el WHERE; repasar codEmpre ***
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        'CadenaConsulta = "select * from " & NombreTabla & " WHERE codempre = " & codEmpre & " AND " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
Dim indice As Byte
'    indice = CByte(Me.cmdAux(0).Tag + 2)
'    txtaux1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
    Text1(18).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de campo
    FormateaCampo Text1(18)
    PonerDatosCampo Text1(18).Text
End Sub

Private Sub frmCam1_DatoSeleccionado(CadenaSeleccion As String)
    txtaux4(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codcampo
    FormateaCampo txtaux4(2)
    PonerDatosCampoLineas txtaux4(2)

End Sub


Private Sub frmMen2_DatoSeleccionado(CadenaSeleccion As String)
    Orden = CadenaSeleccion
    If CadenaSeleccion = "" Then Orden = "pOrden={rpozos.hidrante}"
End Sub

Private Sub frmMen3_DatoSeleccionado(CadenaSeleccion As String)
    CadB = ""
    If CadenaSeleccion <> "" Then
        CadB = "hidrante in (" & Mid(CadenaSeleccion, 1, Len(CadenaSeleccion) - 1) & ")"
        HacerBusquedaDiferencias
    End If
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    Text1(18).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de campo
    FormateaCampo Text1(18)
    PonerDatosCampo Text1(18).Text
End Sub

Private Sub frmPar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de partida
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de partida
End Sub

Private Sub frmPoz_DatoSeleccionado(CadenaSeleccion As String)
    Text1(13).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de pozo
    FormateaCampo Text1(13)
    Text2(13).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de pozo
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo Text1(2)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub frmSoc1_DatoSeleccionado(CadenaSeleccion As String)
    txtaux3(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo txtaux3(2)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub


Private Sub imgFec_Click(Index As Integer)
       
       Screen.MousePointer = vbHourglass
       
       Dim esq As Long
       Dim dalt As Long
       Dim menu As Long
       Dim obj As Object
    
       Set frmC1 = New frmCal
        
       esq = imgFec(Index).Left
       dalt = imgFec(Index).Top
        
       Set obj = imgFec(Index).Container
    
       While imgFec(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
       Wend
        
       menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
    
       frmC1.Left = esq + imgFec(Index).Parent.Left + 30
       frmC1.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
       
       frmC1.NovaData = Now
       Select Case Index
            Case 0
                indice = 8
            Case 1
                indice = 10
            Case 2
                indice = 16
            Case 3
                indice = 17
       End Select
       
       Me.imgFec(0).Tag = indice
       
       PonerFormatoFecha Text1(indice)
       If Text1(indice).Text <> "" Then frmC1.NovaData = CDate(Text1(indice).Text)
    
       Screen.MousePointer = vbDefault
       frmC1.Show vbModal
       Set frmC1 = Nothing
       PonerFoco Text1(indice)
    
End Sub

Private Sub imgZoom_Click(Index As Integer)
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            indice = 11
            frmZ.pTitulo = "Observaciones del Hidrante"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(indice)
    End Select
End Sub

Private Sub mnActualizar_Click()
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    If BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2") Then BotonActualizar
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnDiferencias_Click()
    BotonBuscarDiferencias
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
Dim NroCopias As String
Dim Lin As String

    printNou
    
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2") Then BotonModificar
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


Private Sub SSTab2_Click(PreviousTab As Integer)
'    PonerCampos
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 3  'Búscar
           mnBuscar_Click
        Case 4  'Tots
            mnVerTodos_Click
        Case 5 ' buscar diferencias
            mnDiferencias_Click
        Case 6 ' Actualizar registro con datos de indefa
            mnActualizar_Click
        Case 8  'Nou
            mnNuevo_Click
        Case 9  'Modificar
            mnModificar_Click
        Case 10  'Borrar
            mnEliminar_Click
        Case 12 'Imprimir
'            AbrirListado (10)
            mnImprimir_Click
        Case 13    'Eixir
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

Private Sub BotonActualizar()
Dim i As Integer
Dim Sql As String
Dim Contador As String
Dim RS As ADODB.Recordset
Dim CADENA As String

Dim Orden As Long

    On Error GoTo eBotonActualizar

    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Then Exit Sub
    
    Sql = "select poligono, parcelas, hanegadas, socio_revisado, toma from rae_visitas_hidtomas where sector = " & DBSet(CInt(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante = " & DBSet(CInt(Mid(Contador, 3, 2)), "T")
    '[Monica]18/07/2013:cambio
                                    '[Monica]27/01/2014: lo cambio a numerico
    Sql = Sql & " and salida_tch = " & DBSet(CInt(Mid(Contador, 5, 2)), "N")

    Set RS = New ADODB.Recordset
    RS.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
      '[Monica]19/11/2013: si el socio_revisado esta vacio no hacemos ninguna comprobacion
      If DBLet(RS!socio_revisado, "N") <> 0 Then
      
        '[Monica]07/11/2013: puede que indefa nos haya metido un socio que no existe
        If CInt(DBLet(RS!socio_revisado, "N") <> Text1(2).Text And DBLet(RS!socio_revisado, "N") <> 0) Then
            Sql = "select * from rsocios where codsocio = " & DBSet(RS!socio_revisado, "N")
            If TotalRegistros(Sql) = 0 Then
                MsgBox "Debe dar de alta el socio " & DBLet(RS!socio_revisado, "N"), vbExclamation
                Exit Sub
            End If
        End If
    
        If (DBLet(RS!Poligono, "T") <> Text1(4).Text) Or (Mid(DBLet(RS!parcelas, "T"), 1, 25) <> Mid(Text1(5).Text, 1, 25)) Or (Int(ComprobarCero(DBLet(RS!Hanegadas, "N"))) <> Int(ComprobarCero(Text1(6).Text))) Or CInt(DBLet(RS!socio_revisado, "N") <> Text1(2).Text And DBLet(RS!socio_revisado, "N") <> 0) Or _
           (CLng(ComprobarCero(DBLet(RS!toma, "N"))) <> CLng(ComprobarCero(Text1(1).Text)) Mod 100) Then
            
            CADENA = ""
            If (DBLet(RS!Poligono, "T") <> Text1(4).Text) Then
                CADENA = CADENA & " Pol:" & Trim(Text1(4).Text) & "-" & DBLet(RS!Poligono, "T") & "·"
            End If
            If (Mid(DBLet(RS!parcelas, "T"), 1, 25) <> Mid(Text1(5).Text, 1, 25)) Then
                CADENA = CADENA & "Par:" & Trim(Text1(5).Text) & "-" & (Mid(DBLet(RS!parcelas, "T"), 1, 25)) & "·"
            End If
            If (Int(ComprobarCero(DBLet(RS!Hanegadas, "N"))) <> Int(ComprobarCero(Text1(6).Text))) Then
                CADENA = CADENA & "Hdas:" & Int(ComprobarCero(Text1(6).Text)) & "-" & Int(ComprobarCero(DBLet(RS!Hanegadas, "N"))) & "·"
            End If
            If CLng((DBLet(RS!socio_revisado, "N")) <> CLng(Text1(2).Text) And DBLet(RS!socio_revisado, "N") <> 0) Then
                CADENA = CADENA & "Soc:" & Trim(Text1(2).Text) & "-" & DBLet(RS!socio_revisado, "N") & "·"
            End If
            If (CLng(ComprobarCero(DBLet(RS!toma, "N"))) <> CLng(ComprobarCero(Text1(1).Text)) Mod 100) Then
                CADENA = CADENA & "Toma:" & CLng(Text1(1).Text) Mod 100 & "-" & CLng(ComprobarCero(DBLet(RS!toma, "N"))) & "·"
            End If
            

            '------------------------------------------------------------------------------
            '  LOG de acciones
            Set LOG = New cLOG
            LOG.Insertar 11, vUsu, "Contador:" & Contador & vbCrLf & " " & CADENA
            Set LOG = Nothing
            '-----------------------------------------------------------------------------
            
            Sql = "update rpozos set poligono = " & DBSet(RS!Poligono, "T")
            Sql = Sql & ", parcelas = " & DBSet(Mid(RS!parcelas, 1, 25), "T")
            Sql = Sql & ", hanegada = " & DBSet(RS!Hanegadas, "N")
            '[Monica]23/10/2013: daba error cuando no me han insertado el socio
            If DBLet(RS!socio_revisado, "N") <> 0 Then
                Sql = Sql & ", codsocio = " & DBSet(RS!socio_revisado, "N")
            End If
            '[Monica]30/10/2013: hemos de actualizar tambien el nro de orden con la toma de indefa
            Orden = (CLng(Text1(1).Text) \ 100) * 100 + CLng(ComprobarCero(DBLet(RS!toma, "N")))
            Sql = Sql & ", nroorden = " & DBSet(Orden, "N")
            
            Sql = Sql & " where hidrante = " & DBSet(Contador, "T")
            
            conn.Execute Sql
            TerminaBloquear
            Data1.Refresh
    '        PonerCampos
            SituarData Data1, "hidrante = " & DBSet(Contador, "T"), Me.lblIndicador, False
            PonerCampos
        
            MsgBox "Proceso realizado correctamente.", vbExclamation
            
        End If
      End If
    End If
    Exit Sub
    
eBotonActualizar:
    MuestraError Err.Number, "Actualizar Datos Indefa", Err.Description
End Sub

Private Sub BotonBuscarDiferencias()
Dim i As Integer

    LimpiarCampos
    
    
    Set frmMen3 = New frmMensajes
    frmMen3.OpcionMensaje = 44
    frmMen3.Show vbModal
    
    Set frmMen3 = Nothing
    
    
' ******************************************************************************
End Sub

Private Sub HacerBusquedaDiferencias()
    
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

Private Sub HacerBusqueda()

    '[Monica]09/01/2014: nuevo tipo para que no lleve los asteriscos implicitos
    Text1(4).Tag = "Polígono|TT|S|||rpozos|poligono||N|"
    
    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    Text1(4).Tag = "Polígono|T|S|||rpozos|poligono||N|"
    
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
    Dim cad As String
    Dim NombreTabla1 As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & "Hidrante|rpozos.hidrante|N||18·"
    cad = cad & "Socio|rpozos.codsocio|N|000000|12·"
    cad = cad & "Nombre|rsocios.nomsocio|T||55·"
    cad = cad & "Nro.Orden|rpozos.nroorden|T||15·"
    
    NombreTabla1 = "(rpozos inner join rsocios on rpozos.codsocio = rsocios.codsocio)"
    
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = NombreTabla1
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Hidrantes" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 0

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
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
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
    
    PonerModo 0
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        ' *** canviar o llevar, si cal, el WHERE; repasar codEmpre ***
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        'CadenaConsulta = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
        ' ******************************************
        PonerCadenaBusqueda
        ' *** si n'hi han llínies sense grids ***
'        CargaFrame 0, True
        ' ************************************
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("rentradas", "numnotac")
'    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    ' ***********************************************************
    ' *** si n'hi han camps de descripció a la capçalera ***
    PosarDescripcions
    ' ******************************************************

    Text1(7).Text = "0"
    Text1(9).Text = "0"
    
    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
    ' ********************************************
End Sub


Private Sub BotonModificar()
Dim i As Integer

    PonerModo 4
    
    SocioAnt = Text1(2).Text


    Set ListOrigen = New Collection

    For i = 0 To Text1.Count - 1
        ListOrigen.Add Text1(i).Text
    Next i


    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
    ' *********************************************************
End Sub



Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar el Hidrante?"
    cad = cad & vbCrLf & "Hidrante: " & Data1.Recordset.Fields(0)
    ' **************************************************************************
    
    'borrem
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
        ' ********************************************************
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Cliente", Err.Description
End Sub


Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 2, "Frame2"   'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    For i = 0 To DataGridAux.Count - 1
        CargaGrid i, True
        If Not AdoAux(i).Recordset.EOF Then _
            PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
    Next i
    ' *******************************************
    SumaTotalPorcentajes 0
    
    PosarDescripcions
    
    '[Monica]15/05/2013: Visualizamos los cobros pendientes del socio
    ComprobarCobrosSocio CStr(Data1.Recordset!Codsocio), ""
    
    If ConexionIndefa Then
        PosarDescripcionsIndefa
        PosarDescripcionsIndefa2
        PosarDescripcionsIndefa3
        PosarDescripcionsIndefa4
        PosarDescripcionsIndefa5
    End If

' lo he quitado de aqui pq el consumo está almacenado en un campo de la tabla rpozos
'    CalcularConsumo
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari
    
End Sub


Private Sub PosarDescripcionsIndefa()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String
Dim Contador As String
Dim RS As ADODB.Recordset
Dim i As Integer

    On Error GoTo EPosarDescripcions


    ' Limpiamos los campos de indefa
    For i = 0 To 19
        txtaux1(i).Text = ""
    Next i
    
    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Or Not IsNumeric(Contador) Then Exit Sub
    '[Monica]19/02/2014: en la columna de nombre_socio(cc1) antes estaba ""codigotch"", no viene en la base de datos nombre_socio de rae_visitas_hidtomas
    Sql = "select ""Contador"" as contador1, socio_revisado, '' as ccc1, poligono, parcelas, hanegadas, dn_contador, dn_valvula_esfera, dn_toma, codigotoma,"
    Sql = Sql & """Recibido"" as aaa , fecha_entrada, instaladora, fecha_revision, fecha_turno,turno, q_instantaneo,verificacion,""observaciones_RAE"" as ccc,"
    Sql = Sql & """Comentarios_INDEFA"" as hhh"
    Sql = Sql & " from rae_visitas_hidtomas where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante = " & DBSet(Int(Mid(Contador, 3, 2)), "T")
    
    '[Monica]18/07/2013: cambio
                                    '[Monica]27/01/2014: lo cambio a numerico
    Sql = Sql & " and salida_tch = " & DBSet(Int(Mid(Contador, 5, 2)), "N")
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        For i = 0 To 19
            txtaux1(i).Text = DBLet(RS.Fields(i).Value)
        Next i
    End If
    If txtaux1(5).Text <> "" Then txtaux1(5).Text = Format(txtaux1(5).Text, "###,##0.00")
    Set RS = Nothing
    
EPosarDescripcions:
    If Err.Number <> 0 Then MsgBox "Han cambiado datos de Contadores de Indefa", vbExclamation
End Sub
' ************************************************************

Private Sub PosarDescripcionsIndefa2()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String
Dim Contador As String
Dim RS As ADODB.Recordset
Dim i As Integer

    On Error GoTo EPosarDescripcions


    ' Limpiamos los campos de indefa
    For i = 0 To txtaux8.Count - 1
        txtaux8(i).Text = ""
    Next i
    
    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Or Not IsNumeric(Contador) Then Exit Sub
    
    Sql = "select sector, hidrante1, hidrante2, valvula_aislamiento, punto_entrega, fto_valvula, tipo_arqueta, tipo_tapa, observaciones, foto_desague, foto_general "
    Sql = Sql & " from rae_visitas_desagues "
    Sql = Sql & " where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante1 = " & DBSet(Int(Mid(Contador, 3, 2)), "T")

    Set RS = New ADODB.Recordset
    RS.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Me.Toolbar3(2).Tag = ""
    Me.Toolbar3(3).Tag = ""
    
    If Not RS.EOF Then
        For i = 0 To txtaux8.Count - 1
            txtaux8(i).Text = DBLet(RS.Fields(i).Value)
        Next i
    
        If Dir(App.Path & "\FotosHts\" & RS!foto_desague) <> "" Then
            Me.Toolbar3(2).Tag = App.Path & "\FotosHts\" & RS!foto_desague
            Image8.Picture = LoadPicture(Me.Toolbar3(2).Tag)
            
        End If
        
        If Dir(App.Path & "\FotosHts\" & RS!foto_general) <> "" Then
            Me.Toolbar3(3).Tag = App.Path & "\FotosHts\" & RS!foto_general
            Image9.Picture = LoadPicture(Me.Toolbar3(3).Tag)
        End If
    End If
    Set RS = Nothing
    
    
EPosarDescripcions:
    If Err.Number <> 0 Then MsgBox "Han cambiado datos de Desagües de Indefa", vbExclamation
End Sub
' ************************************************************



Private Sub PosarDescripcionsIndefa3()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String
Dim Contador As String
Dim RS As ADODB.Recordset
Dim i As Integer
Dim vTabAnt As Integer

    On Error GoTo EPosarDescripcions



    ' Limpiamos los campos de indefa
    For i = 0 To txtaux5.Count - 1
        txtaux5(i).Text = ""
    Next i
    
    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Or Not IsNumeric(Contador) Then Exit Sub
    
    Sql = "select codigo_hidrante, constructora, fecha, valvula_compuerta, estado_colector, tch_tipo, tch_fijacion_colector, caja_empalmes_tch, nivelacion_arqueta_verticalidad, observaciones,  "
    Sql = Sql & " foto1, foto2"
    Sql = Sql & " from rae_visitas_hidrantes "
    Sql = Sql & " where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante = " & DBSet(Int(Mid(Contador, 3, 2)), "T")
    

    Set RS = New ADODB.Recordset
    RS.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Me.Toolbar3(0).Tag = ""
    Me.Toolbar3(1).Tag = ""
    Image2.Picture = LoadPicture("")
    Image3.Picture = LoadPicture("")
    
    If Not RS.EOF Then
    
        For i = 0 To txtaux5.Count - 1
            txtaux5(i).Text = DBLet(RS.Fields(i).Value)
        Next i
        If Dir(App.Path & "\FotosHts\" & RS!foto1 & ".jpg") <> "" Then
            Me.Toolbar3(0).Tag = App.Path & "\FotosHts\" & RS!foto2 & ".jpg"
            Image2.Picture = LoadPicture(Me.Toolbar3(0).Tag)
        End If
        
        If Dir(App.Path & "\FotosHts\" & RS!foto2 & ".jpg") <> "" Then
            Me.Toolbar3(1).Tag = App.Path & "\FotosHts\" & RS!foto1 & ".jpg"
            Image3.Picture = LoadPicture(Me.Toolbar3(1).Tag)
        End If
    End If
    DoEvents

'    Me.Toolbar3(0).Buttons(1).Enabled = (Me.Toolbar3(0).Tag <> "")
'    Me.Toolbar3(1).Buttons(1).Enabled = (Me.Toolbar3(1).Tag <> "")
    Set RS = Nothing

EPosarDescripcions:
    If Err.Number <> 0 Then MsgBox "Han cambiado datos de Hidrantes de Indefa", vbExclamation
End Sub
' ************************************************************


Private Sub PosarDescripcionsIndefa4()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String
Dim Contador As String
Dim RS As ADODB.Recordset
Dim i As Integer

    On Error GoTo EPosarDescripcions


    ' Limpiamos los campos de indefa
    For i = 0 To txtaux6.Count - 1
        txtaux6(i).Text = ""
    Next i
    
    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Or Not IsNumeric(Contador) Then Exit Sub
    
    Sql = "select sector, hidrante1, hidrante2, dn_tuberia_instalada, valvula_mariposa, uniones, eje_en_valvula_de_mariposa, comprobacion,"
    Sql = Sql & " tipologia_arqueta, tapa_fundicion_en_conos, situacion_tapa_arqueta, observaciones, foto2, foto_valvulas_aislamiento"
    Sql = Sql & " from rae_visitas_valvulas "
    Sql = Sql & " where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante1 = " & DBSet(Int(Mid(Contador, 3, 2)), "T")

    Set RS = New ADODB.Recordset
    RS.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    Me.Toolbar3(4).Tag = ""
    Me.Toolbar3(5).Tag = ""
    Me.Image4.Picture = LoadPicture("")
    Me.Image5.Picture = LoadPicture("")
    
    If Not RS.EOF Then
        For i = 0 To txtaux6.Count - 1
            txtaux6(i).Text = DBLet(RS.Fields(i).Value)
        Next i
    
        If InStr(1, RS!foto_valvulas_aislamiento, "http") <> 0 Then
            Me.Toolbar3(4).Tag = RS!foto_valvulas_aislamiento
'            Image5.Picture = LoadPicture(Me.Toolbar3(4).Tag)
        Else
            If DBLet(RS!foto_valvulas_aislamiento, "T") <> "" Then
                If Dir(App.Path & "\FotosHts\" & RS!foto_valvulas_aislamiento) <> "" Then
                    Me.Toolbar3(4).Tag = App.Path & "\FotosHts\" & RS!foto_valvulas_aislamiento
                    Image5.Picture = LoadPicture(Me.Toolbar3(4).Tag)
                End If
            End If
        End If
        If InStr(1, RS!foto2, "http") <> 0 Then
            Me.Toolbar3(5).Tag = RS!foto2
'            Image4.Picture = LoadPicture(Me.Toolbar3(5).Tag)
        Else
            If DBLet(RS!foto2, "T") <> "" Then
                If Dir(App.Path & "\FotosHts\" & RS!foto2) <> "" Then
                    Me.Toolbar3(5).Tag = App.Path & "\FotosHts\" & RS!foto2
                    Image4.Picture = LoadPicture(Me.Toolbar3(5).Tag)
                End If
            End If
        End If
    End If
    
    Set RS = Nothing
    
    
EPosarDescripcions:
    If Err.Number <> 0 Then MsgBox "Han cambiado datos de Válvulas de Indefa", vbExclamation
End Sub
' ************************************************************



Private Sub PosarDescripcionsIndefa5()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String
Dim Contador As String
Dim RS As ADODB.Recordset
Dim i As Integer

    On Error GoTo EPosarDescripcions


    ' Limpiamos los campos de indefa
    For i = 0 To txtaux7.Count - 1
        txtaux7(i).Text = ""
    Next i
    
    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Or Not IsNumeric(Contador) Then Exit Sub
    
    Sql = "select sector, hidrante1, hidrante2, diametro_ventosa, diametro_tuberia_plano, aislamiento, comprobacion, tipologia_arqueta, tapa_arqueta, situacion, observaciones, foto2, foto_ventosa"
    Sql = Sql & " from rae_visitas_ventosas "
    Sql = Sql & " where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante1 = " & DBSet(Int(Mid(Contador, 3, 2)), "T")

    Set RS = New ADODB.Recordset
    RS.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Me.Toolbar3(6).Tag = ""
    Me.Toolbar3(7).Tag = ""
    
    If Not RS.EOF Then
        For i = 0 To txtaux7.Count - 1
            txtaux7(i).Text = DBLet(RS.Fields(i).Value)
        Next i
        
        If InStr(1, RS!foto_ventosa, "http") <> 0 Then
            Me.Toolbar3(6).Tag = RS!foto_ventosa
'            Image6.Picture = LoadPicture(Me.Toolbar3(6).Tag)
        Else
            If DBLet(RS!foto_ventosa, "T") <> "" Then
                If Dir(App.Path & "\FotosHts\" & RS!foto_ventosa) <> "" Then
                    Me.Toolbar3(6).Tag = App.Path & "\FotosHts\" & RS!foto_ventosa
                    Image6.Picture = LoadPicture(Me.Toolbar3(6).Tag)
                End If
            End If
        End If
        If InStr(1, RS!foto2, "http") <> 0 Then
            Me.Toolbar3(7).Tag = RS!foto2
'            Image7.Picture = LoadPicture(Me.Toolbar3(7).Tag)
        Else
            If DBLet(RS!foto2, "T") <> "" Then
                If Dir(App.Path & "\FotosHts\" & RS!foto2 & ".jpg") <> "" Then
                    Me.Toolbar3(7).Tag = App.Path & "\FotosHts\" & RS!foto2
                    Image7.Picture = LoadPicture(Me.Toolbar3(7).Tag)
                End If
            End If
        End If
    End If
    
    
    Set RS = Nothing
    
    
EPosarDescripcions:
    If Err.Number <> 0 Then MsgBox "Han cambiado datos de Ventosas de Indefa", vbExclamation
End Sub
' ************************************************************











Private Sub CalcularConsumo()
Dim Sql As String
Dim Inicio As Long
Dim Fin As Long
Dim Consumo As Long
Dim NroDig As Integer
Dim Limite As Long

    If Text1(9).Text = "" Then Exit Sub

    Inicio = 0
    Fin = 0
    
    If Text1(7).Text <> "" Then Inicio = CLng(Text1(7).Text)
    If Text1(9).Text <> "" Then Fin = CLng(Text1(9).Text)
    
    NroDig = CCur(Text1(12).Text)
    Limite = (10 ^ NroDig)
    
    If Fin >= Inicio Then
        Consumo = Fin - Inicio
    Else
        If MsgBox("¿ Es un reinicio de contador ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Consumo = (Limite - Inicio) + Fin
        Else
            Consumo = Fin - Inicio
        End If
    End If
    
    If Consumo > (Limite - 1) Or Consumo < 0 Then
        MsgBox "Error en la lectura.", vbExclamation
        PonerFoco Text1(9)
    End If
    
   
'    Text2(0).Text = Format(Consumo, "#,###,##0")
    '[Monica]11/06/2013: cambio el formato del consumo
    Text1(19).Text = Format(Consumo, "########0")

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
                ' ***************************************************

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
                ' *******************************************
        
        Case 5 'LLÍNIES
           Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 3 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        ' ********************************************************
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                    ' *** els tabs que no tenen datagrid ***
                    ElseIf NumTabMto = 3 Then
                        If BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2") Then BotonModificar 'Modificar
'                        CargaFrame 3, True
                    End If

                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ************************

                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0

                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ***********************

                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

                    ' ***  bloquejar els camps fora dels grids ***

                    ' ***  bloquejar els camps fora dels grids ***
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select
            
            SumaTotalPorcentajes NumTabMto

            PosicionarData

            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
          
          
          
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    b = CompForm2(Me, 2, "Frame2")
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    If b And (Modo = 3 Or Modo = 4) Then
        If b Then
'            If Not EstaSocioDeAlta(Text1(2).Text) Then
'            ' comprobamos que el socio no este dado de baja
'                SQL = "El socio introducido está dado de baja. Reintroduzca. " & vbCrLf & vbCrLf
'                MsgBox SQL, vbExclamation
'                b = False
'                PonerFoco Text1(2)
'            End If

'            If Text2(0).Text = "" Then Text2(0).Text = "0"
'            Text1(19).Text = Text2(0).Text
'            If CCur(Text2(0).Text) < 0 Then
'                MsgBox "El consumo no puede ser negativo. Revise.", vbExclamation
'                PonerFoco Text1(9)
'                b = False
'            End If
        
            CalcularConsumo
            If CCur(ComprobarCero(Text1(19).Text)) < 0 Then
                MsgBox "El consumo no puede ser negativo. Revise.", vbExclamation
                PonerFoco Text1(9)
                b = False
            End If
        End If
        
        '[Monica]04/02/2013: si el campo introducido no existe daba un error sin controlar
        If b Then
            If Text1(18).Text <> "" Then
                Sql = "select count(*) from rcampos where codcampo = " & DBSet(Text1(18).Text, "N")
                If TotalRegistros(Sql) = 0 Then
                    MsgBox "No existe el campo. Reintroduzca.", vbExclamation
                    PonerFoco Text1(18)
                    b = False
                End If
            End If
        End If
        
        If b Then
            If Text1(18).Text <> "" And Text1(2).Text <> "" Then
                Sql = "select count(*) from rcampos where codcampo = " & DBSet(Text1(18).Text, "N")
                Sql = Sql & " and codsocio = " & DBSet(Text1(2).Text, "N")
                If TotalRegistros(Sql) = 0 Then
                    If MsgBox("El campo introducido no es del socio. ¿ Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                        PonerFoco Text1(18)
                        b = False
                    End If
                End If
            End If
        
        End If
    
    End If
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(hidrante='" & Text1(0).Text & "')"
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
    ' ***********************************************************************************
End Sub


Private Function eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE hidrante='" & Trim(Data1.Recordset!Hidrante) & "'"
        ' ***********************************************************************
        
    conn.Execute "Delete from rpozos_cooprop " & vWhere
    conn.Execute "Delete from rpozos_campos " & vWhere
        
    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        eliminar = False
    Else
        conn.CommitTrans
        eliminar = True
    End If
End Function


Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Sql As String


    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Select Case Index
        Case 1 'nro de orden
            PonerFormatoEntero Text1(1)

        Case 2 'SOCIO
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
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
                    If Modo = 3 Then
                        If TotalRegistros("select count(*) from rcampos where codsocio = " & DBSet(Text1(Index).Text, "N")) > 0 Then
                            Set frmMens = New frmMensajes
                            frmMens.cadwhere = "and rcampos.codsocio = " & DBSet(Text1(Index).Text, "N")
                            frmMens.campo = Text1(Index).Text
                            frmMens.OpcionMensaje = 29
                            frmMens.Show vbModal
                            Set frmMens = Nothing
                        End If
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 3 'PARTIDA
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rpartida", "nomparti", "codparti", "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Partida: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPar = New frmManPartidas
                        frmPar.DatosADevolverBusqueda = "0|1|"
                        frmPar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPar.Show vbModal
                        Set frmPar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                '[Monica]01/10/2014: en el caso de que exista la partida hemos de refrescar el nombre de la poblacion
                Else
                    
                    Text2(1).Text = DevuelveValor("select despobla from rpartida inner join rpueblos on rpartida.codpobla = rpueblos.codpobla where rpartida.codparti = " & DBSet(Text1(3).Text, "N"))
                
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 6 'hanegadas
            PonerFormatoDecimal Text1(Index), 7
                
        Case 7, 9 ' CONTADORES
            If Modo = 1 Then Exit Sub
            PonerFormatoEntero Text1(Index)
            CalcularConsumo
                
        Case 8, 10 'Fecha no comprobaremos que esté dentro de campaña
                    'Fecha de alta y fecha de baja
            PonerFormatoFecha Text1(Index), True
            
        Case 14 ' calibre
            PonerFormatoEntero Text1(Index)
            
        Case 15 ' acciones
            PonerFormatoEntero Text1(Index)
            
        Case 16, 17 ' fecha de alta y baja del contador
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            PonerFormatoFecha Text1(Index)
            
        Case 13 'Tipo de pozo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rtipopozos", "nompozo", "codpozo", "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Pozo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPoz = New frmPOZPozos
                        frmPoz.DatosADevolverBusqueda = "0|1|"
                        frmPoz.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPoz.Show vbModal
                        Set frmPoz = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        
        Case 18
            If PonerFormatoEntero(Text1(Index)) Then
                PonerDatosCampo (Text1(Index))
            End If
                        
    End Select
End Sub

Private Sub PonerDatosCampo(campo As String)
Dim cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim RS As ADODB.Recordset
Dim Sql As String


    If campo = "" Then Exit Sub
    
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

'[Monica]08/02/2013: quito esto pq si quieren traer los datos del campo desplegaran la lupa
'    '[Monica]22/11/2012: Preguntamos si quiere traer los datos del socio del campo
'    If (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) And Modo = 4 Then
'        Sql = "select rcampos.codsocio, rsocios.nomsocio from rcampos inner join rsocios on rcampos.codsocio = rsocios.codsocio where rcampos.codcampo = " & DBSet(Text1(18).Text, "N")
'
'        Set RS = New ADODB.Recordset
'        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        If DBLet(RS.Fields(0)) <> CLng(ComprobarCero(Text1(2).Text)) And Modo = 3 Then
'            Text1(2).Text = Format(DBLet(RS!CodSocio, "N"), "000000") ' codigo de socio del campo
'            Text2(2).Text = DBLet(RS!nomsocio, "T") ' nombre de socio
'
'           'If MsgBox("¿ Desea traer los datos de RAE al contador ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
'        End If
'
'        Set RS = Nothing
'
'        Exit Sub
'
'    End If

    cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rcampos.poligono, rcampos.parcela, rcampos.supcoope, rpueblos.despobla, rcampos.subparce, rcampos.codsocio "
    Cad1 = Cad1 & " from rcampos, rpartida, rpueblos "
    Cad1 = Cad1 & " where " & cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla"
     
    Set RS = New ADODB.Recordset
    RS.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Text1(3).Text = ""
        Text1(4).Text = ""
        Text1(5).Text = ""
        Text1(6).Text = ""
        Text2(1).Text = ""
        
        Text1(18).Text = campo
        PonerFormatoEntero Text1(18)
        Text1(3).Text = DBLet(RS.Fields(0).Value, "N") ' codigo de partida
        If Text1(3).Text <> "" Then Text1(3).Text = Format(Text1(3).Text, "0000")
        Text2(3).Text = DBLet(RS.Fields(1).Value, "T") ' nombre de partida
        Text2(1).Text = DBLet(RS.Fields(5).Value, "T") ' nombre de poblacion
        Text1(4).Text = DBLet(RS.Fields(2).Value, "N") ' poligono
'[Monica]03/08/2012: quito el formato de poligono y parcela
'        If Text1(4).Text <> "" Then Text1(4).Text = Format(Text1(4).Text, "0000")
        Text1(5).Text = DBLet(RS.Fields(3).Value, "N") ' parcela
        
        If vParamAplic.Cooperativa = 10 Then Text1(5).Text = Text1(5).Text & " " & DBLet(RS.Fields(6).Value)
        
'        If Text1(5).Text <> "" Then Text1(5).Text = Format(Text1(5).Text, "000000")
        
        'hanegadas
        Text1(6).Text = Format(Round2(DBLet(RS.Fields(4).Value, "N") / vParamAplic.Faneca, 2), "##,##0.00")
    End If
    
    Set RS = Nothing
    
End Sub

Private Sub PonerDatosCampoLineas(campo As String)
Dim cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim RS As ADODB.Recordset
Dim i As Integer


    If campo = "" Then Exit Sub

    cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rcampos.poligono, rcampos.parcela, rcampos.supcoope, rpueblos.despobla "
    Cad1 = Cad1 & " from rcampos, rpartida, rpueblos "
    Cad1 = Cad1 & " where " & cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla"
     
    Set RS = New ADODB.Recordset
    RS.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        For i = 1 To 5
            txtAux2(i).Text = ""
        Next i
        
        txtaux4(2).Text = campo
        PonerFormatoEntero txtaux4(2)
        txtAux2(1).Text = DBLet(RS.Fields(1).Value, "T") ' nombre de partida
        txtAux2(2).Text = DBLet(RS.Fields(5).Value, "T") ' nombre de poblacion
        txtAux2(4).Text = DBLet(RS.Fields(2).Value, "N") ' poligono
        If txtAux2(4).Text <> "" Then txtAux2(4).Text = Format(txtAux2(4).Text, "0000")
        txtAux2(5).Text = DBLet(RS.Fields(3).Value, "N") ' parcela
        If txtAux2(5).Text <> "" Then txtAux2(5).Text = Format(txtAux2(5).Text, "000000")
        
        'hanegadas
        txtAux2(3).Text = Format(Round2(DBLet(RS.Fields(4).Value, "N") / vParamAplic.Faneca, 2), "##,##0.00")
    End If
    
    Set RS = Nothing
    
End Sub


Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'socio
                Case 3: KEYBusqueda KeyAscii, 1 'partida
                Case 8: KEYFecha KeyAscii, 0 'fecha desde
                Case 10: KEYFecha KeyAscii, 1 'fecha desde
            End Select
        End If
    Else
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

' **** si n'hi han camps de descripció a la capçalera ****
Private Sub PosarDescripcions()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String

    On Error GoTo EPosarDescripcions

    Text2(2).Text = PonerNombreDeCod(Text1(2), "rsocios", "nomsocio", "codsocio", "N")
    Text2(3).Text = PonerNombreDeCod(Text1(3), "rpartida", "nomparti", "codparti", "N")
    Text2(13).Text = PonerNombreDeCod(Text1(13), "rtipopozos", "nompozo", "codpozo", "N")
        
        
    If Text1(3).Text <> "" Then
        Sql = "select despobla from rpueblos, rpartida where rpartida.codparti = " & DBSet(Text1(3).Text, "N")
        Sql = Sql & " and rpueblos.codpobla = rpartida.codpobla "
        
        Text2(1).Text = DevuelveValor(Sql) ' nombre de poblacion
    End If
    
EPosarDescripcions:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo descripciones", Err.Description
End Sub
' ************************************************************



' *** si n'hi han formularis de buscar codi a les llínies ***
Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
       Case 0 'Socios
            Set frmSoc = New frmManSocios
'            frmSoc.DeConsulta = True
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(1).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(2)
    
       Case 1 'Partidas
            Set frmPar = New frmManPartidas
            frmPar.DeConsulta = True
            frmPar.DatosADevolverBusqueda = "0|1|"
            frmPar.CodigoActual = Text1(3).Text
            frmPar.Show vbModal
            Set frmPar = Nothing
            PonerFoco Text1(3)
    
       Case 2 'Tipo de Pozos
            Set frmPoz = New frmPOZPozos
            frmPoz.DeConsulta = True
            frmPoz.DatosADevolverBusqueda = "0|1|"
            frmPoz.CodigoActual = Text1(3).Text
            frmPoz.Show vbModal
            Set frmPoz = Nothing
            PonerFoco Text1(13)
    
       Case 3 'Campo
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|1|"
'            frmCam.CodigoActual = Text1(18).Text
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(18)
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 2, "Frame2"
End Sub


' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    If numTab = 0 Then
        SSTab1.Tab = 1
    ElseIf numTab = 1 Then
        SSTab1.Tab = 2
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codsocio=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function


Private Sub printNou()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim Sql As String
Dim ConFechaBaja As Boolean

    ' pedimos el orden del informe
    Set frmMen2 = New frmMensajes
    
    frmMen2.OpcionMensaje = 38
    frmMen2.Show vbModal
    
    Set frmMen2 = Nothing
    
    ConFechaBaja = False
    If MsgBox("¿ Desea imprimir los contadores con fecha de baja ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        ConFechaBaja = True
    End If
    
    indRPT = 78 ' personalizacion del informe de hidrantes
    
    If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
    
    
    With frmImprimir2
        .cadTabla2 = "rpozos"
        .Informe2 = nomDocu
        If CadB <> "" Then
            If InStr(CadB, "in (") <> 0 Then
                .cadRegSelec = ""
            
                Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
                conn.Execute Sql
            
                Sql = "insert into tmpinformes (codusu, nombre1) select " & vUsu.Codigo & ", " & Replace(Data1.RecordSource, "select * ", "hidrante ")
                conn.Execute Sql
            
                .Informe2 = Replace(nomDocu, ".rpt", "1.rpt")
            
            Else
                .cadRegSelec = SQL2SF(CadB)
            
            End If
        Else
            .cadRegSelec = ""
        End If
        If ConFechaBaja Then
            If .cadRegSelec <> "" Then .cadRegSelec = .cadRegSelec & " and "
            .cadRegSelec = .cadRegSelec & "isnull({rpozos.fechabaja})"
        End If
        .cadRegActua = POS2SF(Data1, Me)
        '[Monica]13/07/2012: falla si hay un solo registro seleccionado y apretamos registros buscados
        If Data1.Recordset.RecordCount = 1 Then .cadRegSelec = .cadRegActua
        .cadTodosReg = ""
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|" & Orden & "|pUsu=" & vUsu.Codigo & "|"
        .NumeroParametros2 = 2
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = True
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub

Private Sub printNouIndefa()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
    
    
    indRPT = 78 ' personalizacion del informe de hidrantes
    
    If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
    
    nomDocu = "EscContadorIndefa.rpt"
    
    With frmImprimir2
        .cadTabla2 = "rpozos"
        .Informe2 = nomDocu
        If CadB <> "" Then
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(Data1, Me)
        '[Monica]13/07/2012: falla si hay un solo registro seleccionado y apretamos registros buscados
        If Data1.Recordset.RecordCount = 1 Then .cadRegSelec = .cadRegActua
        .cadTodosReg = ""
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|" & Orden
        .NumeroParametros2 = 2
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = True
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub




'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadparam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadparam = ""
    numParam = 0
End Sub

Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If InsertarOferta(Sql, vTipoMov) Then
            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            PonerCadenaBusqueda
            PonerModo 2
        End If
    End If
    Text1(0).Text = Format(Text1(0).Text, "0000000")
End Sub

Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String

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
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador."
    vTipoMov.IncrementarContador (CodTipoMov)
    
EInsertarOferta:
    If Err.Number <> 0 Then
        MenError = "Insertando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
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
        Case 0
            Sql = "select rpozos_cooprop.hidrante, rpozos_cooprop.numlinea, rpozos_cooprop.codsocio, rsocios.nomsocio, "
            Sql = Sql & " rpozos_cooprop.porcentaje "
            Sql = Sql & " FROM rpozos_cooprop INNER JOIN rsocios ON rpozos_cooprop.codsocio = rsocios.codsocio "
            Sql = Sql & " and rpozos_cooprop.codsocio = rsocios.codsocio "
            If enlaza Then
                Sql = Sql & " WHERE rpozos_cooprop.hidrante = '" & Trim(Text1(0).Text) & "'"
            Else
                Sql = Sql & " WHERE rpozos_cooprop.hidrante is null"
            End If
            Sql = Sql & " ORDER BY rpozos_cooprop.codsocio "
        
        Case 1
            Sql = "select rpozos_campos.hidrante, rpozos_campos.numlinea, rpozos_campos.codcampo, rpartida.nomparti, "
            Sql = Sql & " rpueblos.despobla, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) supcoope, rcampos.poligono, rcampos.parcela "
            Sql = Sql & " FROM ((rpozos_campos INNER JOIN rcampos ON rpozos_campos.codcampo = rcampos.codcampo)"
            Sql = Sql & " INNER JOIN rpartida ON rcampos.codparti = rpartida.codparti) "
            Sql = Sql & " INNER JOIN rpueblos ON rpartida.codpobla = rpueblos.codpobla "
            If enlaza Then
                Sql = Sql & " WHERE rpozos_campos.hidrante = '" & Trim(Text1(0).Text) & "'"
            Else
                Sql = Sql & " WHERE rpozos_campos.hidrante is null"
            End If
            Sql = Sql & " ORDER BY rpozos_campos.codcampo "
       
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function
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
Dim eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'coopropietarios
            Sql = "¿Seguro que desea eliminar el coopropietario?"
            Sql = Sql & vbCrLf & "Coopropietario: " & AdoAux(Index).Recordset!Codsocio & " - " & AdoAux(Index).Recordset!nomsocio
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rpozos_cooprop"
                Sql = Sql & " WHERE rpozos_cooprop.hidrante = " & DBSet(AdoAux(Index).Recordset!Hidrante, "T")
                Sql = Sql & " and codsocio = " & AdoAux(Index).Recordset!Codsocio
            End If
        Case 1 ' campos
            Sql = "¿Seguro que desea eliminar el campo del hidrante?"
            Sql = Sql & vbCrLf & "Campo: " & AdoAux(Index).Recordset!codcampo
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rpozos_campos"
                Sql = Sql & " WHERE rpozos_campos.hidrante = " & DBSet(AdoAux(Index).Recordset!Hidrante, "T")
                Sql = Sql & " and numlinea = " & AdoAux(Index).Recordset!numlinea
            End If
        
    End Select

    If eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
            
        End If
        SumaTotalPorcentajes NumTabMto
        ' *** si n'hi han tabs sense datagrid ***
        ' ***************************************
        If BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2") Then BotonModificar
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto)
        ' ************************
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
Dim vWhere As String, vTabla As String
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
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vTabla = "rpozos_cooprop"
        Case 1: vTabla = "rpozos_campos"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 0, 1, 2 'clasificacion
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            Select Case Index
                Case 0
                    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", "rpozos_cooprop.hidrante = '" & Trim(Text1(0).Text) & "'")
                Case 1
                    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", "rpozos_campos.hidrante = '" & Trim(Text1(0).Text) & "'")
            End Select
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)

            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

            LLamaLineas Index, ModoLineas, anc

            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'copropietarios
                    For i = 0 To txtaux3.Count - 1
                        txtaux3(i).Text = ""
                    Next i
                    txtAux2(0).Text = ""
                    txtaux3(0).Text = Text1(0).Text 'codcampo
                    txtaux3(1).Text = NumF 'numlinea
                    txtaux3(2).Text = ""
                    PonerFoco txtaux3(2)
                Case 1 ' campos
                    For i = 0 To txtaux4.Count - 1
                        txtaux4(i).Text = ""
                    Next i
                    For i = 1 To 5
                        txtAux2(i).Text = ""
                    Next i
                    txtaux4(0).Text = Text1(0).Text ' codcampo
                    txtaux4(1).Text = NumF 'numlinea
                    PonerFoco txtaux4(2)
                
            End Select


'        ' *** si n'hi han llínies sense datagrid ***
'        Case 3
'            LimpiarCamposLin "FrameAux3"
'            txtaux(42).Text = text1(0).Text 'codclien
'            txtaux(43).Text = vSesion.Empresa
'            Me.cmbAux(28).ListIndex = 0
'            Me.cmbAux(29).ListIndex = 1
'            PonerFoco txtaux(25)
'        ' ******************************************
    End Select
End Sub


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub

    ModoLineas = 2 'Modificar llínia

    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *********************************

    Select Case Index
        Case 0, 1, 2 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
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
        Case 0 'coopropietarios
            txtaux3(0).Text = DataGridAux(Index).Columns(0).Text
            txtaux3(1).Text = DataGridAux(Index).Columns(1).Text
            txtaux3(2).Text = DataGridAux(Index).Columns(2).Text
            
            txtAux2(0).Text = DataGridAux(Index).Columns(3).Text
            txtaux3(3).Text = DataGridAux(Index).Columns(4).Text
        
        Case 1 ' campos
            txtaux4(0).Text = DataGridAux(Index).Columns(0).Text
            txtaux4(1).Text = DataGridAux(Index).Columns(1).Text
            txtaux4(2).Text = DataGridAux(Index).Columns(2).Text
            
            txtAux2(1).Text = DataGridAux(Index).Columns(3).Text
            txtAux2(2).Text = DataGridAux(Index).Columns(4).Text
            txtAux2(3).Text = DataGridAux(Index).Columns(5).Text
            txtAux2(4).Text = DataGridAux(Index).Columns(6).Text
            txtAux2(5).Text = DataGridAux(Index).Columns(7).Text
        
    
    End Select

    LLamaLineas Index, ModoLineas, anc

    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'coopropietarios
            PonerFoco txtaux3(2)
        Case 1 ' campos
            PonerFoco txtaux4(2)
    End Select
    ' ***************************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************

    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 ' coopropietarios
            For jj = 2 To txtaux3.Count - 1
                txtaux3(jj).visible = b
                txtaux3(jj).Top = alto
            Next jj
            txtAux2(0).visible = b
            txtAux2(0).Top = alto
            cmdAux(0).visible = b
            cmdAux(0).Top = txtaux3(2).Top
            cmdAux(0).Height = txtaux3(2).Height
        Case 1 ' campos
            For jj = 2 To txtaux4.Count - 1
                txtaux4(jj).visible = b
                txtaux4(jj).Top = alto
            Next jj
            For jj = 1 To 5
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            cmdAux(1).visible = b
            cmdAux(1).Top = txtaux4(2).Top
            cmdAux(1).Height = txtaux4(2).Height
    
    End Select
End Sub




Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Búscar
            cmdSigpac_Click
    End Select
End Sub


Private Sub Toolbar3_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    
    Select Case Index
        Case 0, 1, 2, 3
            Select Case Button.Index
                Case 1  'Búscar
                    ' si existe el fichero de ampliacion lo mostramos
                    If Dir(Toolbar3(Index).Tag) <> "0" Then
                    
                        Set frmMensImg = New frmMensajes
                        
                        frmMensImg.CADENA = Toolbar3(Index).Tag
                        frmMensImg.OpcionMensaje = 45
                        frmMensImg.Show vbModal
                        
                        Set frmMens = Nothing
                    
                    End If
            End Select
        Case 4, 5, 6, 7
            Select Case Button.Index
                Case 1  'Búscar
                    If InStr(1, Toolbar3(Index).Tag, "http") <> 0 Then
                        Screen.MousePointer = vbHourglass
                    
                        If LanzaHomeGnral(Toolbar3(Index).Tag) Then espera 2
                        
                        Screen.MousePointer = vbDefault
                    Else
                        If Dir(Toolbar3(Index).Tag) <> "0" Then
                        
                            Set frmMensImg = New frmMensajes
                            
                            frmMensImg.CADENA = Toolbar3(Index).Tag
                            frmMensImg.OpcionMensaje = 45
                            frmMensImg.Show vbModal
                            
                            Set frmMensImg = Nothing
                        
                        End If
                    End If
            End Select
    End Select
     
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtaux3(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2 'NIF
            If PonerFormatoEntero(txtaux3(Index)) Then
                txtAux2(0).Text = PonerNombreDeCod(txtaux3(Index), "rsocios", "nomsocio")
                If txtAux2(0).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc1 = New frmManSocios
                        frmSoc1.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
                        txtaux3(Index).Text = ""
                        TerminaBloquear
                        frmSoc1.Show vbModal
                        Set frmSoc1 = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtaux3(Index).Text = ""
                    End If
                    PonerFoco txtaux3(Index)
                Else
                    ' comprobamos que el socio no esté dado de baja
                    If Not EstaSocioDeAlta(txtaux3(Index).Text) Then
                        If MsgBox("Este socio tiene fecha de baja. ¿ Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                            txtaux3(Index).Text = ""
                            txtAux2(0).Text = ""
                            PonerFoco txtaux3(Index)
                        End If
                    End If
                End If
            Else
                txtAux2(0).Text = ""
            End If
            
        Case 3 'porcentaje de
            PonerFormatoDecimal txtaux3(Index), 4
            If txtaux3(2).Text <> "" Then cmdAceptar.SetFocus
    
    End Select

    ' ******************************************************************************
End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
    If Not txtaux3(Index).MultiLine Then ConseguirFocoLin txtaux3(Index)
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtaux3(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'AAAAAAAAAAAAAAAAAAAAAAA
Private Sub TxtAux4_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim i As Integer
Dim Sql As String


    If Not PerderFocoGnral(txtaux4(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2 ' campo
            If PonerFormatoEntero(txtaux4(Index)) Then
                Sql = ""
                Sql = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", txtaux4(Index).Text, "N")
                If Sql = "" Then
                    cadMen = "No existe el Campo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCam1 = New frmManCampos
                        frmCam1.DatosADevolverBusqueda = "0|1|"
'                        frmCamp.NuevoCodigo = Text1(Index).Text
                        txtaux4(Index).Text = ""
                        TerminaBloquear
                        frmCam1.Show vbModal
                        Set frmCam1 = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtaux4(Index).Text = ""
                    End If
                    PonerFoco txtaux4(Index)
                Else
                    If Not EstaCampoDeAlta(txtaux4(Index).Text) Then
                        MsgBox "El campo no está dado de alta. Reintroduzca.", vbExclamation
                        txtaux4(Index).Text = ""
                        PonerFoco txtaux4(Index)
                    Else
                        PonerDatosCampoLineas (txtaux4(Index))
                    End If
                End If
            Else
                For i = 1 To 5
                    txtAux2(i).Text = ""
                Next i
            End If
            
        Case 3 'porcentaje de
            PonerFormatoDecimal txtaux4(Index), 4
            If txtaux4(2).Text <> "" Then cmdAceptar.SetFocus
    
    End Select

End Sub

Private Sub TxtAux4_GotFocus(Index As Integer)
    If Not txtaux4(Index).MultiLine Then ConseguirFocoLin txtaux4(Index)
End Sub

Private Sub TxtAux4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtaux4(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux4_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'AAAAAAAAAAAAAAAAAAAAAAA
Private Function DatosOkLlin(nomframe As String) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    
    If b And (Modo = 5 And ModoLineas = 1) And nomframe = "FrameAux0" Then  'insertar
        'comprobar que el porcentaje sea distinto de cero
        If txtaux3(3).Text = "" Then
            MsgBox "El porcentaje de coopropiedad debe ser superior a 0.", vbExclamation
            PonerFoco txtaux3(3)
            b = False
        Else
            If CInt(txtaux3(3).Text) = 0 Then
                MsgBox "El porcentaje de coopropiedad debe ser superior a 0.", vbExclamation
                PonerFoco txtaux3(3)
                b = False
            End If
        End If
    End If
    
'
'    ' *** si cal fer atres comprovacions a les llínies (en o sense tab) ***
'    Select Case NumTabMto
'        Case 0  'CUENTAS BANCARIAS
'            SQL = "SELECT COUNT(ctaprpal) FROM cltebanc "
'            SQL = SQL & ObtenerWhereCab(True) & " AND ctaprpal=1"
'            If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
'
'            RS.Close
'            Set RS = Nothing
''yo
''            'no n'hi ha cap conter principal i ha seleccionat que no
''            If (Cant = 0) And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 0) Then
''                Mens = "Debe una haber una cuenta principal"
''            ElseIf (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) And (cmbAux(9).ItemData(cmbAux(9).ListIndex) = 0) Then
''                Mens = "Debe seleccionar que esta cuenta está activa si desea que sea la principal"
''            End If
'
''            'No puede haber más de una cuenta principal
''            If cant > 0 And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
''                Mens = "No puede haber más de una cuenta principal."
''            End If
''yo
''            'No pueden haber registros con el mismo: codbanco-codsucur-digcontr-ctabanc
''            If Mens = "" Then
''                SQL = "SELECT count(codclien) FROM cltebanc "
''                SQL = SQL & " WHERE codclien=" & text1(0).Text & " AND codempre= " & vSesion.Empresa
''                If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
''                SQL = SQL & " AND codnacio=" & cmbAux(0).ItemData(cmbAux(0).ListIndex)
''                SQL = SQL & " AND codbanco=" & txtaux(3).Text & " AND codsucur=" & txtaux(4).Text
''                SQL = SQL & " AND digcontr='" & txtaux(5).Text & "' AND ctabanco='" & txtaux(6).Text & "'"
''                Set RS = New ADODB.Recordset
''                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''                Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
''                If Cant > 0 Then
''                    Mens = "Ya Existe la cuenta bancaria: " & cmbAux(0).List(cmbAux(0).ListIndex) & " - " & txtaux(3).Text & "-" & txtaux(4).Text & "-" & txtaux(5).Text & "-" & txtaux(6).Text
''                End If
''                RS.Close
''                Set RS = Nothing
''            End If
''
''            If Mens <> "" Then
''                Screen.MousePointer = vbNormal
''                MsgBox Mens, vbExclamation
''                DatosOkLlin = False
''                'PonerFoco txtAux(3)
''                Exit Function
''            End If
''
'    End Select
'    ' ******************************************************************************
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    b = DataGridAux(Index).Enabled
    DataGridAux(Index).Enabled = False
    
    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    AdoAux(Index).Refresh
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 290
    If PrimeraVez Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    End If

    For i = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(i).AllowSizing = False
    Next i
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    'DataGridAux(Index).Enabled = b
'    PrimeraVez = False
    
    Select Case Index
        Case 0 ' coopropietarios
            tots = "N||||0|;N||||0|;S|txtaux3(2)|T|Cód.|1000|;S|cmdAux(0)|B|||;" 'codsocio,numlinea
            tots = tots & "S|txtAux2(0)|T|Nombre|3870|;"
            tots = tots & "S|txtaux3(3)|T|Porcentaje|1200|;"
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgRight
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

'            BloquearTxt txtAux(14), Not b
'            BloquearTxt txtAux(15), Not b

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                SumaTotalPorcentajes
            Else
                For i = 0 To 3
                    txtaux3(i).Text = ""
                Next i
                txtAux2(0).Text = ""
            End If
        Case 1 ' CAMPOS
            tots = "N||||0|;N||||0|;S|txtaux4(2)|T|Campo|1000|;S|cmdAux(1)|B|||;" 'codsocio,numlinea
            tots = tots & "S|txtAux2(1)|T|Partida|1800|;"
            tots = tots & "S|txtAux2(2)|T|Población|1470|;"
            tots = tots & "S|txtAux2(3)|T|Hdas|800|;"
            tots = tots & "S|txtAux2(4)|T|Pol|400|;"
            tots = tots & "S|txtAux2(5)|T|Par|600|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(5).Alignment = dbgRight
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))


            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                SumaTotalPorcentajes
            Else
                For i = 2 To 2
                    txtaux4(i).Text = ""
                Next i
                For i = 1 To 5
                    txtAux2(i).Text = ""
                Next i
            End If
         
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
        LimpiarCamposFrame Index
    End If
    ' **********************************************************
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub LimpiarCamposFrame(Index As Integer)
'Dim I As Integer
'    On Error Resume Next
'
'    Select Case Index
'        Case 0 'telefonos
'            For I = 0 To txtAux.Count - 1
'                txtAux(I).Text = ""
'            Next I
'        Case 1 'secciones
'            For I = 0 To txtaux1.Count - 1
'                txtaux1(I).Text = ""
'            Next I
'    End Select
'
'    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
'Dim I As Byte
'
'    If ModoLineas <> 1 Then
'        Select Case Index
'            Case 0 'telefonos
'                If DataGridAux(Index).Columns.Count > 2 Then
'                    For I = 5 To txtAux.Count - 1
'                        txtAux(I).Text = DataGridAux(Index).Columns(I).Text
'                    Next I
'                    Me.chkAbonos(1).Value = DataGridAux(Index).Columns(17).Text
'
'                End If
'            Case 1 'secciones
'                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux2(4).Text = ""
'                    txtAux2(5).Text = ""
'                    txtAux2(0).Text = ""
'                    Set vSeccion = New CSeccion
'                    If vSeccion.LeerDatos(AdoAux(1).Recordset!codsecci) Then
'                        If vSeccion.AbrirConta Then
'                            If DBLet(AdoAux(1).Recordset!codmaccli, "T") <> "" Then
'                                txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", AdoAux(1).Recordset!codmaccli, "T")
'                            End If
'                            If DBLet(AdoAux(1).Recordset!codmacpro, "T") <> "" Then
'                                txtAux2(5).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", AdoAux(1).Recordset!codmacpro, "T")
'                            End If
'                            txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", AdoAux(1).Recordset!CodIVA, "N")
'                            vSeccion.CerrarConta
'                        End If
'                    End If
'                    Set vSeccion = Nothing
'                End If
'        End Select
'    End If
End Sub


Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'coopropietarios
        Case 1: nomframe = "FrameAux1" 'campos
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            ' *************************************************
            b = BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2")
            
            '++monica: en caso de estar insertando seccion y que no existan las
            'cuentas contables hacemos esto para que las inserte en contabilidad.
'            If NumTabMto = 1 Then
'               txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 3, Text1(0))
'               txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 3, Text1(0))
'            End If
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
'                Case 3 ' *** els index dels tabs que NO tenen grid ***
'                    CargaFrame 3, True
'                    If b Then BotonModificar
'                    ModoLineas = 0
'                LLamaLineas NumTabMto, 0
                
            End Select
           
            'SituarTab (NumTabMto)
        End If
    End If
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'coopropietarios
        Case 1: nomframe = "FrameAux1" 'campos
        Case 2: nomframe = "FrameAux2" 'parcelas
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
            End If
            ' ******************************************************
            If BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2") Then BotonModificar
            ModoLineas = 0

            If NumTabMto <> 3 Then
                V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
            'SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
            
        End If
    End If
        
End Sub


Private Sub SumaTotalPorcentajes(numTab As Integer)
Dim Sql As String
Dim i As Currency
Dim RS As ADODB.Recordset
   
   Select Case numTab
        Case 0 ' coopropietarios
            Sql = "select sum(porcentaje) from rpozos_cooprop where rpozos_cooprop.hidrante = " & DBSet(Data1.Recordset!Hidrante, "T")
            
            Set RS = New ADODB.Recordset
            RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            i = 0
            If Not RS.EOF Then
                i = DBLet(RS.Fields(0).Value, "N")
            End If
            
            If i = 0 Then Exit Sub
            
            If i <> 100 Then
                NumTabMto = 0
                SituarTab numTab
                MsgBox "La suma de porcentajes es " & i & ". Debe de ser 100%. Revise.", vbExclamation
            End If
   
        
   End Select

End Sub


Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
'    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
'    ' ****************************************************
    
    SepuedeBorrar = True
End Function
    
    
Private Sub cmdSigpac_Click()
Dim Direccion As String
Dim Pobla As String
Dim Municipio As String
Dim RS As ADODB.Recordset
Dim Sql As String

    TerminaBloquear

    
    'http://sigpac.mapa.es/fega/visor/LayerInfo.aspx?layer=PARCELA&id=OID&image=ORTOFOTOS
'    Direccion = "http://sigpac.mapa.es/fega/visor/LayerInfo.aspx?layer=PARCELA&id=" & Trim(Text1(18).Text) & "&image=ORTOFOTOS"
    
    If vParamAplic.SigPac <> "" Then
        If InStr(1, vParamAplic.SigPac, "NUMOID") <> 0 Then
            Sql = "select numeroid from rcampos where codcampo = " & DBSet(Text1(18).Text, "N")
            
            Set RS = New ADODB.Recordset
            RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            If Not RS.EOF Then
                Direccion = Replace(vParamAplic.SigPac, "NUMOID", DBLet(RS!numeroid))
            End If
        Else
            If txtaux1(3).Text <> "" And txtaux1(4).Text <> "" Then
'                Sql = "select codparti, recintos from rcampos where codcampo = " & DBSet(Text1(18).Text, "N")
'
'                Set Rs = New ADODB.Recordset
'                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Not Rs.EOF Then
            
'                    Pobla = DevuelveDesdeBDNew(cAgro, "rpartida", "codpobla", "codparti", DBLet(Rs!codparti), "N")
                    Pobla = vParam.CPostal
                    If Pobla = "" Then
                        MsgBox "No existe el código de poblacion de la partida", vbExclamation
                    Else
                        Municipio = DevuelveDesdeBDNew(cAgro, "rpueblos", "codsigpa", "codpobla", Pobla, "T")
                        Direccion = Replace(vParamAplic.SigPac, "[PR]", Mid(Pobla, 1, 2))
                        Direccion = Replace(Direccion, "[MN]", CInt(Municipio))
                        Direccion = Replace(Direccion, "[PL]", CInt(ComprobarCero(txtaux1(3).Text)))
                        
                        If InStr(txtaux1(14).Text, ",") Then
                            'cogemos unicamente la primera parcela
                            Direccion = Replace(Direccion, "[PC]", CInt(ComprobarCero(Mid(txtaux1(4).Text, 1, InStr(txtaux1(4).Text, ",") - 1))))
                        Else
                            Direccion = Replace(Direccion, "[PC]", CInt(ComprobarCero(txtaux1(4).Text)))
                        End If
                        Direccion = Replace(Direccion, "[RC]", 1) 'CInt(ComprobarCero(Rs!recintos)))
                    End If
'                End If
            Else
                MsgBox "No existe el polígono y/o parcela Indefa.", vbExclamation
                Exit Sub
            End If
        End If
    Else
        MsgBox "No tiene configurada en parámetros la dirección de Sigpac. Llame a Soporte.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass

    If LanzaHomeGnral(Direccion) Then espera 2
    Screen.MousePointer = vbDefault


End Sub

