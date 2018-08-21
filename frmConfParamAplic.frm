VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfParamAplic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de la Aplicación"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   Icon            =   "frmConfParamAplic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   210
      TabIndex        =   430
      Top             =   90
      Width           =   1365
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   431
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   150
         TabIndex        =   432
         Top             =   150
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Añadir"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7665
      Left            =   210
      TabIndex        =   83
      Top             =   870
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   13520
      _Version        =   393216
      Tabs            =   13
      TabsPerRow      =   7
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Contabilidad"
      TabPicture(0)   =   "frmConfParamAplic.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Internet"
      TabPicture(1)   =   "frmConfParamAplic.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame17"
      Tab(1).Control(2)=   "Frame21"
      Tab(1).Control(3)=   "Frame7"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Entradas"
      TabPicture(2)   =   "frmConfParamAplic.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(2)=   "Label19"
      Tab(2).Control(3)=   "Label20"
      Tab(2).Control(4)=   "Label21"
      Tab(2).Control(5)=   "Label22"
      Tab(2).Control(6)=   "Label1(101)"
      Tab(2).Control(7)=   "imgAyuda(2)"
      Tab(2).Control(8)=   "Frame3"
      Tab(2).Control(9)=   "chkTaraTractor"
      Tab(2).Control(10)=   "chkTraza"
      Tab(2).Control(11)=   "Text1(24)"
      Tab(2).Control(12)=   "Text1(31)"
      Tab(2).Control(13)=   "chkAgruparNotas"
      Tab(2).Control(14)=   "Text1(64)"
      Tab(2).Control(15)=   "Text1(65)"
      Tab(2).Control(16)=   "Text1(66)"
      Tab(2).Control(17)=   "chkRespetarNroNota"
      Tab(2).Control(18)=   "chkNotaManual"
      Tab(2).Control(19)=   "Text1(109)"
      Tab(2).Control(20)=   "chkCoopro"
      Tab(2).ControlCount=   21
      TabCaption(3)   =   "Aridoc"
      TabPicture(3)   =   "frmConfParamAplic.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(28)"
      Tab(3).Control(1)=   "imgBuscar(9)"
      Tab(3).Control(2)=   "SSTab2"
      Tab(3).Control(3)=   "Text1(13)"
      Tab(3).Control(4)=   "Text2(13)"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Otros"
      TabPicture(4)   =   "frmConfParamAplic.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label12"
      Tab(4).Control(1)=   "Label13"
      Tab(4).Control(2)=   "imgBuscar(0)"
      Tab(4).Control(3)=   "Label1(0)"
      Tab(4).Control(4)=   "Label1(58)"
      Tab(4).Control(5)=   "Label1(4)"
      Tab(4).Control(6)=   "Label15"
      Tab(4).Control(7)=   "imgZoom(0)"
      Tab(4).Control(8)=   "imgZoom(1)"
      Tab(4).Control(9)=   "Label16"
      Tab(4).Control(10)=   "Label17"
      Tab(4).Control(11)=   "Label1(102)"
      Tab(4).Control(12)=   "Label34"
      Tab(4).Control(13)=   "Label1(127)"
      Tab(4).Control(14)=   "imgAyuda(4)"
      Tab(4).Control(15)=   "Text1(25)"
      Tab(4).Control(16)=   "Text1(26)"
      Tab(4).Control(17)=   "Text1(27)"
      Tab(4).Control(18)=   "Text2(27)"
      Tab(4).Control(19)=   "Text1(28)"
      Tab(4).Control(20)=   "Frame5"
      Tab(4).Control(21)=   "Text1(37)"
      Tab(4).Control(22)=   "Text1(38)"
      Tab(4).Control(23)=   "Text1(39)"
      Tab(4).Control(24)=   "Text1(41)"
      Tab(4).Control(25)=   "Text1(110)"
      Tab(4).Control(26)=   "Text1(136)"
      Tab(4).Control(27)=   "Text1(142)"
      Tab(4).ControlCount=   28
      TabCaption(5)   =   "Terc/Trans"
      TabPicture(5)   =   "frmConfParamAplic.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame18"
      Tab(5).Control(1)=   "Frame19"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Almazara"
      TabPicture(6)   =   "frmConfParamAplic.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label1(34)"
      Tab(6).Control(1)=   "imgBuscar(3)"
      Tab(6).Control(2)=   "Label1(107)"
      Tab(6).Control(3)=   "Label1(108)"
      Tab(6).Control(4)=   "imgAyuda(1)"
      Tab(6).Control(5)=   "Frame10"
      Tab(6).Control(6)=   "Text2(48)"
      Tab(6).Control(7)=   "Text1(48)"
      Tab(6).Control(8)=   "Frame16"
      Tab(6).Control(9)=   "Text1(115)"
      Tab(6).Control(10)=   "Text1(116)"
      Tab(6).ControlCount=   11
      TabCaption(7)   =   "ADV"
      TabPicture(7)   =   "frmConfParamAplic.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label1(36)"
      Tab(7).Control(1)=   "imgBuscar(4)"
      Tab(7).Control(2)=   "imgBuscar(5)"
      Tab(7).Control(3)=   "Label1(42)"
      Tab(7).Control(4)=   "Label1(44)"
      Tab(7).Control(5)=   "imgBuscar(58)"
      Tab(7).Control(6)=   "Label1(106)"
      Tab(7).Control(7)=   "imgBuscar(23)"
      Tab(7).Control(8)=   "Text2(56)"
      Tab(7).Control(9)=   "Text1(56)"
      Tab(7).Control(10)=   "Text1(57)"
      Tab(7).Control(11)=   "Text2(57)"
      Tab(7).Control(12)=   "Text2(58)"
      Tab(7).Control(13)=   "Text1(58)"
      Tab(7).Control(14)=   "Text2(114)"
      Tab(7).Control(15)=   "Text1(114)"
      Tab(7).ControlCount=   16
      TabCaption(8)   =   "Suministros"
      TabPicture(8)   =   "frmConfParamAplic.frx":00EC
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label1(46)"
      Tab(8).Control(1)=   "imgBuscar(60)"
      Tab(8).Control(2)=   "Label1(52)"
      Tab(8).Control(3)=   "Text2(60)"
      Tab(8).Control(4)=   "Text1(60)"
      Tab(8).Control(5)=   "Text1(62)"
      Tab(8).ControlCount=   6
      TabCaption(9)   =   "Bodega"
      TabPicture(9)   =   "frmConfParamAplic.frx":0108
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Text1(128)"
      Tab(9).Control(1)=   "Frame15"
      Tab(9).Control(2)=   "Text1(76)"
      Tab(9).Control(3)=   "Text2(76)"
      Tab(9).Control(4)=   "Text1(75)"
      Tab(9).Control(5)=   "Text1(69)"
      Tab(9).Control(6)=   "Text2(69)"
      Tab(9).Control(7)=   "Text1(59)"
      Tab(9).Control(8)=   "Text2(59)"
      Tab(9).Control(9)=   "ChkContadorManual"
      Tab(9).Control(10)=   "Text2(63)"
      Tab(9).Control(11)=   "Text1(63)"
      Tab(9).Control(12)=   "imgAyuda(3)"
      Tab(9).Control(13)=   "Label1(120)"
      Tab(9).Control(14)=   "imgBuscar(16)"
      Tab(9).Control(15)=   "Label1(76)"
      Tab(9).Control(16)=   "Label1(75)"
      Tab(9).Control(17)=   "imgBuscar(69)"
      Tab(9).Control(18)=   "Label1(65)"
      Tab(9).Control(19)=   "imgBuscar(59)"
      Tab(9).Control(20)=   "Label1(45)"
      Tab(9).Control(21)=   "Label1(53)"
      Tab(9).Control(22)=   "imgBuscar(10)"
      Tab(9).ControlCount=   23
      TabCaption(10)  =   "Telefonia"
      TabPicture(10)  =   "frmConfParamAplic.frx":0124
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Text1(71)"
      Tab(10).Control(1)=   "Text2(70)"
      Tab(10).Control(2)=   "Text1(70)"
      Tab(10).Control(3)=   "Label1(67)"
      Tab(10).Control(4)=   "Label1(66)"
      Tab(10).Control(5)=   "imgBuscar(70)"
      Tab(10).ControlCount=   6
      TabCaption(11)  =   "Nóminas"
      TabPicture(11)  =   "frmConfParamAplic.frx":0140
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Label1(68)"
      Tab(11).Control(1)=   "imgBuscar(13)"
      Tab(11).Control(2)=   "Label31"
      Tab(11).Control(3)=   "Label1(94)"
      Tab(11).Control(4)=   "Label1(95)"
      Tab(11).Control(5)=   "Label1(96)"
      Tab(11).Control(6)=   "Label1(97)"
      Tab(11).Control(7)=   "Label1(100)"
      Tab(11).Control(8)=   "Text2(72)"
      Tab(11).Control(9)=   "Text1(72)"
      Tab(11).Control(10)=   "Text1(97)"
      Tab(11).Control(11)=   "Text1(98)"
      Tab(11).Control(12)=   "Text1(99)"
      Tab(11).Control(13)=   "Text1(100)"
      Tab(11).Control(14)=   "Text1(101)"
      Tab(11).Control(15)=   "Text1(108)"
      Tab(11).ControlCount=   16
      TabCaption(12)  =   "Pozos"
      TabPicture(12)  =   "frmConfParamAplic.frx":015C
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Label24"
      Tab(12).Control(1)=   "Label25"
      Tab(12).Control(2)=   "Label27"
      Tab(12).Control(3)=   "imgBuscar(17)"
      Tab(12).Control(4)=   "Label1(85)"
      Tab(12).Control(5)=   "Label28"
      Tab(12).Control(6)=   "Label26"
      Tab(12).Control(7)=   "Label29"
      Tab(12).Control(8)=   "Label1(113)"
      Tab(12).Control(9)=   "imgBuscar(25)"
      Tab(12).Control(10)=   "Label1(114)"
      Tab(12).Control(11)=   "imgBuscar(122)"
      Tab(12).Control(12)=   "imgBuscar(123)"
      Tab(12).Control(13)=   "Label1(115)"
      Tab(12).Control(14)=   "imgBuscar(124)"
      Tab(12).Control(15)=   "Label1(116)"
      Tab(12).Control(16)=   "Label1(118)"
      Tab(12).Control(17)=   "imgBuscar(126)"
      Tab(12).Control(18)=   "Label1(119)"
      Tab(12).Control(19)=   "imgBuscar(127)"
      Tab(12).Control(20)=   "imgBuscar(129)"
      Tab(12).Control(21)=   "Label1(121)"
      Tab(12).Control(22)=   "imgBuscar(130)"
      Tab(12).Control(23)=   "Label1(122)"
      Tab(12).Control(24)=   "Label1(131)"
      Tab(12).Control(25)=   "imgBuscar(131)"
      Tab(12).Control(26)=   "Label1(123)"
      Tab(12).Control(27)=   "Label1(124)"
      Tab(12).Control(28)=   "Label1(125)"
      Tab(12).Control(29)=   "imgBuscar(134)"
      Tab(12).Control(30)=   "imgBuscar(135)"
      Tab(12).Control(31)=   "Label1(126)"
      Tab(12).Control(32)=   "Label8"
      Tab(12).Control(33)=   "Label35"
      Tab(12).Control(34)=   "Label36"
      Tab(12).Control(35)=   "Label37"
      Tab(12).Control(36)=   "Text1(88)"
      Tab(12).Control(37)=   "Text1(86)"
      Tab(12).Control(38)=   "Text1(89)"
      Tab(12).Control(39)=   "Text1(87)"
      Tab(12).Control(40)=   "Text1(90)"
      Tab(12).Control(41)=   "Text2(90)"
      Tab(12).Control(42)=   "Text1(92)"
      Tab(12).Control(43)=   "Text1(91)"
      Tab(12).Control(44)=   "Text2(121)"
      Tab(12).Control(45)=   "Text1(121)"
      Tab(12).Control(46)=   "Text2(122)"
      Tab(12).Control(47)=   "Text1(122)"
      Tab(12).Control(48)=   "Text1(123)"
      Tab(12).Control(49)=   "Text2(123)"
      Tab(12).Control(50)=   "Text1(124)"
      Tab(12).Control(51)=   "Text2(124)"
      Tab(12).Control(52)=   "Text2(126)"
      Tab(12).Control(53)=   "Text1(126)"
      Tab(12).Control(54)=   "Text2(127)"
      Tab(12).Control(55)=   "Text1(127)"
      Tab(12).Control(56)=   "Text1(129)"
      Tab(12).Control(57)=   "Text2(129)"
      Tab(12).Control(58)=   "Text1(130)"
      Tab(12).Control(59)=   "Text2(130)"
      Tab(12).Control(60)=   "Text2(131)"
      Tab(12).Control(61)=   "Text1(131)"
      Tab(12).Control(62)=   "Text1(132)"
      Tab(12).Control(63)=   "Text1(133)"
      Tab(12).Control(64)=   "Text2(134)"
      Tab(12).Control(65)=   "Text1(134)"
      Tab(12).Control(66)=   "Text1(135)"
      Tab(12).Control(67)=   "Text2(135)"
      Tab(12).Control(68)=   "Text1(137)"
      Tab(12).Control(69)=   "Text1(138)"
      Tab(12).Control(70)=   "Text1(139)"
      Tab(12).Control(71)=   "Text1(140)"
      Tab(12).Control(72)=   "Text1(141)"
      Tab(12).ControlCount=   73
      Begin VB.TextBox Text1 
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
         Index           =   142
         Left            =   -71805
         MaxLength       =   10
         TabIndex        =   65
         Tag             =   "Precio Capital Social|N|S|||rparam|eurcapsocial|###,##0.00||"
         Top             =   6660
         Width           =   1020
      End
      Begin VB.TextBox Text1 
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
         Index           =   141
         Left            =   -65955
         MaxLength       =   8
         TabIndex        =   322
         Tag             =   "Coef Consumo Pozos|N|S|||rparam|coefsuministropoz|###,##0.00||"
         Text            =   "cuota"
         Top             =   4050
         Width           =   1515
      End
      Begin VB.TextBox Text1 
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
         Index           =   140
         Left            =   -71940
         MaxLength       =   8
         TabIndex        =   321
         Tag             =   "Coef Consumo Pozos|N|S|||rparam|coefconsumopoz|###,##0.00||"
         Text            =   "cuota"
         Top             =   4050
         Width           =   1425
      End
      Begin VB.TextBox Text1 
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
         Index           =   139
         Left            =   -65955
         MaxLength       =   8
         TabIndex        =   320
         Tag             =   "Canon Contador Pozos|N|S|||rparam|imporcanonpoz|###,##0.00||"
         Text            =   "Canon"
         Top             =   3645
         Width           =   1560
      End
      Begin VB.TextBox Text1 
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
         Index           =   138
         Left            =   -73200
         MaxLength       =   8
         TabIndex        =   311
         Tag             =   "Consumo 3|N|S|||rparam|hastametcub3poz|0000000||"
         Text            =   "m3"
         Top             =   2070
         Width           =   1560
      End
      Begin VB.TextBox Text1 
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
         Index           =   137
         Left            =   -71370
         MaxLength       =   8
         TabIndex        =   312
         Tag             =   "Precio 3|N|S|||rparam|precio3poz|#,##0.00||"
         Text            =   "precio3"
         Top             =   2070
         Width           =   1170
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
         Index           =   136
         Left            =   -71820
         MaxLength       =   250
         TabIndex        =   64
         Tag             =   "Path Impresión Entradas|T|S|||rparam|directorioentradas|||"
         Top             =   6210
         Width           =   7245
      End
      Begin VB.TextBox Text2 
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
         Index           =   135
         Left            =   -70500
         TabIndex        =   406
         Top             =   6750
         Width           =   6120
      End
      Begin VB.TextBox Text1 
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
         Index           =   135
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   230
         Tag             =   "Cta Contable Recargos Pozos|T|S|||rparam|ctarecargospoz|||"
         Top             =   6750
         Width           =   1395
      End
      Begin VB.TextBox Text1 
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
         Index           =   134
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   229
         Tag             =   "Cta Contable Ventas Manta Pozos|T|S|||rparam|ctaventasmantapoz|||"
         Top             =   6390
         Width           =   1395
      End
      Begin VB.TextBox Text2 
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
         Index           =   134
         Left            =   -70500
         TabIndex        =   404
         Top             =   6390
         Width           =   6120
      End
      Begin VB.TextBox Text1 
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
         Index           =   133
         Left            =   -66990
         MaxLength       =   9
         TabIndex        =   314
         Tag             =   "Consumo Máximo Pozos|N|N|||rparam|consumomaxpoz|000000000||"
         Top             =   1590
         Width           =   1425
      End
      Begin VB.TextBox Text1 
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
         Index           =   132
         Left            =   -66990
         MaxLength       =   9
         TabIndex        =   313
         Tag             =   "Consumo Mínimo Pozos|N|N|||rparam|consumominpoz|000000000||"
         Top             =   1200
         Width           =   1425
      End
      Begin VB.TextBox Text1 
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
         Index           =   131
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   323
         Tag             =   "Carta Reclamación Pozos|T|S|||rparam|codcartapoz|||"
         Top             =   4440
         Width           =   1395
      End
      Begin VB.TextBox Text2 
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
         Index           =   131
         Left            =   -70500
         TabIndex        =   400
         Top             =   4440
         Width           =   6120
      End
      Begin VB.TextBox Text2 
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
         Index           =   130
         Left            =   -70500
         TabIndex        =   398
         Top             =   6030
         Width           =   6120
      End
      Begin VB.TextBox Text1 
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
         Index           =   130
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   228
         Tag             =   "Cta Contable Ventas Mto. Pozos|T|S|||rparam|ctaventasmtopoz|||"
         Top             =   6030
         Width           =   1395
      End
      Begin VB.TextBox Text2 
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
         Index           =   129
         Left            =   -70500
         TabIndex        =   396
         Top             =   5640
         Width           =   6120
      End
      Begin VB.TextBox Text1 
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
         Index           =   129
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   227
         Tag             =   "Cta Contable Ventas Talla Pozos|T|S|||rparam|ctaventastalpoz|||"
         Top             =   5640
         Width           =   1395
      End
      Begin VB.TextBox Text1 
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
         Index           =   128
         Left            =   -72510
         MaxLength       =   10
         TabIndex        =   178
         Tag             =   "Porcentaje Incr.kilos entrada|N|S|||rparam|porckilosbod|##0.00||"
         Top             =   3840
         Width           =   585
      End
      Begin VB.TextBox Text1 
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
         Index           =   127
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   317
         Tag             =   "FP Recibo Pozos|N|S|||rparam|forparecpoz|000||"
         Top             =   3225
         Width           =   585
      End
      Begin VB.TextBox Text2 
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
         Index           =   127
         Left            =   -71970
         TabIndex        =   393
         Top             =   3225
         Width           =   6390
      End
      Begin VB.TextBox Text1 
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
         Index           =   126
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   316
         Tag             =   "FP contado Pozos|N|S|||rparam|forpaconpoz|000||"
         Top             =   2835
         Width           =   585
      End
      Begin VB.TextBox Text2 
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
         Index           =   126
         Left            =   -71970
         TabIndex        =   391
         Top             =   2835
         Width           =   6390
      End
      Begin VB.Frame Frame7 
         Caption         =   "Envio E-Mail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1905
         Left            =   -74580
         TabIndex        =   92
         Top             =   870
         Width           =   10140
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
            Index           =   6
            Left            =   1470
            MaxLength       =   50
            TabIndex        =   410
            Tag             =   "Servidor SMTP|T|S|||rparam|smtpHost|||"
            Text            =   "3"
            Top             =   690
            Width           =   8430
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
            Index           =   5
            Left            =   1470
            MaxLength       =   50
            TabIndex        =   409
            Tag             =   "Direccion e-mail|T|S|||rparam|diremail|||"
            Text            =   "3"
            Top             =   300
            Width           =   8430
         End
         Begin VB.CheckBox chkOutlook 
            Caption         =   "Enviar desde Outlook"
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
            Left            =   5310
            TabIndex        =   12
            Tag             =   "Outlook|N|N|||rparam|EnvioDesdeOutlook|||"
            Top             =   1500
            Width           =   4005
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
            IMEMode         =   3  'DISABLE
            Index           =   125
            Left            =   2730
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "LanzaMailOutlook|T|S|||rparam|arigesmail|||"
            Text            =   "3"
            Top             =   1470
            Width           =   1680
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
            Index           =   7
            Left            =   1470
            MaxLength       =   50
            TabIndex        =   9
            Tag             =   "Usuario SMTP|T|S|||rparam|smtpUser|||"
            Text            =   "3"
            Top             =   1080
            Width           =   2940
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
            IMEMode         =   3  'DISABLE
            Index           =   8
            Left            =   6480
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   10
            Tag             =   "Password SMTP|T|S|||rparam|smtpPass|||"
            Text            =   "3"
            Top             =   1080
            Width           =   2220
         End
         Begin VB.Label Label1 
            Caption         =   "Lanza pantalla mail outlook"
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
            Index           =   117
            Left            =   120
            TabIndex        =   390
            Top             =   1500
            Width           =   2040
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
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
            Left            =   120
            TabIndex        =   96
            Top             =   330
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
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
            Left            =   120
            TabIndex        =   95
            Top             =   750
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
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
            Left            =   120
            TabIndex        =   94
            Top             =   1140
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Left            =   5310
            TabIndex        =   93
            Top             =   1110
            Width           =   1440
         End
      End
      Begin VB.TextBox Text2 
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
         Index           =   124
         Left            =   -70500
         TabIndex        =   382
         Top             =   7140
         Width           =   6120
      End
      Begin VB.TextBox Text1 
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
         Index           =   124
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   326
         Tag             =   "Centro Coste Pozos|T|S|||rparam|codccostpoz|||"
         Top             =   7140
         Width           =   1395
      End
      Begin VB.TextBox Text2 
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
         Index           =   123
         Left            =   -70500
         TabIndex        =   380
         Top             =   5250
         Width           =   6120
      End
      Begin VB.TextBox Text1 
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
         Index           =   123
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   325
         Tag             =   "Cta Contable Ventas Cuotas Pozos|T|S|||rparam|ctaventascuopoz|||"
         Top             =   5250
         Width           =   1395
      End
      Begin VB.TextBox Text1 
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
         Index           =   122
         Left            =   -71955
         MaxLength       =   10
         TabIndex        =   324
         Tag             =   "Cta Contable Ventas Consumo Pozos|T|S|||rparam|ctaventasconspoz|||"
         Top             =   4860
         Width           =   1395
      End
      Begin VB.TextBox Text2 
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
         Index           =   122
         Left            =   -70500
         TabIndex        =   378
         Top             =   4860
         Width           =   6120
      End
      Begin VB.TextBox Text1 
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
         Index           =   121
         Left            =   -72030
         MaxLength       =   10
         TabIndex        =   306
         Tag             =   "Sección Pozos|N|S|||rparam|seccionpozos|000||"
         Top             =   720
         Width           =   555
      End
      Begin VB.TextBox Text2 
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
         Index           =   121
         Left            =   -71400
         TabIndex        =   376
         Top             =   720
         Width           =   5820
      End
      Begin VB.Frame Frame21 
         Caption         =   "Envio SMS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1605
         Left            =   -74580
         TabIndex        =   372
         Top             =   2850
         Width           =   10155
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
            Index           =   120
            Left            =   1500
            MaxLength       =   11
            TabIndex        =   15
            Tag             =   "Remitente SMS|T|S|||rparam|smsremitente|||"
            Text            =   "3"
            Top             =   1050
            Width           =   3150
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
            IMEMode         =   3  'DISABLE
            Index           =   119
            Left            =   1500
            MaxLength       =   50
            PasswordChar    =   "*"
            TabIndex        =   14
            Tag             =   "Clave SMS|T|S|||rparam|smsclave|||"
            Text            =   "3"
            Top             =   660
            Width           =   8400
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
            Index           =   118
            Left            =   1500
            MaxLength       =   50
            TabIndex        =   13
            Tag             =   "Direccion e-mail SMS|T|S|||rparam|smsemail|||"
            Text            =   "3"
            Top             =   270
            Width           =   8400
         End
         Begin VB.Label Label1 
            Caption         =   "Remitente"
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
            Index           =   112
            Left            =   150
            TabIndex        =   375
            Top             =   1080
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Clave"
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
            Index           =   111
            Left            =   150
            TabIndex        =   374
            Top             =   690
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
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
            Index           =   110
            Left            =   150
            TabIndex        =   373
            Top             =   300
            Width           =   1380
         End
      End
      Begin VB.CheckBox chkCoopro 
         Caption         =   "Desdoble Coopropietarios "
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
         Left            =   -70920
         TabIndex        =   49
         Tag             =   "Desdoble Coopropietarios Entradas|N|S|||rparam|cooproentradas|0||"
         Top             =   6030
         Width           =   2895
      End
      Begin VB.TextBox Text1 
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
         Index           =   116
         Left            =   -66360
         MaxLength       =   10
         TabIndex        =   365
         Tag             =   "Precio por litro Gto.Envasado|N|S|||rparam|gtoenvasado||#,##0.0000|"
         Top             =   5940
         Width           =   1425
      End
      Begin VB.TextBox Text1 
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
         Index           =   115
         Left            =   -66360
         MaxLength       =   10
         TabIndex        =   364
         Tag             =   "Precio por Kilo Gto.Molturación|N|S|||rparam|gtomoltura||#,##0.0000|"
         Top             =   5520
         Width           =   1425
      End
      Begin VB.TextBox Text1 
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
         Index           =   114
         Left            =   -71970
         MaxLength       =   10
         TabIndex        =   77
         Tag             =   "Cod.Iva Pozos|N|S|||rparam|codivaexeadv|000||"
         Top             =   2820
         Width           =   705
      End
      Begin VB.TextBox Text2 
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
         Index           =   114
         Left            =   -71220
         TabIndex        =   368
         Top             =   2820
         Width           =   6930
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   53
         Tag             =   "Faneca|N|N|||rparam|faneca|0.0000||"
         Top             =   1230
         Width           =   1215
      End
      Begin VB.TextBox Text1 
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
         Left            =   -64950
         MaxLength       =   10
         TabIndex        =   50
         Tag             =   "Porc.Incr/Decr.Aforo|N|N|||rparam|porcincreaforo||##0.00|"
         Top             =   6090
         Width           =   735
      End
      Begin VB.TextBox Text1 
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
         Left            =   -71850
         MaxLength       =   4
         TabIndex        =   272
         Tag             =   "Nro.Maximo Jornadas|N|S|0|1000|rparam|nromaxjornadas||###0|"
         Top             =   3930
         Width           =   855
      End
      Begin VB.CheckBox chkNotaManual 
         Caption         =   "Nro.Nota manual"
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
         Left            =   -74550
         TabIndex        =   44
         Tag             =   "Nro Nota Manual|N|S|||rparam|nronotamanual|0||"
         Top             =   6030
         Width           =   2505
      End
      Begin VB.TextBox Text1 
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
         Index           =   101
         Left            =   -71850
         MaxLength       =   10
         TabIndex        =   268
         Tag             =   "Porcentaje Jornadas|N|S|0|100|rparam|porcjornada||##0.00|"
         Top             =   2100
         Width           =   855
      End
      Begin VB.TextBox Text1 
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
         Index           =   100
         Left            =   -71850
         MaxLength       =   10
         TabIndex        =   271
         Tag             =   "Porcentaje IRPF|N|S|0|100|rparam|porcirpf||##0.00|"
         Top             =   3450
         Width           =   855
      End
      Begin VB.TextBox Text1 
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
         Index           =   99
         Left            =   -71850
         MaxLength       =   10
         TabIndex        =   270
         Tag             =   "Porcentaje Seg.Social 2|N|S|0|100|rparam|porcsegso2||##0.00|"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text1 
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
         Index           =   98
         Left            =   -71850
         MaxLength       =   10
         TabIndex        =   269
         Tag             =   "Porcentaje Seg.Social 1|N|S|0|100|rparam|porcsegso1||##0.00|"
         Top             =   2550
         Width           =   855
      End
      Begin VB.TextBox Text1 
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
         Index           =   97
         Left            =   -72360
         MaxLength       =   8
         TabIndex        =   267
         Tag             =   "Euros Trabajador/dia|N|N|||rparam|eurostrabdia|#,##0.00||"
         Text            =   "cost.h"
         Top             =   1530
         Width           =   1380
      End
      Begin VB.Frame Frame19 
         Caption         =   "Terceros"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1305
         Left            =   -74790
         TabIndex        =   336
         Top             =   990
         Width           =   10335
         Begin VB.TextBox Text2 
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
            Index           =   42
            Left            =   4170
            TabIndex        =   338
            Top             =   750
            Width           =   5820
         End
         Begin VB.TextBox Text1 
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
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   60
            Tag             =   "Cta Contable Retencion|T|S|||rparam|ctaterreten|||"
            Top             =   750
            Width           =   1350
         End
         Begin VB.TextBox Text2 
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
            Index           =   40
            Left            =   3390
            TabIndex        =   337
            Top             =   330
            Width           =   6600
         End
         Begin VB.TextBox Text1 
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
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   59
            Tag             =   "Cod.Iva Extranjero|N|N|||rparam|codivaintracom|000||"
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Retención"
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
            Index           =   13
            Left            =   330
            TabIndex        =   340
            Top             =   810
            Width           =   1980
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   2460
            ToolTipText     =   "Buscar cuenta"
            Top             =   780
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cod.IVA Extranjero"
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
            Index           =   5
            Left            =   330
            TabIndex        =   339
            Top             =   390
            Width           =   1950
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   2460
            ToolTipText     =   "Buscar Iva"
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Transporte"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   4155
         Left            =   -74790
         TabIndex        =   335
         Top             =   2490
         Width           =   10395
         Begin VB.TextBox Text1 
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
            Index           =   93
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   421
            Tag             =   "Tarifa transporte local|N|S|||rparam|tracodtarif|00||"
            Top             =   870
            Width           =   855
         End
         Begin VB.TextBox Text2 
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
            Index           =   93
            Left            =   3690
            TabIndex        =   420
            Top             =   870
            Width           =   6300
         End
         Begin VB.TextBox Text1 
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
            Index           =   94
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   419
            Tag             =   "Concepto Gasto Transporte|N|S|||rparam|tracodgasto|00||"
            Top             =   1770
            Width           =   855
         End
         Begin VB.TextBox Text2 
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
            Index           =   94
            Left            =   3690
            TabIndex        =   418
            Top             =   1770
            Width           =   6300
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
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   417
            Tag             =   "Tipo de Transporte|N|N|||rparam|tratipoportes||N|"
            Top             =   420
            Width           =   2730
         End
         Begin VB.TextBox Text1 
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
            Index           =   102
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   416
            Tag             =   "Porcentaje Retención Transportista|N|S|||rparam|porcretenfactra||##0.00|"
            Top             =   2220
            Width           =   555
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
            Index           =   29
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   415
            Tag             =   "Tipo contador Factura Transporte|N|N|||rparam|tratipocontador||N|"
            Top             =   3120
            Width           =   2730
         End
         Begin VB.TextBox Text2 
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
            Index           =   113
            Left            =   3690
            TabIndex        =   414
            Top             =   1320
            Width           =   6300
         End
         Begin VB.TextBox Text1 
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
            Index           =   113
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   413
            Tag             =   "Tarifa transporte local|N|S|||rparam|tracodtarif2|00||"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox Text2 
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
            Index           =   117
            Left            =   4170
            TabIndex        =   412
            Top             =   3570
            Width           =   5880
         End
         Begin VB.TextBox Text1 
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
            Index           =   117
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   411
            Tag             =   "Cta Contable Retencion|T|S|||rparam|ctatrareten|||"
            Top             =   3570
            Width           =   1350
         End
         Begin VB.TextBox Text1 
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
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   61
            Tag             =   "Precio por Kilo Transportado|N|S|||rparam|preciotra||#,##0.0000|"
            Top             =   2670
            Width           =   1155
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   18
            Left            =   2460
            ToolTipText     =   "Buscar Tarifa"
            Top             =   900
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tarifa 1 Local "
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
            Index           =   86
            Left            =   300
            TabIndex        =   429
            Top             =   870
            Width           =   1830
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   19
            Left            =   2460
            ToolTipText     =   "Buscar Concepto"
            Top             =   1770
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto Gasto"
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
            Index           =   87
            Left            =   300
            TabIndex        =   428
            Top             =   1800
            Width           =   1680
         End
         Begin VB.Label Label1 
            Caption         =   "Se trabaja con "
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
            Index           =   88
            Left            =   300
            TabIndex        =   427
            Top             =   420
            Width           =   1830
         End
         Begin VB.Label Label1 
            Caption         =   "Porcentaje Retención"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   98
            Left            =   300
            TabIndex        =   426
            Top             =   2220
            Width           =   2460
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Contador Factura "
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
            Index           =   99
            Left            =   300
            TabIndex        =   425
            Top             =   3120
            Width           =   2340
         End
         Begin VB.Label Label1 
            Caption         =   "Precio/Kilo Transportado"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   103
            Left            =   300
            TabIndex        =   424
            Top             =   2670
            Width           =   2820
         End
         Begin VB.Label Label1 
            Caption         =   "Tarifa 2 Local "
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
            Index           =   105
            Left            =   300
            TabIndex        =   423
            Top             =   1365
            Width           =   1890
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   22
            Left            =   2460
            ToolTipText     =   "Buscar Tarifa"
            Top             =   1350
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Retención"
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
            Left            =   300
            TabIndex        =   422
            Top             =   3570
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   24
            Left            =   2460
            ToolTipText     =   "Buscar cuenta"
            Top             =   3600
            Width           =   240
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   5700
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   480
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
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
         Index           =   91
         Left            =   -72600
         MaxLength       =   8
         TabIndex        =   318
         Tag             =   "Importe Cuota Pozos|N|S|||rparam|imporcuotapoz|###,##0.00||"
         Text            =   "cuota"
         Top             =   3645
         Width           =   1560
      End
      Begin VB.TextBox Text1 
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
         Index           =   92
         Left            =   -69315
         MaxLength       =   8
         TabIndex        =   319
         Tag             =   "Importe Derrama Pozos|N|S|||rparam|imporderramapoz|###,##0.00||"
         Text            =   "derrama"
         Top             =   3645
         Width           =   1560
      End
      Begin VB.TextBox Text2 
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
         Index           =   90
         Left            =   -71970
         TabIndex        =   330
         Top             =   2475
         Width           =   6390
      End
      Begin VB.TextBox Text1 
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
         Index           =   90
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   315
         Tag             =   "Cod.Iva Pozos|N|S|||rparam|codivapoz|000||"
         Top             =   2475
         Width           =   585
      End
      Begin VB.TextBox Text1 
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
         Index           =   87
         Left            =   -71370
         MaxLength       =   8
         TabIndex        =   308
         Tag             =   "Precio 1|N|S|||rparam|precio1poz|#,##0.00||"
         Text            =   "precio1"
         Top             =   1290
         Width           =   1170
      End
      Begin VB.TextBox Text1 
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
         Index           =   89
         Left            =   -71370
         MaxLength       =   8
         TabIndex        =   310
         Tag             =   "Precio 2|N|S|||rparam|precio2poz|#,##0.00||"
         Text            =   "precio2"
         Top             =   1680
         Width           =   1170
      End
      Begin VB.TextBox Text1 
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
         Index           =   86
         Left            =   -73200
         MaxLength       =   8
         TabIndex        =   307
         Tag             =   "Consumo 1|N|S|||rparam|hastametcub1poz|0000000||"
         Text            =   "m3"
         Top             =   1290
         Width           =   1560
      End
      Begin VB.TextBox Text1 
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
         Index           =   88
         Left            =   -73200
         MaxLength       =   8
         TabIndex        =   309
         Tag             =   "Consumo 2|N|S|||rparam|hastametcub2poz|0000000||"
         Text            =   "m3"
         Top             =   1680
         Width           =   1560
      End
      Begin VB.Frame Frame17 
         Caption         =   "Direcciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1185
         Left            =   -74580
         TabIndex        =   304
         Top             =   5490
         Width           =   10185
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
            Index           =   95
            Left            =   1560
            MaxLength       =   250
            TabIndex        =   18
            Tag             =   "Goolzoom|T|S|||rparam|goolzoom|||"
            Top             =   660
            Width           =   8370
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
            Index           =   85
            Left            =   1560
            MaxLength       =   250
            TabIndex        =   17
            Tag             =   "Sigpac|T|S|||rparam|sigpac|||"
            Top             =   270
            Width           =   8370
         End
         Begin VB.Label Label30 
            Caption         =   "GoolZoom"
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
            Left            =   180
            TabIndex        =   341
            Top             =   660
            Width           =   1035
         End
         Begin VB.Label Label23 
            Caption         =   "Sigpac"
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
            Left            =   180
            TabIndex        =   305
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Última Facturación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -74730
         TabIndex        =   295
         Top             =   5250
         Width           =   4845
         Begin VB.TextBox Text1 
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
            Index           =   84
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   299
            Tag             =   "Ult.Fact.Liquidación almazara|N|S|||rparam|ultfactliqalmz|0000000||"
            Top             =   855
            Width           =   1515
         End
         Begin VB.TextBox Text1 
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
            Index           =   83
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   298
            Tag             =   "Prim.Fact.Liquidación Almazara|N|S|||rparam|primfactliqalmz|0000000||"
            Top             =   855
            Width           =   1515
         End
         Begin VB.TextBox Text1 
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
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   297
            Tag             =   "Prim.Fact.Anticipo Almazara|N|S|||rparam|primfactantalmz|0000000||"
            Top             =   450
            Width           =   1515
         End
         Begin VB.TextBox Text1 
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
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   296
            Tag             =   "Ult.Fact.Anticipo Almazara|N|S|||rparam|ultfactantalmz|0000000||"
            Top             =   450
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Liquidación:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   84
            Left            =   210
            TabIndex        =   303
            Top             =   915
            Width           =   1830
         End
         Begin VB.Label Label1 
            Caption         =   "Anticipos:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   83
            Left            =   210
            TabIndex        =   302
            Top             =   450
            Width           =   1080
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   79
            Left            =   2430
            TabIndex        =   301
            Top             =   180
            Width           =   630
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   77
            Left            =   4050
            TabIndex        =   300
            Top             =   180
            Width           =   630
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Última Facturación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   -74610
         TabIndex        =   286
         Top             =   5010
         Width           =   4845
         Begin VB.TextBox Text1 
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
            Left            =   3420
            MaxLength       =   10
            TabIndex        =   290
            Tag             =   "Ult.Fact.Anticipo Bodega|N|S|||rparam|ultfactantbod|0000000||"
            Top             =   510
            Width           =   1245
         End
         Begin VB.TextBox Text1 
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
            Left            =   2100
            MaxLength       =   10
            TabIndex        =   289
            Tag             =   "Prim.Fact.Anticipo Bodega|N|S|||rparam|primfactantbod|0000000||"
            Top             =   510
            Width           =   1245
         End
         Begin VB.TextBox Text1 
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
            Index           =   79
            Left            =   2100
            MaxLength       =   10
            TabIndex        =   288
            Tag             =   "Prim.Fact.Liquidación Bodega|N|S|||rparam|primfactliqbod|0000000||"
            Top             =   945
            Width           =   1245
         End
         Begin VB.TextBox Text1 
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
            Left            =   3420
            MaxLength       =   10
            TabIndex        =   287
            Tag             =   "Ult.Fact.Liquidación bodega|N|S|||rparam|ultfactliqbod|0000000||"
            Top             =   945
            Width           =   1245
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   82
            Left            =   4050
            TabIndex        =   294
            Top             =   210
            Width           =   630
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   81
            Left            =   2730
            TabIndex        =   293
            Top             =   210
            Width           =   630
         End
         Begin VB.Label Label1 
            Caption         =   "Anticipos:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   80
            Left            =   330
            TabIndex        =   292
            Top             =   450
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Liquidación:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   78
            Left            =   330
            TabIndex        =   291
            Top             =   975
            Width           =   1830
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72525
         MaxLength       =   3
         TabIndex        =   177
         Tag             =   "Codigo Gasto|N|S|||rparam|codgastobod|00||"
         Top             =   3450
         Width           =   585
      End
      Begin VB.TextBox Text2 
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
         Index           =   76
         Left            =   -71895
         TabIndex        =   284
         Top             =   3450
         Width           =   7185
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72525
         MaxLength       =   10
         TabIndex        =   176
         Tag             =   "Porcentaje Gasto Mant|N|N|||rparam|bodporcenmant|##0.00||"
         Top             =   3030
         Width           =   585
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   266
         Tag             =   "Almacen Nominas|N|N|||rparam|codalmacnomi|000||"
         Top             =   1050
         Width           =   855
      End
      Begin VB.TextBox Text2 
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
         Index           =   72
         Left            =   -71415
         TabIndex        =   265
         Top             =   1035
         Width           =   6930
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72570
         MaxLength       =   1
         TabIndex        =   261
         Tag             =   "Letra Serie Almazara|T|S|||rparam|letraserietel|||"
         Top             =   1485
         Width           =   465
      End
      Begin VB.TextBox Text2 
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
         Index           =   70
         Left            =   -71100
         TabIndex        =   262
         Top             =   1035
         Width           =   6660
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72570
         MaxLength       =   10
         TabIndex        =   260
         Tag             =   "Cta Contable Ventas Telefonia|T|S|||rparam|ctaventastel|||"
         Top             =   1035
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72525
         MaxLength       =   10
         TabIndex        =   175
         Tag             =   "Cta Contable Ventas Bodega|T|S|||rparam|ctaventasbod|||"
         Top             =   2565
         Width           =   1350
      End
      Begin VB.TextBox Text2 
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
         Index           =   69
         Left            =   -71115
         TabIndex        =   258
         Top             =   2595
         Width           =   6420
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72525
         MaxLength       =   10
         TabIndex        =   174
         Tag             =   "Cta Contable Banco|T|S|||rparam|ctabancobod|||"
         Top             =   2190
         Width           =   1350
      End
      Begin VB.TextBox Text2 
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
         Index           =   59
         Left            =   -71115
         TabIndex        =   185
         Top             =   2190
         Width           =   6420
      End
      Begin VB.CheckBox ChkContadorManual 
         Caption         =   "Contador de Albarán de Retirada manual "
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
         Left            =   -74595
         TabIndex        =   173
         Tag             =   "Contador albaran Retirada Manual|N|S|||rparam|albretiradabodman|0||"
         Top             =   1620
         Width           =   5325
      End
      Begin VB.CheckBox chkRespetarNroNota 
         Caption         =   "Se respeta Nro.de Nota"
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
         Left            =   -74550
         TabIndex        =   42
         Tag             =   "Se Respeta Nro.Notas|N|N|0|1|rparam|serespetanota|0||"
         Top             =   5415
         Width           =   2775
      End
      Begin VB.TextBox Text1 
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
         Index           =   66
         Left            =   -65670
         MaxLength       =   6
         TabIndex        =   46
         Tag             =   "Peso Caja Llena|N|S|||rparam|pesocajallena|##0.00||"
         Text            =   "KgCajo"
         Top             =   4800
         Width           =   1320
      End
      Begin VB.TextBox Text1 
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
         Left            =   -67080
         MaxLength       =   6
         TabIndex        =   48
         Tag             =   "Kilos Caja Máximo|N|N|||rparam|kiloscajamax|##0.00||"
         Text            =   "kgmax"
         Top             =   5460
         Width           =   1620
      End
      Begin VB.TextBox Text1 
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
         Left            =   -68880
         MaxLength       =   6
         TabIndex        =   47
         Tag             =   "Kilos Caja Mínimo|N|N|||rparam|kiloscajamin|##0.00||"
         Text            =   "kgmin"
         Top             =   5460
         Width           =   1710
      End
      Begin VB.CheckBox chkAgruparNotas 
         Caption         =   "Se agrupan notas"
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
         Left            =   -74550
         TabIndex        =   41
         Tag             =   "Se Agrupan Notas|N|S|||rparam|agruparnotas|0||"
         Top             =   5100
         Width           =   2175
      End
      Begin VB.TextBox Text2 
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
         Left            =   -71940
         TabIndex        =   179
         Top             =   1050
         Width           =   7245
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72570
         MaxLength       =   3
         TabIndex        =   172
         Tag             =   "Sección Bodega|N|S|||rparam|seccionbodega|000||"
         Top             =   1050
         Width           =   585
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
         Index           =   62
         Left            =   -72930
         MaxLength       =   10
         TabIndex        =   169
         Tag             =   "BD Ariges|T|S|||rparam|bdariges|||"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72060
         MaxLength       =   3
         TabIndex        =   168
         Tag             =   "Sección Suministros|N|S|||rparam|seccionsumi|000||"
         Top             =   1020
         Width           =   585
      End
      Begin VB.TextBox Text2 
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
         Index           =   60
         Left            =   -71430
         TabIndex        =   167
         Top             =   1020
         Width           =   5910
      End
      Begin VB.TextBox Text1 
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
         Left            =   -71985
         MaxLength       =   10
         TabIndex        =   76
         Tag             =   "Cta Contable Banco|T|S|||rparam|ctabancoadv|||"
         Top             =   2340
         Width           =   1350
      End
      Begin VB.TextBox Text2 
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
         Index           =   58
         Left            =   -70590
         TabIndex        =   165
         Top             =   2340
         Width           =   6285
      End
      Begin VB.TextBox Text2 
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
         Index           =   57
         Left            =   -71205
         TabIndex        =   163
         Top             =   1470
         Width           =   6900
      End
      Begin VB.TextBox Text1 
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
         Left            =   -71985
         MaxLength       =   10
         TabIndex        =   75
         Tag             =   "Almacen ADV|N|N|||rparam|codalmacadv|000||"
         Top             =   1470
         Width           =   705
      End
      Begin VB.TextBox Text1 
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
         Left            =   -71985
         MaxLength       =   10
         TabIndex        =   74
         Tag             =   "Sección ADV|N|S|||rparam|seccionadv|000||"
         Top             =   1020
         Width           =   705
      End
      Begin VB.TextBox Text2 
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
         Index           =   56
         Left            =   -71205
         TabIndex        =   161
         Top             =   1020
         Width           =   6900
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
         Index           =   41
         Left            =   -71820
         MaxLength       =   250
         TabIndex        =   63
         Tag             =   "Path Traza|T|S|||rparam|directoriotraza|||"
         Top             =   5790
         Width           =   7245
      End
      Begin VB.TextBox Text1 
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
         Left            =   -71730
         MaxLength       =   10
         TabIndex        =   66
         Tag             =   "Sección Almazara|N|S|||rparam|seccionalmaz|000||"
         Top             =   1020
         Width           =   585
      End
      Begin VB.TextBox Text2 
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
         Left            =   -71100
         TabIndex        =   158
         Top             =   1020
         Width           =   6480
      End
      Begin VB.Frame Frame10 
         Caption         =   "Liquidaciones Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3750
         Left            =   -74730
         TabIndex        =   144
         Top             =   1425
         Width           =   10515
         Begin VB.TextBox Text2 
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
            Index           =   112
            Left            =   3660
            TabIndex        =   366
            Top             =   3180
            Width           =   6525
         End
         Begin VB.TextBox Text1 
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
            Index           =   112
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   363
            Tag             =   "Codigo Gasto|N|S|||rparam|codgastoalmz|00||"
            Top             =   3180
            Width           =   585
         End
         Begin VB.TextBox Text2 
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
            Index           =   49
            Left            =   4410
            TabIndex        =   150
            Top             =   1950
            Width           =   5760
         End
         Begin VB.TextBox Text1 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   72
            Tag             =   "Cta Contable Gastos Almazara|T|S|||rparam|ctagastosalmz|||"
            Top             =   2355
            Width           =   1350
         End
         Begin VB.TextBox Text2 
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
            Index           =   50
            Left            =   4410
            TabIndex        =   149
            Top             =   2355
            Width           =   5760
         End
         Begin VB.TextBox Text1 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   71
            Tag             =   "Cta Contable Ventas Almazara|T|S|||rparam|ctaventasalmz|||"
            Top             =   1950
            Width           =   1350
         End
         Begin VB.TextBox Text2 
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
            Index           =   54
            Left            =   4410
            TabIndex        =   148
            Top             =   1575
            Width           =   5760
         End
         Begin VB.TextBox Text1 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   70
            Tag             =   "Cta Contable Banco|T|S|||rparam|ctabancoalmz|||"
            Top             =   1575
            Width           =   1350
         End
         Begin VB.TextBox Text1 
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
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   67
            Tag             =   "Forma Pago Positivas|N|S|||rparam|codforpaposalmz|000||"
            Top             =   360
            Width           =   585
         End
         Begin VB.TextBox Text2 
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
            Index           =   51
            Left            =   3630
            TabIndex        =   147
            Top             =   360
            Width           =   6540
         End
         Begin VB.TextBox Text1 
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
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   68
            Tag             =   "Forma de Pago Negativas|N|S|||rparam|codforpanegalmz|000||"
            Top             =   735
            Width           =   585
         End
         Begin VB.TextBox Text2 
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
            Left            =   3630
            TabIndex        =   146
            Top             =   735
            Width           =   6540
         End
         Begin VB.TextBox Text1 
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
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   69
            Tag             =   "Cta Contable Retencion Socio|T|S|||rparam|ctaretenalmz|||"
            Top             =   1140
            Width           =   1350
         End
         Begin VB.TextBox Text2 
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
            Index           =   53
            Left            =   4410
            TabIndex        =   145
            Top             =   1140
            Width           =   5760
         End
         Begin VB.TextBox Text1 
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
            Left            =   3000
            MaxLength       =   1
            TabIndex        =   73
            Tag             =   "Letra Serie Almazara|T|S|||rparam|letraseriealmz|||"
            Top             =   2760
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "Código Gasto para Liq."
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
            Index           =   104
            Left            =   300
            TabIndex        =   367
            Top             =   3180
            Width           =   2310
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   21
            Left            =   2670
            ToolTipText     =   "Buscar Concepto Gasto"
            Top             =   3210
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Ventas"
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
            Left            =   300
            TabIndex        =   157
            Top             =   2010
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   50
            Left            =   2670
            ToolTipText     =   "Buscar cuenta"
            Top             =   2385
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Gastos"
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
            Index           =   38
            Left            =   300
            TabIndex        =   156
            Top             =   2385
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   49
            Left            =   2670
            ToolTipText     =   "Buscar cuenta"
            Top             =   1980
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Banco Prevista"
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
            Index           =   39
            Left            =   300
            TabIndex        =   155
            Top             =   1605
            Width           =   2310
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   54
            Left            =   2670
            ToolTipText     =   "Buscar cuenta"
            Top             =   1605
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   51
            Left            =   2670
            ToolTipText     =   "Buscar forma Pago"
            Top             =   390
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago Positivas"
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
            Left            =   300
            TabIndex        =   154
            Top             =   420
            Width           =   2340
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   52
            Left            =   2670
            ToolTipText     =   "Buscar forma pago"
            Top             =   765
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago Negativas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   41
            Left            =   300
            TabIndex        =   153
            Top             =   795
            Width           =   2460
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   53
            Left            =   2670
            ToolTipText     =   "Buscar cuenta"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Retención"
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
            Index           =   43
            Left            =   300
            TabIndex        =   152
            Top             =   1200
            Width           =   2340
         End
         Begin VB.Label Label1 
            Caption         =   "Letra Serie Clientes"
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
            Left            =   300
            TabIndex        =   151
            Top             =   2790
            Width           =   2430
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Anticipos / Liquidaciones Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2505
         Left            =   480
         TabIndex        =   133
         Top             =   3840
         Width           =   9855
         Begin VB.TextBox Text2 
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
            Left            =   4230
            TabIndex        =   142
            Top             =   2070
            Width           =   4860
         End
         Begin VB.TextBox Text1 
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
            Left            =   2820
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta Contable Banco|T|S|||rparam|ctabanco|||"
            Top             =   2070
            Width           =   1350
         End
         Begin VB.TextBox Text1 
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
            Left            =   2820
            MaxLength       =   3
            TabIndex        =   4
            Tag             =   "Forma Pago Positivas|N|S|||rparam|codforpaposi|000||"
            Top             =   420
            Width           =   585
         End
         Begin VB.TextBox Text2 
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
            Index           =   43
            Left            =   3450
            TabIndex        =   140
            Top             =   420
            Width           =   5640
         End
         Begin VB.TextBox Text1 
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
            Left            =   2820
            MaxLength       =   3
            TabIndex        =   5
            Tag             =   "Forma de Pago Negativas|N|S|||rparam|codforpanega|000||"
            Top             =   810
            Width           =   585
         End
         Begin VB.TextBox Text2 
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
            Index           =   44
            Left            =   3450
            TabIndex        =   137
            Top             =   810
            Width           =   5640
         End
         Begin VB.TextBox Text1 
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
            Left            =   2820
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Cta Contable Aportacion|T|S|||rparam|ctaaportasoc|||"
            Top             =   1650
            Width           =   1350
         End
         Begin VB.TextBox Text2 
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
            Left            =   4230
            TabIndex        =   136
            Top             =   1650
            Width           =   4860
         End
         Begin VB.TextBox Text1 
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
            Left            =   2820
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta Contable Retencion Socio|T|S|||rparam|ctaretensoc|||"
            Top             =   1230
            Width           =   1350
         End
         Begin VB.TextBox Text2 
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
            Index           =   45
            Left            =   4230
            TabIndex        =   134
            Top             =   1230
            Width           =   4860
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Banco Prevista"
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
            Left            =   300
            TabIndex        =   143
            Top             =   2070
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   270
            Index           =   47
            Left            =   2550
            ToolTipText     =   "Buscar cuenta"
            Top             =   2100
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   270
            Index           =   43
            Left            =   2550
            ToolTipText     =   "Buscar forma Pago"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago Positivas"
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
            Index           =   29
            Left            =   300
            TabIndex        =   141
            Top             =   450
            Width           =   2220
         End
         Begin VB.Image imgBuscar 
            Height          =   270
            Index           =   44
            Left            =   2550
            ToolTipText     =   "Buscar forma pago"
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago Negativas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   32
            Left            =   300
            TabIndex        =   139
            Top             =   840
            Width           =   2670
         End
         Begin VB.Image imgBuscar 
            Height          =   270
            Index           =   46
            Left            =   2550
            ToolTipText     =   "Buscar cuenta"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Aportación"
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
            Index           =   30
            Left            =   300
            TabIndex        =   138
            Top             =   1650
            Width           =   2250
         End
         Begin VB.Image imgBuscar 
            Height          =   270
            Index           =   45
            Left            =   2550
            ToolTipText     =   "Buscar cuenta"
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Retención"
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
            Index           =   24
            Left            =   300
            TabIndex        =   135
            Top             =   1260
            Width           =   2070
         End
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
         Height          =   585
         Index           =   39
         Left            =   -74580
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Tag             =   "Texto Pie Toma Datos|T|S|||rparam|pietomadatos|||"
         Top             =   5040
         Width           =   10005
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
         Height          =   1125
         Index           =   38
         Left            =   -74580
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Tag             =   "Texto Toma Datos|T|S|||rparam|texttomadatos|||"
         Top             =   3600
         Width           =   10005
      End
      Begin VB.TextBox Text1 
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
         Left            =   -71790
         MaxLength       =   10
         TabIndex        =   57
         Tag             =   "Porcentaje AFO|N|S|||rparam|porcenafo|##0.00||"
         Top             =   2850
         Width           =   615
      End
      Begin VB.Frame Frame5 
         Caption         =   "Última Facturación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   -70470
         TabIndex        =   122
         Top             =   1290
         Width           =   5955
         Begin VB.TextBox Text1 
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
            Left            =   4410
            MaxLength       =   10
            TabIndex        =   115
            Tag             =   "Ult.Fact.Liquidación VC|N|S|||rparam|ultfactliqvc|0000000||"
            Top             =   1560
            Width           =   1275
         End
         Begin VB.TextBox Text1 
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
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   114
            Tag             =   "Prim.Fact.Liquidación VC|N|S|||rparam|primfactliqvc|0000000||"
            Top             =   1560
            Width           =   1425
         End
         Begin VB.TextBox Text1 
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
            Left            =   4410
            MaxLength       =   10
            TabIndex        =   113
            Tag             =   "Ult.Fact.Liquidación|N|S|||rparam|ultfactliq|0000000||"
            Top             =   1170
            Width           =   1275
         End
         Begin VB.TextBox Text1 
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
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   112
            Tag             =   "Prim.Fact.Liquidación|N|S|||rparam|primfactliq|0000000||"
            Top             =   1170
            Width           =   1425
         End
         Begin VB.TextBox Text1 
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
            Left            =   4410
            MaxLength       =   10
            TabIndex        =   111
            Tag             =   "Ult.Fact.Anticipo VC|N|S|||rparam|ultfactantvc|0000000||"
            Top             =   780
            Width           =   1275
         End
         Begin VB.TextBox Text1 
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
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   110
            Tag             =   "Prim.Fact.Anticipo VC|N|S|||rparam|primfactantvc|0000000||"
            Top             =   780
            Width           =   1425
         End
         Begin VB.TextBox Text1 
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
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   108
            Tag             =   "Prim.Fact.Anticipo|N|S|||rparam|primfactant|0000000||"
            Top             =   390
            Width           =   1425
         End
         Begin VB.TextBox Text1 
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
            Left            =   4410
            MaxLength       =   10
            TabIndex        =   109
            Tag             =   "Ult.Fact.Anticipo|N|S|||rparam|ultfactant|0000000||"
            Top             =   390
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Liquidación Ventas Campo:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   31
            Left            =   150
            TabIndex        =   128
            Top             =   1560
            Width           =   2700
         End
         Begin VB.Label Label1 
            Caption         =   "Liquidación:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   27
            Left            =   150
            TabIndex        =   127
            Top             =   1170
            Width           =   1830
         End
         Begin VB.Label Label1 
            Caption         =   "Anticipos Ventas Campo:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   150
            TabIndex        =   126
            Top             =   810
            Width           =   2640
         End
         Begin VB.Label Label1 
            Caption         =   "Anticipos:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   125
            Top             =   450
            Width           =   1200
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   2940
            TabIndex        =   124
            Top             =   150
            Width           =   630
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   4410
            TabIndex        =   123
            Top             =   150
            Width           =   990
         End
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
         Index           =   31
         Left            =   -72180
         MaxLength       =   50
         TabIndex        =   39
         Tag             =   "Impresora Entradas|T|N|||rparam|impresoraentradas|||"
         Top             =   4290
         Width           =   7845
      End
      Begin VB.TextBox Text1 
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
         Left            =   -71790
         MaxLength       =   10
         TabIndex        =   56
         Tag             =   "Porcentaje Retención|N|S|||rparam|porcretenfacsoc||##0.00|"
         Top             =   2445
         Width           =   615
      End
      Begin VB.TextBox Text2 
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
         Index           =   27
         Left            =   -71040
         TabIndex        =   118
         Top             =   810
         Width           =   6540
      End
      Begin VB.TextBox Text1 
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
         Left            =   -71820
         MaxLength       =   10
         TabIndex        =   52
         Tag             =   "Sección Hortofrutícola|N|S|||rparam|seccionhorto|000||"
         Top             =   810
         Width           =   675
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72360
         MaxLength       =   8
         TabIndex        =   55
         Tag             =   "Coste seg.soc|N|N|||rparam|costesegso|0.0000||"
         Text            =   "cost.s"
         Top             =   2040
         Width           =   1200
      End
      Begin VB.TextBox Text1 
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
         Left            =   -72360
         MaxLength       =   8
         TabIndex        =   54
         Tag             =   "Coste Horas|N|N|||rparam|costehora|0.0000||"
         Text            =   "cost.h"
         Top             =   1620
         Width           =   1200
      End
      Begin VB.TextBox Text1 
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
         Left            =   -68910
         MaxLength       =   6
         TabIndex        =   45
         Tag             =   "Cajas por Palet|N|N|||rparam|cajasporpalet|###,##0||"
         Text            =   "ncajas"
         Top             =   4800
         Width           =   1200
      End
      Begin VB.CheckBox chkTraza 
         Caption         =   "Hay Trazabilidad"
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
         Left            =   -74550
         TabIndex        =   43
         Tag             =   "Hay Trazabilidad|N|S|||rparam|haytraza|0||"
         Top             =   5730
         Width           =   2145
      End
      Begin VB.CheckBox chkTaraTractor 
         Caption         =   "Se tara tractor de entrada"
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
         Left            =   -74550
         TabIndex        =   40
         Tag             =   "Se Tara Tractor|N|S|||rparam|setaratractor|0||"
         Top             =   4770
         Width           =   3405
      End
      Begin VB.Frame Frame3 
         Height          =   3195
         Left            =   -74820
         TabIndex        =   99
         Top             =   840
         Width           =   10560
         Begin VB.CheckBox ChkVtaFruta 
            Height          =   225
            Index           =   4
            Left            =   9720
            TabIndex        =   389
            Tag             =   "Es VtaFruta 5|N|S|||rparam|esvtafruta5|||"
            Top             =   2490
            Width           =   285
         End
         Begin VB.CheckBox ChkVtaFruta 
            Height          =   225
            Index           =   3
            Left            =   9720
            TabIndex        =   388
            Tag             =   "Es VtaFruta 4|N|S|||rparam|esvtafruta4|||"
            Top             =   2070
            Width           =   285
         End
         Begin VB.CheckBox ChkVtaFruta 
            Height          =   225
            Index           =   2
            Left            =   9720
            TabIndex        =   387
            Tag             =   "Es VtaFruta 3|N|S|||rparam|esvtafruta3|||"
            Top             =   1620
            Width           =   285
         End
         Begin VB.CheckBox ChkVtaFruta 
            Height          =   225
            Index           =   1
            Left            =   9720
            TabIndex        =   386
            Tag             =   "Es VtaFruta 2|N|S|||rparam|esvtafruta2|||"
            Top             =   1170
            Width           =   285
         End
         Begin VB.CheckBox ChkVtaFruta 
            Height          =   225
            Index           =   0
            Left            =   9720
            TabIndex        =   385
            Tag             =   "Es VtaFruta 1|N|S|||rparam|esvtafruta1|||"
            Top             =   720
            Width           =   285
         End
         Begin VB.TextBox Text1 
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
            Index           =   103
            Left            =   6930
            MaxLength       =   6
            TabIndex        =   21
            Tag             =   "Peso Caja 1|N|S|||rparam|pesocaja11|##0.00||"
            Text            =   "peso 1"
            Top             =   675
            Width           =   1410
         End
         Begin VB.TextBox Text1 
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
            Index           =   104
            Left            =   6930
            MaxLength       =   6
            TabIndex        =   25
            Tag             =   "Peso Caja 2|N|S|||rparam|pesocaja12|##0.00||"
            Text            =   "peso 2"
            Top             =   1110
            Width           =   1410
         End
         Begin VB.TextBox Text1 
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
            Index           =   105
            Left            =   6930
            MaxLength       =   6
            TabIndex        =   29
            Tag             =   "Peso Caja 3|N|S|||rparam|pesocaja13|##0.00||"
            Text            =   "peso 3"
            Top             =   1545
            Width           =   1410
         End
         Begin VB.TextBox Text1 
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
            Left            =   6930
            MaxLength       =   6
            TabIndex        =   33
            Tag             =   "Peso Caja 4|N|S|||rparam|pesocaja14|##0.00||"
            Text            =   "peso 4"
            Top             =   1980
            Width           =   1410
         End
         Begin VB.TextBox Text1 
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
            Left            =   6930
            MaxLength       =   6
            TabIndex        =   37
            Tag             =   "Peso Caja 5|N|S|||rparam|pesocaja15|##0.00||"
            Text            =   "peso 5"
            Top             =   2415
            Width           =   1410
         End
         Begin VB.CheckBox ChkCajas 
            Height          =   225
            Index           =   4
            Left            =   8850
            TabIndex        =   38
            Tag             =   "Son Cajas 5|N|S|||rparam|escaja5|||"
            Top             =   2490
            Width           =   285
         End
         Begin VB.CheckBox ChkCajas 
            Height          =   225
            Index           =   3
            Left            =   8850
            TabIndex        =   34
            Tag             =   "Son Cajas 4|N|S|||rparam|escaja4|||"
            Top             =   2070
            Width           =   285
         End
         Begin VB.CheckBox ChkCajas 
            Height          =   225
            Index           =   2
            Left            =   8850
            TabIndex        =   30
            Tag             =   "Son Cajas 3|N|S|||rparam|escaja3|||"
            Top             =   1620
            Width           =   285
         End
         Begin VB.CheckBox ChkCajas 
            Height          =   225
            Index           =   1
            Left            =   8850
            TabIndex        =   26
            Tag             =   "Son Cajas 2|N|S|||rparam|escaja2|||"
            Top             =   1170
            Width           =   285
         End
         Begin VB.CheckBox ChkCajas 
            Height          =   225
            Index           =   0
            Left            =   8850
            TabIndex        =   22
            Tag             =   "Son Cajas 1|N|S|||rparam|escaja1|||"
            Top             =   720
            Width           =   285
         End
         Begin VB.TextBox Text1 
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
            Left            =   5400
            MaxLength       =   6
            TabIndex        =   36
            Tag             =   "Peso Caja 5|N|S|||rparam|pesocaja5|##0.00||"
            Text            =   "peso 5"
            Top             =   2415
            Width           =   1380
         End
         Begin VB.TextBox Text1 
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
            Left            =   5400
            MaxLength       =   6
            TabIndex        =   32
            Tag             =   "Peso Caja 4|N|S|||rparam|pesocaja4|##0.00||"
            Text            =   "peso 4"
            Top             =   1980
            Width           =   1380
         End
         Begin VB.TextBox Text1 
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
            Left            =   5400
            MaxLength       =   6
            TabIndex        =   28
            Tag             =   "Peso Caja 3|N|S|||rparam|pesocaja3|##0.00||"
            Text            =   "peso 3"
            Top             =   1545
            Width           =   1380
         End
         Begin VB.TextBox Text1 
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
            Left            =   5400
            MaxLength       =   6
            TabIndex        =   24
            Tag             =   "Peso Caja 2|N|S|||rparam|pesocaja2|##0.00||"
            Text            =   "peso 2"
            Top             =   1110
            Width           =   1380
         End
         Begin VB.TextBox Text1 
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
            Left            =   5400
            MaxLength       =   6
            TabIndex        =   20
            Tag             =   "Peso Caja 1|N|S|||rparam|pesocaja1|##0.00||"
            Text            =   "peso 1"
            Top             =   675
            Width           =   1380
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
            Index           =   18
            Left            =   420
            MaxLength       =   20
            TabIndex        =   35
            Tag             =   "Tipo Caja 5|T|S|||rparam|tipocaja5|||"
            Text            =   "tipo 5"
            Top             =   2415
            Width           =   4695
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
            Index           =   17
            Left            =   420
            MaxLength       =   20
            TabIndex        =   31
            Tag             =   "Tipo Caja 4|T|S|||rparam|tipocaja4|||"
            Text            =   "tipo 4"
            Top             =   1980
            Width           =   4695
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
            Index           =   16
            Left            =   420
            MaxLength       =   20
            TabIndex        =   27
            Tag             =   "Tipo Caja 3|T|S|||rparam|tipocaja3|||"
            Text            =   "tipo 3"
            Top             =   1545
            Width           =   4695
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
            Index           =   15
            Left            =   420
            MaxLength       =   20
            TabIndex        =   23
            Tag             =   "Tipo Caja 2|T|S|||rparam|tipocaja2|||"
            Text            =   "tipo 2"
            Top             =   1110
            Width           =   4695
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
            Index           =   14
            Left            =   420
            MaxLength       =   20
            TabIndex        =   19
            Tag             =   "Tipo Caja 1|T|S|||rparam|tipocaja1|||"
            Text            =   "tipo 1"
            Top             =   675
            Width           =   4695
         End
         Begin VB.Label Label33 
            Caption         =   "VtaFr."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9570
            TabIndex        =   384
            Top             =   330
            Width           =   870
         End
         Begin VB.Label Label32 
            Caption         =   "PesoTransp."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6990
            TabIndex        =   359
            Top             =   315
            Width           =   1350
         End
         Begin VB.Label Label18 
            Caption         =   "Son Cajas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8400
            TabIndex        =   132
            Top             =   330
            Width           =   1260
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Caja "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   420
            TabIndex        =   106
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "5.-"
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
            Left            =   105
            TabIndex        =   105
            Top             =   2415
            Width           =   240
         End
         Begin VB.Label Label9 
            Caption         =   "4.-"
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
            Left            =   105
            TabIndex        =   104
            Top             =   1980
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "3.-"
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
            Left            =   105
            TabIndex        =   103
            Top             =   1545
            Width           =   285
         End
         Begin VB.Label Label6 
            Caption         =   "2.-"
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
            Left            =   105
            TabIndex        =   102
            Top             =   1110
            Width           =   285
         End
         Begin VB.Label Label5 
            Caption         =   "1.-"
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
            Left            =   105
            TabIndex        =   101
            Top             =   675
            Width           =   285
         End
         Begin VB.Label Label4 
            Caption         =   "Peso de Caja"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5400
            TabIndex        =   100
            Top             =   315
            Width           =   1590
         End
      End
      Begin VB.TextBox Text2 
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
         Left            =   -70680
         TabIndex        =   98
         Top             =   870
         Width           =   5340
      End
      Begin VB.TextBox Text1 
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
         Left            =   -71970
         MaxLength       =   10
         TabIndex        =   51
         Tag             =   "Extension|N|N|||rparam|codextension|000||"
         Top             =   870
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Soporte"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   885
         Left            =   -74580
         TabIndex        =   90
         Top             =   4530
         Width           =   10185
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
            Index           =   9
            Left            =   1530
            MaxLength       =   100
            TabIndex        =   16
            Tag             =   "Web Soporte|T|S|||rparam|websoporte|||"
            Top             =   300
            Width           =   8370
         End
         Begin VB.Label Label2 
            Caption         =   "Web soporte"
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
            Left            =   180
            TabIndex        =   91
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2025
         Left            =   480
         TabIndex        =   84
         Top             =   1230
         Width           =   9825
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
            Index           =   1
            Left            =   2235
            MaxLength       =   20
            TabIndex        =   0
            Tag             =   "Servidor Contabilidad|T|S|||rparam|serconta|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   210
            Width           =   6945
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   4230
            MaxLength       =   15
            TabIndex        =   89
            Tag             =   "Código Parámetros Aplic|N|N|||sparam|codparam||S|"
            Text            =   "1"
            Top             =   240
            Width           =   645
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
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   2235
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   2
            Tag             =   "Password Contabilidad|T|S|||rparam|pasconta|||"
            Text            =   "3"
            Top             =   990
            Width           =   6945
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
            Index           =   2
            Left            =   2235
            MaxLength       =   20
            TabIndex        =   1
            Tag             =   "Usuario Contabilidad|T|S|||rparam|usuconta|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   585
            Width           =   6945
         End
         Begin VB.TextBox Text1 
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
            Left            =   2235
            MaxLength       =   2
            TabIndex        =   3
            Tag             =   "Nº Contabilidad|N|S|||rparam|numconta|||"
            Text            =   "3"
            Top             =   1395
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Left            =   300
            TabIndex        =   88
            Top             =   1050
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   17
            Left            =   300
            TabIndex        =   87
            Top             =   660
            Width           =   1710
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
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
            Left            =   300
            TabIndex        =   86
            Top             =   1470
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
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
            Index           =   19
            Left            =   300
            TabIndex        =   85
            Top             =   240
            Width           =   1560
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4725
         Left            =   -74730
         TabIndex        =   187
         Top             =   1410
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   8334
         _Version        =   393216
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Anticipos / Liquid."
         TabPicture(0)   =   "frmConfParamAplic.frx":0178
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame9"
         Tab(0).Control(1)=   "Frame8"
         Tab(0).Control(2)=   "Text1(11)"
         Tab(0).Control(3)=   "Text2(11)"
         Tab(0).Control(4)=   "Text1(10)"
         Tab(0).Control(5)=   "Text2(10)"
         Tab(0).Control(6)=   "imgBuscar(7)"
         Tab(0).Control(7)=   "Label1(7)"
         Tab(0).Control(8)=   "imgBuscar(6)"
         Tab(0).Control(9)=   "Label1(6)"
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "ADV / Recibos Campo"
         TabPicture(1)   =   "frmConfParamAplic.frx":0194
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame14"
         Tab(1).Control(1)=   "Text1(74)"
         Tab(1).Control(2)=   "Text2(73)"
         Tab(1).Control(3)=   "Text2(74)"
         Tab(1).Control(4)=   "Text1(73)"
         Tab(1).Control(5)=   "Frame11"
         Tab(1).Control(6)=   "Text2(61)"
         Tab(1).Control(7)=   "Text1(61)"
         Tab(1).Control(8)=   "imgBuscar(15)"
         Tab(1).Control(9)=   "Label1(70)"
         Tab(1).Control(10)=   "Label1(69)"
         Tab(1).Control(11)=   "imgBuscar(14)"
         Tab(1).Control(12)=   "Label1(47)"
         Tab(1).Control(13)=   "imgBuscar(8)"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   "Almazara / Bodega"
         TabPicture(2)   =   "frmConfParamAplic.frx":01B0
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label1(63)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "imgBuscar(11)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label1(64)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "imgBuscar(12)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Text2(67)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Text1(67)"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Text2(68)"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "Text1(68)"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "Frame12"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "Frame13"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).ControlCount=   10
         TabCaption(3)   =   "Transporte"
         TabPicture(3)   =   "frmConfParamAplic.frx":01CC
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "imgBuscar(20)"
         Tab(3).Control(1)=   "Label1(93)"
         Tab(3).Control(2)=   "Frame20"
         Tab(3).Control(3)=   "Text1(96)"
         Tab(3).Control(4)=   "Text2(96)"
         Tab(3).ControlCount=   5
         Begin VB.TextBox Text2 
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
            Index           =   96
            Left            =   -70965
            TabIndex        =   352
            Top             =   870
            Width           =   5370
         End
         Begin VB.TextBox Text1 
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
            Index           =   96
            Left            =   -72240
            MaxLength       =   10
            TabIndex        =   347
            Tag             =   "Carpeta Facturas|N|N|||rparam|codcarpetatran|000||"
            Top             =   870
            Width           =   1215
         End
         Begin VB.Frame Frame20 
            Caption         =   "Transporte"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   1170
            Left            =   -74760
            TabIndex        =   342
            Top             =   1440
            Width           =   9240
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
               Index           =   25
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   348
               Tag             =   "C1 Liquidacion|N|N|||rparam|c1tranaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   26
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   349
               Tag             =   "C2 Liquidación|N|N|||rparam|c2tranaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   27
               Left            =   4155
               Style           =   2  'Dropdown List
               TabIndex        =   350
               Tag             =   "C3 Liquidación|N|N|||rparam|c3tranaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   28
               Left            =   6165
               Style           =   2  'Dropdown List
               TabIndex        =   351
               Tag             =   "C4 Liquidación|N|N|||rparam|c4tranaridoc||N|"
               Top             =   585
               Width           =   1800
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 4"
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
               Index           =   92
               Left            =   6165
               TabIndex        =   346
               Top             =   315
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 3"
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
               Index           =   91
               Left            =   4155
               TabIndex        =   345
               Top             =   315
               Width           =   1620
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 2"
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
               Index           =   90
               Left            =   2100
               TabIndex        =   344
               Top             =   315
               Width           =   1755
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 1"
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
               Index           =   89
               Left            =   90
               TabIndex        =   343
               Top             =   315
               Width           =   1620
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Recibos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   1050
            Left            =   -74685
            TabIndex        =   278
            Top             =   3150
            Width           =   9240
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
               Index           =   24
               Left            =   6165
               Style           =   2  'Dropdown List
               TabIndex        =   198
               Tag             =   "C4 Recibo|N|N|||sparam|c4recaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   23
               Left            =   4155
               Style           =   2  'Dropdown List
               TabIndex        =   197
               Tag             =   "C3 Recibo|N|N|||sparam|c3recaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   22
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   196
               Tag             =   "C2 Recibo|N|N|||sparam|c2recaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   21
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   195
               Tag             =   "C1 Recibo|N|N|||sparam|c1recaridoc||N|"
               Top             =   585
               Width           =   1800
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 1"
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
               Index           =   74
               Left            =   90
               TabIndex        =   282
               Top             =   315
               Width           =   1620
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 2"
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
               Index           =   73
               Left            =   2100
               TabIndex        =   281
               Top             =   315
               Width           =   1755
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 3"
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
               Index           =   72
               Left            =   4155
               TabIndex        =   280
               Top             =   315
               Width           =   1620
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 4"
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
               Index           =   71
               Left            =   6165
               TabIndex        =   279
               Top             =   315
               Width           =   1305
            End
         End
         Begin VB.TextBox Text1 
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
            Left            =   -71925
            MaxLength       =   10
            TabIndex        =   190
            Tag             =   "Carpeta Recibos Almacen|N|N|||sparam|codcarpetareccamp|000||"
            Top             =   1410
            Width           =   1110
         End
         Begin VB.TextBox Text2 
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
            Left            =   -70740
            TabIndex        =   275
            Top             =   1020
            Width           =   5370
         End
         Begin VB.TextBox Text2 
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
            Left            =   -70740
            TabIndex        =   274
            Top             =   1410
            Width           =   5370
         End
         Begin VB.TextBox Text1 
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
            Left            =   -71925
            MaxLength       =   10
            TabIndex        =   189
            Tag             =   "Carpeta Recibos Campo|N|N|||sparam|codcarpetarecalm|000||"
            Top             =   1020
            Width           =   1110
         End
         Begin VB.Frame Frame13 
            Caption         =   "Bodega"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   1050
            Left            =   315
            TabIndex        =   251
            Top             =   3135
            Width           =   9240
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
               Index           =   17
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   241
               Tag             =   "C1 Liquidacion|N|N|||rparam|c1bodearidoc||N|"
               Top             =   570
               Width           =   1800
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
               Index           =   18
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   242
               Tag             =   "C2 Liquidación|N|N|||rparam|c2bodearidoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   19
               Left            =   4155
               Style           =   2  'Dropdown List
               TabIndex        =   243
               Tag             =   "C3 Liquidación|N|N|||rparam|c3bodearidoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   20
               Left            =   6165
               Style           =   2  'Dropdown List
               TabIndex        =   244
               Tag             =   "C4 Liquidación|N|N|||rparam|c4bodearidoc||N|"
               Top             =   585
               Width           =   1800
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 4"
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
               Index           =   62
               Left            =   6165
               TabIndex        =   255
               Top             =   315
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 3"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   61
               Left            =   4155
               TabIndex        =   254
               Top             =   315
               Width           =   1620
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 2"
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
               Index           =   60
               Left            =   2100
               TabIndex        =   253
               Top             =   315
               Width           =   1755
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 1"
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
               Index           =   59
               Left            =   90
               TabIndex        =   252
               Top             =   315
               Width           =   1620
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Almazara"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   1050
            Left            =   330
            TabIndex        =   246
            Top             =   1920
            Width           =   9240
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
               Index           =   14
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   238
               Tag             =   "C2 Anticipo|N|N|||rparam|c2almzaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   13
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   237
               Tag             =   "C1 Almazara|N|N|||rparam|c1almzaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   15
               Left            =   4155
               Style           =   2  'Dropdown List
               TabIndex        =   239
               Tag             =   "C3 Anticipo|N|N|||rparam|c3almzaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   16
               Left            =   6165
               Style           =   2  'Dropdown List
               TabIndex        =   240
               Tag             =   "C4 Anticipo|N|N|||rparam|c4almzaridoc||N|"
               Top             =   585
               Width           =   1800
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 1"
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
               Index           =   57
               Left            =   90
               TabIndex        =   250
               Top             =   315
               Width           =   1665
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 2"
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
               Index           =   56
               Left            =   2100
               TabIndex        =   249
               Top             =   315
               Width           =   1755
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 3"
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
               Index           =   55
               Left            =   4155
               TabIndex        =   248
               Top             =   315
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 4"
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
               Index           =   54
               Left            =   6165
               TabIndex        =   247
               Top             =   315
               Width           =   1305
            End
         End
         Begin VB.TextBox Text1 
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
            Left            =   2595
            MaxLength       =   10
            TabIndex        =   236
            Tag             =   "Carpeta Bodega|N|N|||rparam|codcarpetabode|000||"
            Top             =   1170
            Width           =   1215
         End
         Begin VB.TextBox Text2 
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
            Index           =   68
            Left            =   3870
            TabIndex        =   245
            Top             =   1170
            Width           =   5370
         End
         Begin VB.TextBox Text1 
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
            Left            =   2595
            MaxLength       =   10
            TabIndex        =   235
            Tag             =   "Carpeta Almazara|N|N|||rparam|codcarpetaalmz|000||"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text2 
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
            Index           =   67
            Left            =   3870
            TabIndex        =   234
            Top             =   720
            Width           =   5370
         End
         Begin VB.Frame Frame9 
            Caption         =   "Liquidaciones"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   1050
            Left            =   -74685
            TabIndex        =   222
            Top             =   2835
            Width           =   9240
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
               Index           =   8
               Left            =   6165
               Style           =   2  'Dropdown List
               TabIndex        =   215
               Tag             =   "C4 Liquidación|N|N|||rparam|c4liquaridoc||N|"
               Top             =   585
               Width           =   1710
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
               Index           =   7
               Left            =   4155
               Style           =   2  'Dropdown List
               TabIndex        =   214
               Tag             =   "C3 Liquidación|N|N|||rparam|c3liquaridoc||N|"
               Top             =   585
               Width           =   1710
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
               Index           =   6
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   213
               Tag             =   "C2 Liquidación|N|N|||rparam|c2liquaridoc||N|"
               Top             =   585
               Width           =   1710
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
               Index           =   5
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   212
               Tag             =   "C1 Liquidacion|N|N|||rparam|c1liquaridoc||N|"
               Top             =   585
               Width           =   1710
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 1"
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
               Index           =   26
               Left            =   90
               TabIndex        =   226
               Top             =   315
               Width           =   1620
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 2"
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
               Index           =   25
               Left            =   2100
               TabIndex        =   225
               Top             =   315
               Width           =   1755
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 3"
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
               Index           =   16
               Left            =   4155
               TabIndex        =   224
               Top             =   315
               Width           =   1620
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 4"
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
               Index           =   14
               Left            =   6165
               TabIndex        =   223
               Top             =   315
               Width           =   1305
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Anticipos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   1050
            Left            =   -74685
            TabIndex        =   217
            Top             =   1575
            Width           =   9240
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
               Index           =   4
               Left            =   6165
               Style           =   2  'Dropdown List
               TabIndex        =   211
               Tag             =   "C4 Anticipo|N|N|||rparam|c4antiaridoc||N|"
               Top             =   585
               Width           =   1710
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
               Index           =   3
               Left            =   4155
               Style           =   2  'Dropdown List
               TabIndex        =   210
               Tag             =   "C3 Anticipo|N|N|||rparam|c3antiaridoc||N|"
               Top             =   585
               Width           =   1710
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
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   208
               Tag             =   "C1 Anticipo|N|N|||rparam|c1antiaridoc||N|"
               Top             =   585
               Width           =   1710
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
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   209
               Tag             =   "C2 Anticipo|N|N|||rparam|c2antiaridoc||N|"
               Top             =   585
               Width           =   1710
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 4"
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
               Index           =   12
               Left            =   6165
               TabIndex        =   221
               Top             =   315
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 3"
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
               Index           =   10
               Left            =   4155
               TabIndex        =   220
               Top             =   315
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 2"
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
               Index           =   8
               Left            =   2100
               TabIndex        =   219
               Top             =   315
               Width           =   1755
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 1"
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
               Index           =   9
               Left            =   90
               TabIndex        =   218
               Top             =   315
               Width           =   1665
            End
         End
         Begin VB.TextBox Text1 
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
            Left            =   -72225
            MaxLength       =   10
            TabIndex        =   207
            Tag             =   "Carpeta Facturas|N|N|||rparam|codcarpetaliqu|000||"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox Text2 
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
            Index           =   11
            Left            =   -70950
            TabIndex        =   216
            Top             =   1080
            Width           =   5340
         End
         Begin VB.TextBox Text1 
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
            Left            =   -72225
            MaxLength       =   10
            TabIndex        =   206
            Tag             =   "Carpeta Albaranes|N|N|||rparam|codcarpetaanti|000||"
            Top             =   630
            Width           =   1215
         End
         Begin VB.TextBox Text2 
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
            Index           =   10
            Left            =   -70950
            TabIndex        =   205
            Top             =   630
            Width           =   5340
         End
         Begin VB.Frame Frame11 
            Caption         =   "ADV"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   1050
            Left            =   -74685
            TabIndex        =   200
            Top             =   1980
            Width           =   9240
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
               Index           =   12
               Left            =   6165
               Style           =   2  'Dropdown List
               TabIndex        =   194
               Tag             =   "C4 ADV|N|N|||rparam|c4advaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   11
               Left            =   4155
               Style           =   2  'Dropdown List
               TabIndex        =   193
               Tag             =   "C3 ADV|N|N|||rparam|c3advaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   10
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   192
               Tag             =   "C2 ADV|N|N|||rparam|c2advaridoc||N|"
               Top             =   585
               Width           =   1800
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
               Index           =   9
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   191
               Tag             =   "C1 ADV|N|N|||rparam|c1advaridoc||N|"
               Top             =   600
               Width           =   1800
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 1"
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
               Index           =   48
               Left            =   90
               TabIndex        =   204
               Top             =   315
               Width           =   1620
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 2"
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
               Index           =   49
               Left            =   2070
               TabIndex        =   203
               Top             =   315
               Width           =   1755
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 3"
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
               Index           =   50
               Left            =   4185
               TabIndex        =   202
               Top             =   315
               Width           =   1620
            End
            Begin VB.Label Label1 
               Caption         =   "Campo 4"
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
               Index           =   51
               Left            =   6165
               TabIndex        =   201
               Top             =   315
               Width           =   1305
            End
         End
         Begin VB.TextBox Text2 
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
            Index           =   61
            Left            =   -70740
            TabIndex        =   199
            Top             =   630
            Width           =   5370
         End
         Begin VB.TextBox Text1 
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
            Left            =   -71925
            MaxLength       =   10
            TabIndex        =   188
            Tag             =   "Carpeta Facturas|N|N|||rparam|codcarpetaADV|000||"
            Top             =   630
            Width           =   1110
         End
         Begin VB.Label Label1 
            Caption         =   "Carpeta Transporte"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   93
            Left            =   -74670
            TabIndex        =   353
            Top             =   900
            Width           =   2070
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   20
            Left            =   -72570
            ToolTipText     =   "Buscar Carpeta"
            Top             =   870
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   -72195
            ToolTipText     =   "Buscar Carpeta"
            Top             =   1410
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Carp.Recibos Almacén"
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
            Index           =   70
            Left            =   -74640
            TabIndex        =   277
            Top             =   1065
            Width           =   2340
         End
         Begin VB.Label Label1 
            Caption         =   "Carpeta Recibos Campo"
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
            Index           =   69
            Left            =   -74640
            TabIndex        =   276
            Top             =   1455
            Width           =   2730
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   14
            Left            =   -72195
            ToolTipText     =   "Buscar Carpeta"
            Top             =   1020
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   2265
            ToolTipText     =   "Buscar Carpeta"
            Top             =   1230
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Carpeta Bodega"
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
            Index           =   64
            Left            =   420
            TabIndex        =   257
            Top             =   1215
            Width           =   1770
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   2280
            ToolTipText     =   "Buscar Carpeta"
            Top             =   765
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Carpeta Almazara"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   63
            Left            =   420
            TabIndex        =   256
            Top             =   750
            Width           =   1980
         End
         Begin VB.Label Label1 
            Caption         =   "Carpeta ADV"
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
            Index           =   47
            Left            =   -74640
            TabIndex        =   233
            Top             =   675
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   -72555
            ToolTipText     =   "Buscar Carpeta"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Carpeta Liquidacion"
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
            Index           =   7
            Left            =   -74580
            TabIndex        =   232
            Top             =   1125
            Width           =   2130
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   -72555
            ToolTipText     =   "Buscar Carpeta"
            Top             =   675
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Carpeta Anticipos"
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
            Index           =   6
            Left            =   -74580
            TabIndex        =   231
            Top             =   720
            Width           =   3390
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   -72195
            ToolTipText     =   "Buscar Carpeta"
            Top             =   630
            Width           =   240
         End
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   4
         Left            =   -70680
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   6705
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Euros * Sup.Cultivable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   127
         Left            =   -74550
         TabIndex        =   438
         Top             =   6660
         Width           =   2580
      End
      Begin VB.Label Label37 
         Caption         =   "Coeficiente Suministro Externo "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69285
         TabIndex        =   437
         Top             =   4080
         Width           =   3435
      End
      Begin VB.Label Label36 
         Caption         =   "Coeficiente Consumo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74325
         TabIndex        =   436
         Top             =   4080
         Width           =   2265
      End
      Begin VB.Label Label35 
         Caption         =   "Canon Contador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67620
         TabIndex        =   435
         Top             =   3705
         Width           =   2040
      End
      Begin VB.Label Label8 
         Caption         =   "Rango 3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74340
         TabIndex        =   434
         Top             =   2100
         Width           =   1005
      End
      Begin VB.Label Label34 
         Caption         =   "Path Impresión Entradas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74550
         TabIndex        =   408
         Top             =   6210
         Width           =   2820
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Recargos"
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
         Index           =   126
         Left            =   -74340
         TabIndex        =   407
         Top             =   6780
         Width           =   2040
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   135
         Left            =   -72240
         ToolTipText     =   "Buscar cuenta"
         Top             =   6780
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   134
         Left            =   -72240
         ToolTipText     =   "Buscar cuenta"
         Top             =   6420
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Ventas Manta"
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
         Index           =   125
         Left            =   -74340
         TabIndex        =   405
         Top             =   6420
         Width           =   2070
      End
      Begin VB.Label Label1 
         Caption         =   "Consumo Máximo"
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
         Index           =   124
         Left            =   -68880
         TabIndex        =   403
         Top             =   1665
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "Consumo Mínimo"
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
         Index           =   123
         Left            =   -68880
         TabIndex        =   402
         Top             =   1245
         Width           =   1890
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   131
         Left            =   -72240
         ToolTipText     =   "Buscar carta"
         Top             =   4470
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Carta Reclamación "
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
         Index           =   131
         Left            =   -74340
         TabIndex        =   401
         Top             =   4470
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Ventas Mto."
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
         Index           =   122
         Left            =   -74340
         TabIndex        =   399
         Top             =   6060
         Width           =   2070
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   130
         Left            =   -72240
         ToolTipText     =   "Buscar cuenta"
         Top             =   6060
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Ventas Talla"
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
         Index           =   121
         Left            =   -74340
         TabIndex        =   397
         Top             =   5670
         Width           =   2040
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   129
         Left            =   -72240
         ToolTipText     =   "Buscar cuenta"
         Top             =   5670
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   3
         Left            =   -71850
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   3870
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Porc.Increm.Kilos Entrada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   120
         Left            =   -74490
         TabIndex        =   395
         Top             =   3870
         Width           =   2040
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   127
         Left            =   -72870
         ToolTipText     =   "Buscar forma pago"
         Top             =   3225
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "FP Recibo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   119
         Left            =   -74340
         TabIndex        =   394
         Top             =   3255
         Width           =   1350
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   126
         Left            =   -72870
         ToolTipText     =   "Buscar forma pago"
         Top             =   2865
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "FP Contado"
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
         Index           =   118
         Left            =   -74340
         TabIndex        =   392
         Top             =   2895
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Centro de Coste"
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
         Index           =   116
         Left            =   -74340
         TabIndex        =   383
         Top             =   7170
         Width           =   2100
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   124
         Left            =   -72240
         ToolTipText     =   "Buscar centro coste"
         Top             =   7170
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Ventas Cuota"
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
         Index           =   115
         Left            =   -74340
         TabIndex        =   381
         Top             =   5280
         Width           =   2070
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   123
         Left            =   -72240
         ToolTipText     =   "Buscar cuenta"
         Top             =   5280
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   122
         Left            =   -72240
         ToolTipText     =   "Buscar cuenta"
         Top             =   4890
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Ventas Consumo"
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
         Index           =   114
         Left            =   -74340
         TabIndex        =   379
         Top             =   4890
         Width           =   2070
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   -72300
         ToolTipText     =   "Buscar Sección"
         Top             =   750
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sección de Pozos"
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
         Index           =   113
         Left            =   -74370
         TabIndex        =   377
         Top             =   750
         Width           =   1800
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   2
         Left            =   -68010
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   6090
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   1
         Left            =   -64830
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   5550
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Precio por litro Gtos.Envasado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   108
         Left            =   -69780
         TabIndex        =   371
         Top             =   5985
         Width           =   3210
      End
      Begin VB.Label Label1 
         Caption         =   "Precio por Kilo Gtos.Molturación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   107
         Left            =   -69780
         TabIndex        =   370
         Top             =   5565
         Width           =   3210
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   -72300
         ToolTipText     =   "Buscar Iva"
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cód.IVA Fact.Internas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   106
         Left            =   -74610
         TabIndex        =   369
         Top             =   2880
         Width           =   2265
      End
      Begin VB.Label Label1 
         Caption         =   "Constante Faneca"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   102
         Left            =   -74550
         TabIndex        =   362
         Top             =   1230
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje Inc/Dec.Aforo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   101
         Left            =   -67560
         TabIndex        =   361
         Top             =   6090
         Width           =   2880
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Máximo Jornadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   100
         Left            =   -74430
         TabIndex        =   360
         Top             =   3960
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje Jornadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   97
         Left            =   -74430
         TabIndex        =   358
         Top             =   2130
         Width           =   2280
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje IRPF"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   96
         Left            =   -74430
         TabIndex        =   357
         Top             =   3480
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje Seg.Social 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   95
         Left            =   -74430
         TabIndex        =   356
         Top             =   3030
         Width           =   2790
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje Seg.Social 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   94
         Left            =   -74430
         TabIndex        =   355
         Top             =   2580
         Width           =   2550
      End
      Begin VB.Label Label31 
         Caption         =   "Euros Trab/dia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74460
         TabIndex        =   354
         Top             =   1530
         Width           =   1695
      End
      Begin VB.Label Label29 
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74340
         TabIndex        =   334
         Top             =   3675
         Width           =   1680
      End
      Begin VB.Label Label26 
         Caption         =   "Derrama"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70800
         TabIndex        =   333
         Top             =   3705
         Width           =   1590
      End
      Begin VB.Label Label28 
         Caption         =   "Hasta Metros Cúbicos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73200
         TabIndex        =   332
         Top             =   1050
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Código IVA"
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
         Index           =   85
         Left            =   -74340
         TabIndex        =   331
         Top             =   2535
         Width           =   1290
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   -72870
         ToolTipText     =   "Buscar Iva"
         Top             =   2505
         Width           =   240
      End
      Begin VB.Label Label27 
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
         Height          =   285
         Left            =   -71370
         TabIndex        =   329
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Rango 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74340
         TabIndex        =   328
         Top             =   1290
         Width           =   1635
      End
      Begin VB.Label Label24 
         Caption         =   "Rango 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74340
         TabIndex        =   327
         Top             =   1710
         Width           =   1590
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   -72795
         ToolTipText     =   "Buscar Concepto Gasto"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Código Gasto para Liq."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   76
         Left            =   -74505
         TabIndex        =   285
         Top             =   3450
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje Gasto Mto."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   75
         Left            =   -74505
         TabIndex        =   283
         Top             =   3060
         Width           =   2040
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   -72645
         ToolTipText     =   "Buscar Almacén"
         Top             =   1065
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Almacén"
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
         Left            =   -74460
         TabIndex        =   273
         Top             =   1050
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Letra Serie Clientes"
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
         Left            =   -74505
         TabIndex        =   264
         Top             =   1515
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Ventas"
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
         Left            =   -74505
         TabIndex        =   263
         Top             =   1050
         Width           =   1650
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   70
         Left            =   -72840
         ToolTipText     =   "Buscar cuenta"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   69
         Left            =   -72795
         ToolTipText     =   "Buscar cuenta"
         Top             =   2595
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Ventas"
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
         Left            =   -74505
         TabIndex        =   259
         Top             =   2625
         Width           =   1650
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   59
         Left            =   -72795
         ToolTipText     =   "Buscar cuenta"
         Top             =   2190
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Banco Prevista"
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
         Left            =   -74505
         TabIndex        =   186
         Top             =   2220
         Width           =   1650
      End
      Begin VB.Label Label22 
         Caption         =   "Peso Caja Llena"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67440
         TabIndex        =   184
         Top             =   4830
         Width           =   1830
      End
      Begin VB.Label Label21 
         Caption         =   "Mínimo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -68880
         TabIndex        =   183
         Top             =   5190
         Width           =   990
      End
      Begin VB.Label Label20 
         Caption         =   "Máximo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67050
         TabIndex        =   182
         Top             =   5190
         Width           =   990
      End
      Begin VB.Label Label19 
         Caption         =   "Límites Kilos Caja"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70860
         TabIndex        =   181
         Top             =   5520
         Width           =   1950
      End
      Begin VB.Label Label1 
         Caption         =   "Sección Bodega"
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
         Index           =   53
         Left            =   -74550
         TabIndex        =   180
         Top             =   1050
         Width           =   1650
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   -72840
         ToolTipText     =   "Buscar Sección"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Base Datos Ariges"
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
         Index           =   52
         Left            =   -74580
         TabIndex        =   171
         Top             =   1590
         Width           =   1650
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   60
         Left            =   -72330
         ToolTipText     =   "Buscar Sección"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sección Suministros"
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
         Index           =   46
         Left            =   -74580
         TabIndex        =   170
         Top             =   1050
         Width           =   2340
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   58
         Left            =   -72300
         ToolTipText     =   "Buscar cuenta"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta Banco Prevista"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   44
         Left            =   -74610
         TabIndex        =   166
         Top             =   2370
         Width           =   1980
      End
      Begin VB.Label Label1 
         Caption         =   "Almacén"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   42
         Left            =   -74580
         TabIndex        =   164
         Top             =   1500
         Width           =   1140
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   -72255
         ToolTipText     =   "Buscar Almacén"
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   -72255
         ToolTipText     =   "Buscar Sección"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sección ADV"
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
         Left            =   -74580
         TabIndex        =   162
         Top             =   1050
         Width           =   1770
      End
      Begin VB.Label Label17 
         Caption         =   "Path Ficheros clasificación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74550
         TabIndex        =   160
         Top             =   5790
         Width           =   3000
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   -72060
         ToolTipText     =   "Buscar Sección"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sección Almazara"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   34
         Left            =   -74490
         TabIndex        =   159
         Top             =   1050
         Width           =   2340
      End
      Begin VB.Label Label16 
         Caption         =   "Texto Pie de Toma de Datos"
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
         Left            =   -74580
         TabIndex        =   131
         Top             =   4740
         Width           =   2925
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   1
         Left            =   -71640
         ToolTipText     =   "Zoom descripción"
         Top             =   4770
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   -71010
         ToolTipText     =   "Zoom descripción"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label15 
         Caption         =   "Texto Cabecera de Toma de Datos"
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
         Left            =   -74580
         TabIndex        =   130
         Top             =   3300
         Width           =   3525
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje AFO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   -74550
         TabIndex        =   129
         Top             =   2850
         Width           =   2040
      End
      Begin VB.Label Label14 
         Caption         =   "Impresora de Entradas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74580
         TabIndex        =   121
         Top             =   4320
         Width           =   2940
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje Retención"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   58
         Left            =   -74550
         TabIndex        =   120
         Top             =   2460
         Width           =   2700
      End
      Begin VB.Label Label1 
         Caption         =   "Sección Hortofrutícola"
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
         Index           =   0
         Left            =   -74550
         TabIndex        =   119
         Top             =   840
         Width           =   2220
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   -72120
         ToolTipText     =   "Buscar Sección"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Coste Seg.Social"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74550
         TabIndex        =   117
         Top             =   2040
         Width           =   2130
      End
      Begin VB.Label Label12 
         Caption         =   "Coste Horas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74550
         TabIndex        =   116
         Top             =   1620
         Width           =   2025
      End
      Begin VB.Label Label11 
         Caption         =   "Cajas por Palet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70860
         TabIndex        =   107
         Top             =   4830
         Width           =   1590
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -72300
         ToolTipText     =   "Buscar Extensión"
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Extensión"
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
         Left            =   -74340
         TabIndex        =   97
         Top             =   915
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   10110
      TabIndex        =   79
      Top             =   8700
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   240
      TabIndex        =   81
      Top             =   8565
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   210
         Width           =   2760
      End
   End
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
      Left            =   8970
      TabIndex        =   78
      Top             =   8700
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   10125
      TabIndex        =   80
      Top             =   8700
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3630
      Top             =   5250
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   10920
      TabIndex        =   433
      Top             =   180
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnAñadir 
         Caption         =   "&Añadir"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ### [Monica] 06/09/2006
' procedimiento nuevo introducido de la gestion

Option Explicit

Private Const IdPrograma = 1002


Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmIva As frmTipIVAConta
Attribute frmIva.VB_VarHelpID = -1
Private WithEvents frmDoc As frmCarpetaAridoc
Attribute frmDoc.VB_VarHelpID = -1
Private WithEvents frmExt As frmExtAridoc
Attribute frmExt.VB_VarHelpID = -1
Private WithEvents frmAri As frmCarpAridoc
Attribute frmAri.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSeccion
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmAlm As frmBasico2
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmAlm2 As frmBasico2
Attribute frmAlm2.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmConGasto As frmManConcepGasto
Attribute frmConGasto.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarTra
Attribute frmTar.VB_VarHelpID = -1
Private WithEvents frmCCos As frmCCosConta 'centros de coste
Attribute frmCCos.VB_VarHelpID = -1
Private WithEvents frmArtADV As frmADVArticulos 'articulos de adv
Attribute frmArtADV.VB_VarHelpID = -1
Private WithEvents frmCarSocio As frmCartasSocio 'cartas de socio
Attribute frmCarSocio.VB_VarHelpID = -1


Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Dim Indice As Byte
Dim Encontrado As Boolean
Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: Añadir
'4: Modificar

Dim vSeccion As CSeccion
Dim indCodigo As Integer


Private Sub chkAgruparNotas_Click()
    If (chkAgruparNotas.Value = 1) Then
        Me.chkRespetarNroNota.Enabled = False
        Me.chkRespetarNroNota.Value = 0
    Else
        Me.chkRespetarNroNota.Enabled = True
    End If
End Sub

Private Sub chkOutlook_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkOutlook_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkTaraTractor_GotFocus()
    PonerFocoChk chkTaraTractor
End Sub

Private Sub chkTaraTractor_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkTraza_GotFocus()
    PonerFocoChk chkTraza
End Sub

Private Sub chkTraza_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkNotaManual_GotFocus()
    PonerFocoChk chkNotaManual
End Sub

Private Sub chkNotaManual_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkCoopro_GotFocus()
    PonerFocoChk chkCoopro
End Sub

Private Sub chkCoopro_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub cmdAceptar_Click()
Dim actualiza As Boolean
Dim kms As Currency

    
'    If Modo = 3 Then
'        If DatosOk Then
'            'Cambiamos el path
'            'CambiaPath True
'            If InsertarDesdeForm(Me) Then
'                PonerModo 0
''                ActualizaNombreEmpresa
'                MsgBox "Debe salir de la aplicacion para que los cambios tengan efecto", vbExclamation
'            End If
'
'        End If
'    End If


    If Modo = 4 Then 'MODIFICAR
        If DatosOk Then
            If Not vParamAplic Is Nothing Then
                'Datos contabilidad
                vParamAplic.ServidorConta = Text1(1).Text
                vParamAplic.UsuarioConta = Text1(2).Text
                vParamAplic.PasswordConta = Text1(3).Text
                vParamAplic.NumeroConta = ComprobarCero(Text1(4).Text)
                
                vParamAplic.WebSoporte = Text1(9).Text
                vParamAplic.DireMail = Text1(5).Text
                vParamAplic.Smtphost = Text1(6).Text
                vParamAplic.SmtpUser = Text1(7).Text
                vParamAplic.Smtppass = Text1(8).Text
                
                ' envio de email por outlook
                vParamAplic.EnvioDesdeOutlook = Me.chkOutlook.Value
            
                ' Para utilizar el arigesmail
                vParamAplic.ExeEnvioMail = Trim(Text1(125).Text)
                
                
                ' SMS
                vParamAplic.SMSemail = Text1(118).Text
                vParamAplic.SMSclave = Text1(119).Text
                vParamAplic.SMSremitente = Text1(120).Text
               
                ' entradas de almacen
                vParamAplic.SeTaraTractor = Me.chkTaraTractor.Value
                vParamAplic.HayTraza = Me.chkTraza.Value
                vParamAplic.CajasporPalet = ComprobarCero(Text1(24).Text)
                vParamAplic.SeAgrupanNotas = Me.chkAgruparNotas.Value
                vParamAplic.SeRespetaNota = Me.chkRespetarNroNota.Value
                vParamAplic.NroNotaManual = Me.chkNotaManual.Value
                vParamAplic.CooproenEntradas = Me.chkCoopro.Value
                
                
                vParamAplic.TipoCaja1 = Text1(14).Text
                vParamAplic.TipoCaja2 = Text1(15).Text
                vParamAplic.TipoCaja3 = Text1(16).Text
                vParamAplic.TipoCaja4 = Text1(17).Text
                vParamAplic.TipoCaja5 = Text1(18).Text
                vParamAplic.PesoCaja1 = ComprobarCero(Text1(19).Text)
                vParamAplic.PesoCaja2 = ComprobarCero(Text1(20).Text)
                vParamAplic.PesoCaja3 = ComprobarCero(Text1(21).Text)
                vParamAplic.PesoCaja4 = ComprobarCero(Text1(22).Text)
                vParamAplic.PesoCaja5 = ComprobarCero(Text1(23).Text)
                
                vParamAplic.PesoCaja11 = ComprobarCero(Text1(103).Text)
                vParamAplic.PesoCaja12 = ComprobarCero(Text1(104).Text)
                vParamAplic.PesoCaja13 = ComprobarCero(Text1(105).Text)
                vParamAplic.PesoCaja14 = ComprobarCero(Text1(106).Text)
                vParamAplic.PesoCaja15 = ComprobarCero(Text1(107).Text)
                
                vParamAplic.EsCaja1 = Me.ChkCajas(0).Value
                vParamAplic.EsCaja2 = Me.ChkCajas(1).Value
                vParamAplic.EsCaja3 = Me.ChkCajas(2).Value
                vParamAplic.EsCaja4 = Me.ChkCajas(3).Value
                vParamAplic.EsCaja5 = Me.ChkCajas(4).Value
                vParamAplic.KilosCajaMin = ComprobarCero(Text1(64).Text)
                vParamAplic.KilosCajaMax = ComprobarCero(Text1(65).Text)
                vParamAplic.PesoCajaLLena = ComprobarCero(Text1(66).Text)
                
                
                vParamAplic.EsVtaFruta1 = Me.ChkVtaFruta(0).Value
                vParamAplic.EsVtaFruta2 = Me.ChkVtaFruta(1).Value
                vParamAplic.EsVtaFruta3 = Me.ChkVtaFruta(2).Value
                vParamAplic.EsVtaFruta4 = Me.ChkVtaFruta(3).Value
                vParamAplic.EsVtaFruta5 = Me.ChkVtaFruta(4).Value
                
                vParamAplic.PorcIncreAforo = ComprobarCero(Text1(109).Text)
                
                vParamAplic.ImpresoraEntradas = Replace(Text1(31).Text, "\", "\\")
                
                vParamAplic.SigPac = Text1(85).Text
                vParamAplic.GoolZoom = Text1(95).Text
                
                '[Monicax]12/03/2018: euros de capital social
                vParamAplic.EurCapSocial = Text1(142).Text
                
                'aridoc
                vParamAplic.CarpetaAnt = ComprobarCero(Text1(10))
                vParamAplic.CarpetaLiq = ComprobarCero(Text1(11))
                vParamAplic.CarpetaADV = ComprobarCero(Text1(61))
                vParamAplic.CarpetaRecAlmacen = ComprobarCero(Text1(73))
                vParamAplic.CarpetaRecCampo = ComprobarCero(Text1(74))
                vParamAplic.CarpetaAlmz = ComprobarCero(Text1(67))
                vParamAplic.CarpetaBOD = ComprobarCero(Text1(68))
                vParamAplic.CarpetaTra = ComprobarCero(Text1(96))

                vParamAplic.Extension = Text1(13)
                
                vParamAplic.C1Anticipo = Combo1(1).ListIndex
                vParamAplic.C2Anticipo = Combo1(2).ListIndex
                vParamAplic.C3Anticipo = Combo1(3).ListIndex
                vParamAplic.C4Anticipo = Combo1(4).ListIndex
                vParamAplic.C1Liquidacion = Combo1(5).ListIndex
                vParamAplic.C2Liquidacion = Combo1(6).ListIndex
                vParamAplic.C3Liquidacion = Combo1(7).ListIndex
                vParamAplic.C4Liquidacion = Combo1(8).ListIndex
                vParamAplic.C1ADV = Combo1(9).ListIndex
                vParamAplic.C2ADV = Combo1(10).ListIndex
                vParamAplic.C3ADV = Combo1(11).ListIndex
                vParamAplic.C4ADV = Combo1(12).ListIndex
                
                vParamAplic.C1Recibo = Combo1(21).ListIndex
                vParamAplic.C2Recibo = Combo1(22).ListIndex
                vParamAplic.C3Recibo = Combo1(23).ListIndex
                vParamAplic.C4Recibo = Combo1(24).ListIndex
                
                vParamAplic.C1Almz = Combo1(13).ListIndex
                vParamAplic.C2Almz = Combo1(14).ListIndex
                vParamAplic.C3Almz = Combo1(15).ListIndex
                vParamAplic.C4Almz = Combo1(16).ListIndex
                
                vParamAplic.C1BOD = Combo1(17).ListIndex
                vParamAplic.C2BOD = Combo1(18).ListIndex
                vParamAplic.C3BOD = Combo1(19).ListIndex
                vParamAplic.C4BOD = Combo1(20).ListIndex
                
                vParamAplic.C1Transporte = Combo1(25).ListIndex
                vParamAplic.C2Transporte = Combo1(26).ListIndex
                vParamAplic.C3Transporte = Combo1(27).ListIndex
                vParamAplic.C4Transporte = Combo1(28).ListIndex
                
                
                vParamAplic.Faneca = ComprobarCero(Text1(110).Text)
                vParamAplic.CosteHora = ComprobarCero(Text1(25).Text)
                vParamAplic.CosteSegSo = ComprobarCero(Text1(26).Text)
                vParamAplic.Seccionhorto = Text1(27).Text ' ComprobarCero(Text1(27).Text)
                vParamAplic.SeccionAlmaz = Text1(48).Text 'ComprobarCero(Text1(48).Text)
                vParamAplic.SeccionADV = Text1(56).Text 'ComprobarCero(Text1(56).Text)
                vParamAplic.PorcreteFacSoc = ComprobarCero(Text1(28).Text)
                vParamAplic.SeccionPOZOS = Text1(121).Text ' seccion de pozos
                
                vParamAplic.PrimFactAnt = ComprobarCero(Text1(29).Text)
                vParamAplic.UltFactAnt = ComprobarCero(Text1(30).Text)
                vParamAplic.PrimFactAntVC = ComprobarCero(Text1(32).Text)
                vParamAplic.UltFactAntVC = ComprobarCero(Text1(33).Text)
                vParamAplic.PrimFactLiq = ComprobarCero(Text1(34).Text)
                vParamAplic.UltFactLiq = ComprobarCero(Text1(35).Text)
                vParamAplic.PrimFactLiqVC = ComprobarCero(Text1(36).Text)
                vParamAplic.UltFactLiqVC = ComprobarCero(Text1(12).Text)
                
                vParamAplic.PorcenAFO = ComprobarCero(Text1(37).Text)
                vParamAplic.TTomaDatos = Text1(38).Text
                vParamAplic.PieTomaDatos = Text1(39).Text
                vParamAplic.CodIvaIntra = ComprobarCero(Text1(40).Text)
                vParamAplic.CtaTerReten = Text1(42).Text
                
                vParamAplic.CtaTraReten = Text1(117).Text ' cuenta de retencion de transporte
                
                vParamAplic.PathTraza = Text1(41).Text
                
                vParamAplic.ForpaPosi = ComprobarCero(Text1(43).Text)
                vParamAplic.ForpaNega = ComprobarCero(Text1(44).Text)
                vParamAplic.CtaRetenSoc = Text1(45).Text
                vParamAplic.CtaAportaSoc = Text1(46).Text
                vParamAplic.CtaBancoSoc = Text1(47).Text
                
                ' ALMAZARA
                vParamAplic.ForpaPosiAlmz = ComprobarCero(Text1(51).Text)
                vParamAplic.ForpaNegaAlmz = ComprobarCero(Text1(52).Text)
                vParamAplic.CtaRetenAlmz = Text1(53).Text
                vParamAplic.CtaBancoAlmz = Text1(54).Text
                vParamAplic.CtaVentasAlmz = Text1(49).Text
                vParamAplic.CtaGastosAlmz = Text1(50).Text
                vParamAplic.LetraSerieAlmz = Text1(55).Text
                vParamAplic.CodGastoAlmz = ComprobarCero(Text1(112).Text)
                vParamAplic.GtoMoltura = ComprobarCero(Text1(115).Text)
                vParamAplic.GtoEnvasado = ComprobarCero(Text1(116).Text)
                
                vParamAplic.PrimFactAntAlmz = ComprobarCero(Text1(81).Text)
                vParamAplic.UltFactAntAlmz = ComprobarCero(Text1(82).Text)
                vParamAplic.PrimFactLiqAlmz = ComprobarCero(Text1(83).Text)
                vParamAplic.UltFactLiqAlmz = ComprobarCero(Text1(84).Text)
                
                
                ' ADV
                vParamAplic.AlmacenADV = ComprobarCero(Text1(57).Text)
                vParamAplic.CtaBancoADV = Text1(58).Text
                vParamAplic.CodIvaExeADV = ComprobarCero(Text1(114).Text)
                
                
                ' Suministros
                vParamAplic.SeccionSumi = Text1(60).Text 'ComprobarCero(Text1(60).Text)
                vParamAplic.BDAriges = Text1(62).Text
                
                ' Bodega
                vParamAplic.SeccionBodega = Text1(63).Text ' ComprobarCero(Text1(63).Text)
                vParamAplic.AlbRetiradaManual = Me.ChkContadorManual.Value
                vParamAplic.CtaBancoBOD = Text1(59).Text
                vParamAplic.CtaVentasBOD = Text1(69).Text
                vParamAplic.PorcGtoMantBOD = ComprobarCero(Text1(75).Text)
                vParamAplic.CodGastoBOD = ComprobarCero(Text1(76).Text)
                
                vParamAplic.PrimFactAntBOD = ComprobarCero(Text1(77).Text)
                vParamAplic.UltFactAntBOD = ComprobarCero(Text1(78).Text)
                vParamAplic.PrimFactLiqBOD = ComprobarCero(Text1(79).Text)
                vParamAplic.UltFactLiqBOD = ComprobarCero(Text1(80).Text)
                '[Monica]27/08/2012: incremento de kilos de entrada
                vParamAplic.PorcKilosBOD = ComprobarCero(Text1(128).Text)
                
                ' Telefonia
                vParamAplic.CtaVentasTel = Text1(70).Text
                vParamAplic.LetraSerieTel = Text1(71).Text
                
                ' NOMinas
                vParamAplic.AlmacenNOMI = ComprobarCero(Text1(72).Text)
                vParamAplic.EurosTrabdiaNOMI = ComprobarCero(Text1(97).Text)
                vParamAplic.PorcSegSo1NOMI = ComprobarCero(Text1(98).Text)
                vParamAplic.PorcSegSo2NOMI = ComprobarCero(Text1(99).Text)
                vParamAplic.PorcIRPFNOMI = ComprobarCero(Text1(100).Text)
                vParamAplic.PorcJornadaNOMI = ComprobarCero(Text1(101).Text)
                vParamAplic.NroMaxJornadasNOMI = ComprobarCero(Text1(108).Text)
                
                ' POZOS
                vParamAplic.Consumo1POZ = ComprobarCero(Text1(86).Text)
                vParamAplic.Precio1POZ = ComprobarCero(Text1(87).Text)
                vParamAplic.Consumo2POZ = ComprobarCero(Text1(88).Text)
                vParamAplic.Precio2POZ = ComprobarCero(Text1(89).Text)
                vParamAplic.Consumo3POZ = ComprobarCero(Text1(138).Text)
                vParamAplic.Precio3POZ = ComprobarCero(Text1(137).Text)
                vParamAplic.CodIvaPOZ = ComprobarCero(Text1(90).Text)
                vParamAplic.CuotaPOZ = ComprobarCero(Text1(91).Text)
                vParamAplic.DerramaPOZ = ComprobarCero(Text1(92).Text)
                vParamAplic.CanonPOZ = ComprobarCero(Text1(139).Text)
                
                vParamAplic.CoefConsumoPOZ = ComprobarCero(Text1(140).Text)
                vParamAplic.CoefSuministroPOZ = ComprobarCero(Text1(141).Text)
                
                
                vParamAplic.CtaVentasConsPOZ = Text1(122).Text
                vParamAplic.CtaVentasCuoPOZ = Text1(123).Text
                vParamAplic.CtaVentasTalPOZ = Text1(129).Text
                vParamAplic.CtaVentasMtoPOZ = Text1(130).Text
                vParamAplic.CtaVentasMantaPOZ = Text1(134).Text
                '[Monica]21/01/2016: nueva cuenta de recargos
                vParamAplic.CtaRecargosPOZ = Text1(135).Text
                
                vParamAplic.CodCCostPOZ = Text1(124).Text
                vParamAplic.ForpaConPOZ = ComprobarCero(Text1(126).Text)
                vParamAplic.ForpaRecPOZ = ComprobarCero(Text1(127).Text)
                vParamAplic.CartaPOZ = ComprobarCero(Text1(131).Text)
                
                ' transporte
                vParamAplic.TipoPortesTRA = Combo1(0).ListIndex
                vParamAplic.TarifaTRA = ComprobarCero(Text1(93).Text)
                vParamAplic.TarifaTRA2 = ComprobarCero(Text1(113).Text)
                vParamAplic.CodGastoTRA = ComprobarCero(Text1(94).Text)
                vParamAplic.PorcreteFacTra = ComprobarCero(Text1(102).Text)
                vParamAplic.PrecioKgTra = ComprobarCero(Text1(111).Text)
                vParamAplic.TipoContadorTRA = Combo1(29).ListIndex
                
                vParamAplic.ConsumoMinPOZ = ComprobarCero(Text1(132).Text)
                vParamAplic.ConsumoMaxPOZ = ComprobarCero(Text1(133).Text)
                
                '[Monica]17/10/2016: directorio de impresion de entradas
                vParamAplic.PathEntradas = Text1(136).Text
                
                
                
                actualiza = vParamAplic.Modificar()
                TerminaBloquear
    
                If actualiza Then  'Inserta o Modifica
                    'Abrir la conexion a la conta q hemos modificado
                    CerrarConexionConta
                    If vParamAplic.NumeroConta <> 0 Then
                        If Not AbrirConexionConta() Then End
                        LeerNivelesEmpresa
                    End If
                    BloqueoMenusSegunCooperativa
                    PonerModo 2
                    PonerFocoBtn Me.cmdSalir
                End If
           End If
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    TerminaBloquear
    If Data1.Recordset.EOF Then
        PonerModo 0
    Else
        PonerCampos
        PonerModo 2
    End If
End Sub

Private Sub cmdSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 0 Then PonerCadenaBusqueda
    PonerFoco Text1(0)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Byte
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(4).Image = 11  'Salir
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With
    
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
 
    LimpiarCampos   'Limpia los campos TextBox
   
   'cargar IMAGES de busqueda
    For i = 0 To 25
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 43 To 47
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 49 To 50
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    For i = 51 To 54
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 58 To 60
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 69 To 70
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 122 To 124
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 126 To 127
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 129 To 131
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 134 To 135
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i

    
    ' el codigo de cartas de pozos solo es visible para utxera y escalona
    Me.Label1(131).visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
    Me.imgBuscar(131).Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
    Me.imgBuscar(131).visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
    Text1(131).Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
    Text1(131).visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
    Text2(131).Enabled = False
    Text2(131).visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)


    SSTab1.Tab = 0

    NombreTabla = "rparam"
    Ordenacion = " ORDER BY codparam"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    Encontrado = True
    If Data1.Recordset.EOF Then
        'No hay registro de datos de parametros
        'quitar###
        Encontrado = False
    End If
    
    CargaCombo
        
    Me.SSTab1.TabEnabled(3) = (vParamAplic.HayAridoc = 1)
    Me.SSTab1.TabVisible(3) = (vParamAplic.HayAridoc = 1)
    If (vParamAplic.HayAridoc = 1) Then
        Me.SSTab1.TabsPerRow = 7
        AbrirConexionAridoc "root", "aritel"
        Me.SSTab2.Tab = 0
    Else
        Me.SSTab1.TabsPerRow = 6
    End If
    
    PonerModo 0

    '[Monica]22/09/2017
    If vParamAplic.Cooperativa = 17 Then
        Label29.Caption = "Cuota Mantenim."
        Label26.Caption = "Cuota Servicio"
    End If


End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
'        Me.Toolbar1.Buttons(1).Enabled = False 'Modificar
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CerrarConexionAridoc
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    Text1(57).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    FormateaCampo Text1(57)
    Text2(57).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

' Almacen de la gesion de nominas
Private Sub frmAlm2_DatoSeleccionado(CadenaSeleccion As String)
    Text1(72).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    FormateaCampo Text1(72)
    Text2(72).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub


Private Sub frmArtADV_DatoSeleccionado(CadenaSeleccion As String)
'Articulo de mano de obra de adv para quatretonda
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCarSocio_DatoSeleccionado(CadenaSeleccion As String)
'Carta de reclamacion de pozos
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCCos_DatoSeleccionado(CadenaSeleccion As String)
'Centro de Coste de la contabilidad
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmConGasto_DatoSeleccionado(CadenaSeleccion As String)
'Concepto de gasto de bodega
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codgasto
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nomgasto
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmDoc_DatoSeleccionado(CadenaSeleccion As String)
'Carpetas de Aridoc
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre carpeta
End Sub

Private Sub frmExt_DatoSeleccionado(CadenaSeleccion As String)
'Extension de Aridoc
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmAri_DatoSeleccionado(CadenaSeleccion As String)
Dim cad As String
    cad = RecuperaValor(CadenaSeleccion, 1)
    Text1(Indice).Text = Mid(cad, 2, Len(cad))
    Text1(Indice).Text = Format(Text1(Indice).Text, "000")
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(Indice).Text = Format(Text1(Indice).Text, "000")
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de iva de la contabilidad
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Porceiva
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
' tarifa de transporte
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Éstas son las tarifas base, dependiendo del tipo de Tarifas en el  " & vbCrLf & _
                      "mantenimiento de Tarifas de Transporte. " & vbCrLf & vbCrLf & _
                      "Para el cálculo del canon de caminos, en algunas cooperativas se " & vbCrLf & _
                      "utiliza el tipo de tarifa en función de que la variedad sea de un " & vbCrLf & _
                      "producto o no lo sea." & vbCrLf & vbCrLf
                                            
        Case 1
           ' "____________________________________________________________"
            vCadena = "Precio por kilo de gastos de molturación y precio por litros de   " & vbCrLf & _
                      "gastos de envasado que se aplican en la liquidación de Almazara. " & vbCrLf & vbCrLf & _
                      "Los kilos / litros utilizados para realizar el cálculo son " & vbCrLf & _
                      "únicamente los de autoconsumo." & vbCrLf & vbCrLf
                      
        Case 2
           ' "____________________________________________________________"
            vCadena = "Si está marcado se desdoblan las entradas, según sus coopropietarios," & vbCrLf & _
                      "cuando se actualizan las entradas y pasan a entradas clasificadas. " & vbCrLf & vbCrLf & _
                      "Si no está marcado las entradas se desdoblarán cuando actualicen las" & vbCrLf & _
                      "entradas clasificadas y pasen al Histórico de Entradas." & vbCrLf & vbCrLf
        
        Case 3
           ' "____________________________________________________________"
            vCadena = "Porcentaje de incremento de kilos netos de las entradas de bodega. " & vbCrLf & vbCrLf & _
                      "Se incrementará el resultado de kilos brutos menos la tara, en el " & vbCrLf & _
                      "porcentaje indicado sólo en el caso de que sea una entrada de tipo" & vbCrLf & _
                      "Producto Integrado" & vbCrLf & _
                      vbCrLf
             
        Case 4
           ' "____________________________________________________________"
            vCadena = "Euros que multiplicado por la superficie cultivable de los campos " & vbCrLf & _
                      "nos indica el importe a pagar al socio cuando se da de baja. " & vbCrLf & vbCrLf & _
                      "Corresponde al Capital Social del socio cuando se da de alta." & vbCrLf & _
                      vbCrLf
        
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim numnivel As Byte

TerminaBloquear
    
    If vParamAplic.NumeroConta = 0 Then Exit Sub
    
    Select Case Index
        Case 0, 3, 4, 60, 10, 25
            Select Case Index
                Case 0 ' Seccion hortofrutícola
                    Indice = Index + 27
                Case 3 ' seccion de Almazara
                    Indice = 48
                Case 4 ' seccion de Adv
                    Indice = 56
                Case 60 ' seccion de suministros
                    Indice = Index
                Case 10 ' seccion de bodega
                    Indice = 63
                Case 25 ' seccion de pozos
                    Indice = 121
            End Select
            
            Set frmSec = New frmManSeccion
            frmSec.DatosADevolverBusqueda = "0|1|2|"
            frmSec.CodigoActual = Text1(Indice).Text
            frmSec.Show vbModal
            Set frmSec = Nothing
            PonerFoco Text1(Indice)
        
        Case 1  'Porcentaje iva de factura de terceros de extranjero
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = Index + 39
                    Set frmIva = New frmTipIVAConta
                    frmIva.DatosADevolverBusqueda = "0|1|2|"
                    frmIva.CodigoActual = Text1(Indice).Text
                    frmIva.Show vbModal
                    Set frmIva = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
    
        Case 2  'Cuenta Contable Retencion facturas terceros
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = Index + 40
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
         Case 24 'Cuenta Contable Retencion facturas transporte
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = 117
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing

         Case 6, 7, 8, 11, 12, 14, 15, 20 'carpetas de aridoc
            Select Case Index
                Case 8
                    Indice = 61
                Case 6, 7
                    Indice = Index + 4
                Case 11, 12
                    Indice = Index + 56
                Case 14, 15
                    Indice = Index + 59
                Case 20
                    Indice = Index + 76
            End Select
            
            Set frmAri = New frmCarpAridoc
            frmAri.Opcion = 20
            frmAri.Show vbModal
            Set frmAri = Nothing
            PonerFoco Text1(Indice)
        
         Case 9 'extesion de fichero de aridoc
            Indice = Index + 4
            Set frmExt = New frmExtAridoc
            frmExt.DatosADevolverBusqueda = "0|1|"
            frmExt.CodigoActual = Text1(Indice).Text
            frmExt.Show vbModal
            Set frmExt = Nothing
            PonerFoco Text1(Indice)
                
        Case 43, 44 ' forma de pago de facturas de anticipos / liquidaciones socios
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    AbrirFrmForpaConta (Index)
                End If
            End If
        
        Case 45, 46, 47, 70 ' cuenta de retencion y de aportacion de facturas anti / liqui de socios
                        ' 47 cta de banco prevista
                        ' 70 cta de ventas de telefonia
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = Index
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
        

        '53,54 cta de retencion de almazara y cta banco almazara
        '49,50 cuenta de ventas y de gastos de la almazara
        Case 53, 54, 49, 50
            If Text1(48).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(48).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = Index
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing

        Case 51, 52 ' forma de pago de facturas de almazara
            If Text1(48).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(48).Text) Then
                If vSeccion.AbrirConta Then
                    AbrirFrmForpaConta (Index)
                End If
            End If
            
        Case 5 ' almacen de adv
            Set frmAlm = New frmBasico2
            
            AyudaAlmacenCom frmAlm, Text1(57).Text
            
            Set frmAlm = Nothing

        '58 cta banco prevista adv
        Case 58
            If Text1(48).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(56).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = Index
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
        Case 23 ' codigo de iva de facturas internas de adv
            If Text1(56).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(56).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = 114
                    Set frmIva = New frmTipIVAConta
                    frmIva.DatosADevolverBusqueda = "0|1|2|"
                    frmIva.CodigoActual = Text1(Indice).Text
                    frmIva.Show vbModal
                    Set frmIva = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
        Case 59, 69 ' Cta.banco prevista y cta de ventas de bodega
            If Text1(63).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(63).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = Index
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
        
        
        Case 16 ' concepto de gasto para liquidacion
            Indice = Index + 60
            
            Set frmConGasto = New frmManConcepGasto
            frmConGasto.DatosADevolverBusqueda = "0|1|"
            frmConGasto.CodigoActual = Text1(Indice).Text
            frmConGasto.Show vbModal
            Set frmConGasto = Nothing
            PonerFoco Text1(Indice)
        
        Case 21 ' concepto de gasto para liquidacion
            Indice = Index + 91
            
            Set frmConGasto = New frmManConcepGasto
            frmConGasto.DatosADevolverBusqueda = "0|1|"
            frmConGasto.CodigoActual = Text1(Indice).Text
            frmConGasto.Show vbModal
            Set frmConGasto = Nothing
            PonerFoco Text1(Indice)
        
        ' Nominas
        Case 13 ' almacen de nominas
            Set frmAlm2 = New frmBasico2
            
            AyudaAlmacenCom frmAlm2, Text1(72).Text
            
            Set frmAlm2 = Nothing
        
        
        ' pozos
        Case 17 ' codigo de iva de pozos
            If Text1(121).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(121).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = 90
                    Set frmIva = New frmTipIVAConta
                    frmIva.DatosADevolverBusqueda = "0|1|2|"
                    frmIva.CodigoActual = Text1(Indice).Text
                    frmIva.Show vbModal
                    Set frmIva = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
        
        Case 126, 127 ' forma de pago de facturas de contado y recibos de pozos
            If Text1(121).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(121).Text) Then
                If vSeccion.AbrirConta Then
                    AbrirFrmForpaConta (Index)
                End If
            End If
            
       Case 18, 22  ' 18 = codigo de tarifa 1
                ' 22 = codigo de tarifa 2
            If Index = 18 Then
                indCodigo = 93
            Else
                indCodigo = 113
            End If
            Set frmTar = New frmManTarTra
            frmTar.DatosADevolverBusqueda = "0|1|"
            frmTar.CodigoActual = Text1(indCodigo).Text
            frmTar.Show vbModal
            Set frmTar = Nothing
            PonerFoco Text1(indCodigo)
        
       Case 22  ' codigo de tarifa 2
            Set frmTar = New frmManTarTra
            frmTar.DatosADevolverBusqueda = "0|1|"
            frmTar.CodigoActual = Text1(93).Text
            frmTar.Show vbModal
            Set frmTar = Nothing
            PonerFoco Text1(93)
        
        
       Case 19 ' concepto de gasto para transporte
            Indice = 94
            
            Set frmConGasto = New frmManConcepGasto
            frmConGasto.DatosADevolverBusqueda = "0|1|"
            frmConGasto.CodigoActual = Text1(Indice).Text
            frmConGasto.Show vbModal
            Set frmConGasto = Nothing
            PonerFoco Text1(Indice)
 
        
        
        '122 cuenta de ventas consumo de pozos
        '123 cuenta de ventas cuota pozos
        Case 122, 123, 129, 130, 134, 135
            If Text1(121).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(121).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = Index
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
        ' centro de coste de ventas
        Case 124
            If Text1(121).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(121).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = Index
                    Set frmCCos = New frmCCosConta
                    frmCCos.DatosADevolverBusqueda = "0|1|"
                    frmCCos.CodigoActual = Text1(Indice).Text
                    frmCCos.Show vbModal
                    Set frmCCos = Nothing
                    PonerFoco Text1(Indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
        
        Case 26 ' codigo de articulo de adv
            Indice = Index + 100
            
            Set frmArtADV = New frmADVArticulos
            frmArtADV.DatosADevolverBusqueda = "0|1|"
            frmArtADV.CodigoActual = Text1(Indice).Text
            frmArtADV.Show vbModal
            Set frmArtADV = Nothing
            PonerFoco Text1(Indice)
        
        Case 131 ' codigo de carta de reclamacion (solo utxera y escalona)
            Indice = Index
            
            Set frmCarSocio = New frmCartasSocio
            frmCarSocio.DatosADevolverBusqueda = "0|1|"
            frmCarSocio.CodigoActual = Text1(131).Text
            frmCarSocio.Show vbModal
            Set frmCarSocio = Nothing
            PonerFoco Text1(Indice)
        
        
    End Select

    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.Data1, 1

End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            Indice = 38
            frmZ.pTitulo = "Texto para Cabecera de Toma de Datos"
            frmZ.pValor = Text1(Indice).Text
            frmZ.pModo = Modo
        
            frmZ.Show vbModal
            Set frmZ = Nothing
                
            PonerFoco Text1(Indice)
        Case 1
            Indice = 39
            frmZ.pTitulo = "Texto para Pie de Toma de Datos"
            frmZ.pValor = Text1(Indice).Text
            frmZ.pModo = Modo
        
            frmZ.Show vbModal
            Set frmZ = Nothing
                
            PonerFoco Text1(Indice)
            
    End Select
    
End Sub









'Private Sub mnAñadir_Click()
'    If BLOQUEADesdeFormulario(Me) Then BotonAnyadir
'End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub



Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes estaba esto
'    KEYpress (KeyAscii)
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 5: KEYBusqueda KeyAscii, 0 'tipo de iva transporte
            Case 16: KEYBusqueda KeyAscii, 1 'cuenta de diferencias negativas
            Case 17: KEYBusqueda KeyAscii, 2 'cuenta de diferencias positivas
            Case 43: KEYBusqueda KeyAscii, 43 'forma de pago positiva
            Case 44: KEYBusqueda KeyAscii, 44 'forma de pago negativa
            Case 45: KEYBusqueda KeyAscii, 43 'cuenta de retencion
            Case 46: KEYBusqueda KeyAscii, 44 'cuenta de aportacion
            Case 47: KEYBusqueda KeyAscii, 45 'cuenta de banco prevista
            ' **** almazara
            Case 51: KEYBusqueda KeyAscii, 51 'forma de pago positiva almazara
            Case 52: KEYBusqueda KeyAscii, 52 'forma de pago negativa almazara
            Case 53: KEYBusqueda KeyAscii, 43 'cuenta de retencion almazara
            Case 54: KEYBusqueda KeyAscii, 54 'cuenta de banco prevista almazara
            Case 49: KEYBusqueda KeyAscii, 49 'cuenta de ventas almazara
            Case 50: KEYBusqueda KeyAscii, 50 'cuenta de gastos almazara
            Case 58: KEYBusqueda KeyAscii, 58 'cuenta de banco prevista adv
            Case 59: KEYBusqueda KeyAscii, 59 'cuenta de banco prevista bodega
            Case 69: KEYBusqueda KeyAscii, 69 'cuenta de ventas de bodega
            Case 70: KEYBusqueda KeyAscii, 70 'cuenta de ventas de telefonia
        
            Case 60: KEYBusqueda KeyAscii, 60 'seccion de suministros
            
            Case 63: KEYBusqueda KeyAscii, 10 'seccion de bodega
            
            Case 10:  KEYBusqueda KeyAscii, 6 'carpeta aridoc anticipos
            Case 11: KEYBusqueda KeyAscii, 7 'carpeta aridoc liquidacion
            Case 61: KEYBusqueda KeyAscii, 8 'carpeta aridoc adv
            Case 67: KEYBusqueda KeyAscii, 11 'carpeta aridoc almazara
            Case 68: KEYBusqueda KeyAscii, 12 'carpeta aridoc bodega
            Case 73: KEYBusqueda KeyAscii, 14 'carpeta aridoc recibos almacen
            Case 74: KEYBusqueda KeyAscii, 15 'carpeta aridoc recibos campo
        
            Case 72: KEYBusqueda KeyAscii, 13 'codigo de almacen de nominas
            Case 76: KEYBusqueda KeyAscii, 16 'codigo de concepto de gasto de nominas
            
            Case 112: KEYBusqueda KeyAscii, 21 'codigo de concepto de gasto
            
            ' pozos
            Case 90: KEYBusqueda KeyAscii, 17 'codigo de iva de pozos
            Case 122: KEYBusqueda KeyAscii, 122 'cuenta de ventas consumo pozos
            Case 123: KEYBusqueda KeyAscii, 123 'cuenta de ventas cuotas pozos
            Case 129: KEYBusqueda KeyAscii, 129 'cuenta de ventas talla pozos
            Case 130: KEYBusqueda KeyAscii, 130 'cuenta de ventas mto pozos
            Case 134: KEYBusqueda KeyAscii, 134 'cuenta de ventas mto pozos
            '[Monica]21/01/2016: cuenta de recargos
            Case 135: KEYBusqueda KeyAscii, 135 'cuenta de recargos
            
            Case 126: KEYBusqueda KeyAscii, 126 'forma de pago contado pozos
            Case 127: KEYBusqueda KeyAscii, 127 'forma de pago recibo pozos
            
            Case 93: KEYBusqueda KeyAscii, 18 ' tarifa local 1 de transportes
            Case 113: KEYBusqueda KeyAscii, 22 ' tarifa local 2 de transportes
            Case 94: KEYBusqueda KeyAscii, 19 ' cod.gasto de transportes
            
            Case 96: KEYBusqueda KeyAscii, 20 'carpeta aridoc transporte
            
            Case 114: KEYBusqueda KeyAscii, 23 'codigo de iva de facturas internas de adv
            
            Case 131: KEYBusqueda KeyAscii, 131 'codigo de cartas reclamacion de socio (solo utxera y escalona)
            
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim cad As String

    If Text1(Index).Text = "" Then Exit Sub
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    'Quitar espacios en blanco
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    Select Case Index
        Case 4 'numero de contabilidad
            If Not EsNumerico(Text1(Index).Text) Then
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            Else
                cmdAceptar_Click
            End If
            
            
        Case 10, 11, 61, 67, 68, 73, 74, 96
            If Text1(Index).Text = "" Then Exit Sub
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            If ComprobarCero(Text1(Index)) <> 0 Then
                cad = CargaPath(Text1(Index))
                Text2(Index).Text = Mid(cad, 2, Len(cad))
            End If
        
        Case 13
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "extension", "descripcion", "codext", "N", cAridoc)
        
        Case 14, 15, 16, 17, 18
            If Text1(Index).Text = "" Then Exit Sub
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        
        Case 19, 20, 21, 22, 23 'peso cajas
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoDecimal Text1(Index), 4
       
        Case 103, 104, 105, 106, 107 'peso cajas transportistas
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoDecimal Text1(Index), 4
       
       
        Case 64, 65, 66 ' limite inferior y superior de kilos caja
                        ' 66 peso caja llena
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoDecimal Text1(Index), 4
       
        Case 109 ' porcentaje de incremento / decremento de aforo
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
        
        Case 25, 26 'coste hora y coste seguridad social
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoDecimal Text1(Index), 8
        
        Case 27 ' codigo de seccion hortofruticola
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rseccion", "nomsecci", "codsecci", "N", cAgro)
        
        Case 48 ' codigo de seccion almazara
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rseccion", "nomsecci", "codsecci", "N", cAgro)
        
        Case 56 ' codigo de seccion adv
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rseccion", "nomsecci", "codsecci", "N", cAgro)
        
        
        Case 28, 37 ' porcentaje de retencion de facturas socios
                    ' porcentaje de aportacion de fondo operativo
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
            
        Case 110 ' constante faneca
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 6
            
        Case 38 ' texto de toma de datos
            
        
        Case 40 ' codigo iva intracomunitario
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    If PonerFormatoEntero(Text1(Index)) Then
                        Text2(Index).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(Index), "N")
                    Else
                        Text2(Index).Text = ""
                    End If
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
        Case 43, 44 ' forma de pago en positivo y en negativo
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    If PonerFormatoEntero(Text1(Index)) Then
                        If vParamAplic.ContabilidadNueva Then
                            Text2(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", Text1(Index), "N")
                        Else
                            Text2(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(Index), "N")
                        End If
                    Else
                        Text2(Index).Text = ""
                    End If
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
            
        Case 45, 46, 47, 70, 42, 117 ' cuentas contables de retencion aportacion y banco
                            ' para contabilizacion de facturas de socio
                            ' 70 cta de ventas de telefonia
                            ' 42 cta de retencion de terceros
                            ' 117 cta de retencion de transporte
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    Text2(Index).Text = PonerNombreCuenta(Text1(Index), 2)
                    If Text2(Index).Text = "" Then PonerFoco Text1(Index)
' antes
'                    If PonerFormatoEntero(Text1(Index)) Then
'                        Text2(Index).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(Index), "T")
'                    Else
'                        Text2(Index).Text = ""
'                    End If
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
       Case 71 ' letra de serie de telefonia
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
            
    ' ***********ALMAZARA*********
        Case 51, 52 ' forma de pago en positivo y en negativo
            If Text1(48).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(48).Text) Then
                If vSeccion.AbrirConta Then
                    If PonerFormatoEntero(Text1(Index)) Then
                        If vParamAplic.ContabilidadNueva Then
                            Text2(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", Text1(Index), "N")
                        Else
                            Text2(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(Index), "N")
                        End If
                    Else
                        Text2(Index).Text = ""
                    End If
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
            
        Case 53, 54, 49, 50
            ' 53 cuenta contable de retencion almazara
            ' 54 cuenta banco almazara
            ' 49 cuenta ventas almazara
            ' 50 cuenta gastos almazara
            ' para contabilizacion de facturas de socio
            
            If Text1(48).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(48).Text) Then
                If vSeccion.AbrirConta Then
                    Text2(Index).Text = PonerNombreCuenta(Text1(Index), 2)
                    If Text2(Index).Text = "" Then PonerFoco Text1(Index)
' antes
'                    If PonerFormatoEntero(Text1(Index)) Then
'                        Text2(Index).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(Index), "T")
'                    Else
'                        Text2(Index).Text = ""
'                    End If
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
       Case 112 ' concepto de gasto para el reparto de gastos de la liquidacion
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rconcepgasto", "nomgasto", "codgasto", "N", cAgro)
            
            
       Case 115 ' precio por kilo de gto de molturacion
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 7
            
       Case 116 ' precio por litro de gto de envasado
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 7
            
    ' ***********END ALMAZARA*********
        
        
    ' ***********ADV*********
       Case 57 ' almacen de adv
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "salmpr", "nomalmac", "codalmac", "N", cAgro)
        
        Case 58
            ' 58 cuenta contable de banco adv
            ' para contabilizacion de facturas de adv
            If Text1(56).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(56).Text) Then
                If vSeccion.AbrirConta Then
                    Text2(Index).Text = PonerNombreCuenta(Text1(Index), 2)
                    If Text2(Index).Text = "" Then PonerFoco Text1(Index)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
    
    
       Case 114 ' codigo de iva de exento factura de adv
            If Text1(56).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(56).Text) Then
                If vSeccion.AbrirConta Then
                    If PonerFormatoEntero(Text1(Index)) Then
                        Text2(Index).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(Index), "N")
                    Else
                        Text2(Index).Text = ""
                    End If
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
    
            
    
    ' ***********END ADV*********
        
    ' ***********BODEGA*********
        Case 59, 69
            ' 59 cuenta contable de banco bodega
            ' 69 cuenta contable VENTAS bodega
            ' para contabilizacion de facturas de bodega
            If Text1(63).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(63).Text) Then
                If vSeccion.AbrirConta Then
                    Text2(Index).Text = PonerNombreCuenta(Text1(Index), 2)
                    If Text2(Index).Text = "" Then PonerFoco Text1(Index)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
       
       Case 75 ' porcentaje de gastos de mantenimiento en liqudiacion de bodega
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
       
       Case 60 ' seccion de suministros
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rseccion", "nomsecci", "codsecci", "N", cAgro)
    
       Case 63 ' seccion de bodega
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rseccion", "nomsecci", "codsecci", "N", cAgro)
    
       Case 76 ' concepto de gasto para el reparto de gastos de la liquidacion
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rconcepgasto", "nomgasto", "codgasto", "N", cAgro)
        
       Case 128 ' porcentaje de incremento de kilos de entrada e bodega
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
    
    
    ' ***********END BODEGA*********
       
    
    ' ***********NOMINAS*********
       Case 72 ' almacen de nominas
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "salmpr", "nomalmac", "codalmac", "N", cAgro)
    
       Case 97 'euros trabajador dia
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoDecimal Text1(Index), 3
    
       Case 98, 99, 100, 101 ' porcentajes de seguridad social irpf y de jornada
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
    
       Case 108
            If Text1(Index).Text <> "" Then PonerFormatoEntero Text1(Index)
    
    ' ***********POZOS***********
       Case 86, 88, 138 ' hasta rangos m3
            PonerFormatoEntero Text1(Index)
            
       Case 87, 89, 137  ' precio rangos
            PonerFormatoDecimal Text1(Index), 10
            
       Case 90 ' codigo de iva de pozos
            If Text1(121).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(121).Text) Then
                If vSeccion.AbrirConta Then
                    If PonerFormatoEntero(Text1(Index)) Then
                        Text2(Index).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(Index), "N")
                    Else
                        Text2(Index).Text = ""
                    End If
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
       Case 91, 92, 139, 140, 141 ' cuota y derrama de pozos y canon de contador coeficiente consumo y coef.suministro externo
            PonerFormatoDecimal Text1(Index), 3
    
        Case 122, 123, 129, 130, 134, 135
            ' 122 cuenta contable de ventas consumo pozo
            ' 123 cuenta contable de ventas cuotas pozos
            ' 129 cuenta contable de ventas talla pozos
            ' 130 cuenta contable de ventas mantenimiento pozos
            ' 134 cuenta contable de ventas de consumo a manta
            ' 135 cuenta contable de recargos
            
            If Text1(121).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(121).Text) Then
                If vSeccion.AbrirConta Then
                    Text2(Index).Text = PonerNombreCuenta(Text1(Index), 2)
                    If Text2(Index).Text = "" Then PonerFoco Text1(Index)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
        
        Case 126, 127 ' forma de pago contado y recibo de pozos
            If Text1(121).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(121).Text) Then
                If vSeccion.AbrirConta Then
                    If PonerFormatoEntero(Text1(Index)) Then
                        If vParamAplic.ContabilidadNueva Then
                            Text2(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", Text1(Index), "N")
                        Else
                            Text2(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(Index), "N")
                        End If
                    Else
                        Text2(Index).Text = ""
                    End If
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
        Case 131 ' codigo de carta de reclamacion
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "scartas", "descarta", "codcarta", "N", cAgro)
            
        '[Monica]11/06/2013:
        Case 132, 133 ' consumo minimo y consumo maximo
            PonerFormatoEntero Text1(Index)
            
    ' ***********TRANSPORTE***********
        Case 93, 113 ' 93= tarifa local 1 de transporte
                     ' 113= tarifa local 2 de transporte
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rtarifatra", "nomtarif", "codtarif", "N", cAgro)
        
        Case 94 ' concepto de transporte
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rconcepgasto", "nomgasto", "codgasto", "N", cAgro)
    
        Case 102 ' porcentaje de retencion de modulo transportista
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
    
        Case 111 ' precio por kilo de transporte
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 7
            
        '[Monica]18/03/2018: euros capital social
        Case 142 ' euros por superficie cultivable = capital social
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 10
            
            
    End Select
End Sub


Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
'    Select Case Index
'        Case 6, 7
'            If Text1(Index).Text <> "" Then
'                If Not EsNumerico(Text1(Index).Text) Then
'                    Cancel = True
'                    ConseguirFoco Text1(Index), Modo
'                End If
'            End If
'    End Select
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
'        Case 1  'Anyadir
'            BotonAnyadir
        Case 2  'Modificar
            mnModificar_Click
        Case 5 'Salir
            mnSalir_Click
    End Select
End Sub


'Private Sub BotonAnyadir()
'    LimpiarCampos
'    PonerModo 3
'    Text1(0).Text = 1
'    PonerFoco Text1(1)
'End Sub


Private Sub BotonModificar()
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim i As Integer
Dim NumTaras As Integer

    DatosOk = False
    B = CompForm(Me)
    
    '[Monica]19/12/2011: solo pueden haber 2 y solo 2 taras de esvtafruta marcadas
    If B Then
        NumTaras = 0
        For i = 0 To 4
            If Me.ChkVtaFruta(i).Value Then NumTaras = NumTaras + 1
        Next i
        
        If NumTaras <> 2 Then
            MsgBox "Debe haber marcadas 2 y sólo 2 tipos de caja de Venta Fruta. Revise.", vbExclamation
            B = False
        End If
    End If
    
    DatosOk = B
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(B As Boolean)
    Me.cmdAceptar.visible = Not B
    Me.CmdCancelar.visible = Not B
    Me.cmdSalir.visible = B
'    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
Dim i As Byte
Dim cad As String


On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    ' ************* si hay aridoc **************
    If vParamAplic.HayAridoc = 1 Then
         If ComprobarCero(Text1(10).Text) <> 0 Then
            cad = CargaPath(Text1(10))
            Text2(10).Text = Mid(cad, 2, Len(cad))
         End If
         If ComprobarCero(Text1(11).Text) <> 0 Then
            cad = CargaPath(Text1(11))
            Text2(11).Text = Mid(cad, 2, Len(cad))
         End If
         If ComprobarCero(Text1(61).Text) <> 0 Then
            cad = CargaPath(Text1(61))
            Text2(61).Text = Mid(cad, 2, Len(cad))
         End If
         If ComprobarCero(Text1(67)) <> 0 Then
            cad = CargaPath(Text1(67))
            Text2(67).Text = Mid(cad, 2, Len(cad))
         End If
         If ComprobarCero(Text1(68)) <> 0 Then
            cad = CargaPath(Text1(68))
            Text2(68).Text = Mid(cad, 2, Len(cad))
         End If
         If ComprobarCero(Text1(73)) <> 0 Then
            cad = CargaPath(Text1(73))
            Text2(73).Text = Mid(cad, 2, Len(cad))
         End If
         If ComprobarCero(Text1(74)) <> 0 Then
            cad = CargaPath(Text1(74))
            Text2(74).Text = Mid(cad, 2, Len(cad))
         End If
         If ComprobarCero(Text1(96)) <> 0 Then
            cad = CargaPath(Text1(96))
            Text2(96).Text = Mid(cad, 2, Len(cad))
         End If

         Text2(13).Text = DevuelveDesdeBDNew(cAridoc, "extension", "descripcion", "codext", Text1(13).Text, "N")
    End If
    
    ' ************* configurar els camps de les descripcions de les comptes *************
    If Text1(27).Text <> "" Then  ' si no hemos indicado la seccion
                                  ' no sabemos a que contabilidad va
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Text1(27).Text) Then
            If vSeccion.AbrirConta Then
                ' porcentaje de iva de terceros
                If PonerFormatoEntero(Text1(40)) Then
                    Text2(40).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(40), "N")
                Else
                    Text2(40).Text = ""
                End If
                ' cuenta de retencion de terceros
                If Text1(42).Text <> "" Then
                    Text2(42).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(42), "T")
                End If
                
                ' cuenta de retencion de transportista
                If Text1(117).Text <> "" Then
                    Text2(117).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(117), "T")
                End If
                
                ' cuenta de retencion de facturas de socios
                If Text1(45).Text <> "" Then
                    Text2(45).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(45), "T")
                End If
                ' cuenta de aportacion de facturas de socios
                If Text1(46).Text <> "" Then
                    Text2(46).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(46), "T")
                End If
                ' cuenta de prevista de banco de facturas de socios
                If Text1(47).Text <> "" Then
                    Text2(47).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(47), "T")
                End If
                
                ' forma de pago de facturas anticipos / liquidaciones de socios positivas
                If Text1(43).Text <> "" Then
                    If vParamAplic.ContabilidadNueva Then
                        Text2(43).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", Text1(43), "N")
                    Else
                        Text2(43).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(43), "N")
                    End If
                End If
                ' forma de pago de facturas anticipos / liquidaciones de socios negativas
                If Text1(44).Text <> "" Then
                    If vParamAplic.ContabilidadNueva Then
                        Text2(44).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", Text1(44), "N")
                    Else
                        Text2(44).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(44), "N")
                    End If
                End If
                
                ' telefonia de valsur
                ' cuenta de ventas de telefonia
                If Text1(70).Text <> "" Then
                    Text2(70).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(70), "T")
                End If
                
                
            End If
            vSeccion.CerrarConta
        End If
        Set vSeccion = Nothing
    
        Text2(27).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(27).Text, "N")
    End If
    
    
    If Text1(121).Text <> "" Then  ' si no hemos indicado la seccion
                                  ' no sabemos a que contabilidad va
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Text1(121).Text) Then
            If vSeccion.AbrirConta Then
                ' pozos
                ' codigo de iva de pozos
                If Text1(90).Text <> "" Then
                    Text2(90).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(90), "N")
                End If
                ' cuenta de consumo de pozos
                If Text1(122).Text <> "" Then
                    Text2(122).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(122), "T")
                End If
                ' cuenta de cuotas de pozos
                If Text1(123).Text <> "" Then
                    Text2(123).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(123), "T")
                End If
                ' cuenta de talla de pozos
                If Text1(129).Text <> "" Then
                    Text2(129).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(129), "T")
                End If
                ' cuenta de mantenimiento de pozos
                If Text1(130).Text <> "" Then
                    Text2(130).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(130), "T")
                End If
                ' cuenta de ventas de consumo a manta de pozos
                If Text1(134).Text <> "" Then
                    Text2(134).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(134), "T")
                End If
                '[Monica]21/01/2016: recargos de escalona
                ' cuenta de recargos de pozos
                If Text1(135).Text <> "" Then
                    Text2(135).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(135), "T")
                End If
                
                ' centro de coste de pozos
                If Text1(124).Text <> "" Then
                    If vParamAplic.ContabilidadNueva Then
                        Text2(124).Text = DevuelveDesdeBDNew(cConta, "ccoste", "nomccost", "codccost", Text1(124), "T")
                    Else
                        Text2(124).Text = DevuelveDesdeBDNew(cConta, "cabccost", "nomccost", "codccost", Text1(124), "T")
                    End If
                End If
                ' forma de pago de contado de pozos
                If Text1(126).Text <> "" Then
                    If vParamAplic.ContabilidadNueva Then
                        Text2(126).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", Text1(126), "N")
                    Else
                        Text2(126).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(126), "N")
                    End If
                End If
                ' forma de pago de recibos de pozos
                If Text1(127).Text <> "" Then
                    If vParamAplic.ContabilidadNueva Then
                        Text2(127).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", Text1(127), "N")
                    Else
                        Text2(127).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(127), "N")
                    End If
                End If
            End If
            vSeccion.CerrarConta
        End If
        Set vSeccion = Nothing
    
        Text2(121).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(121).Text, "N")
        Text2(131).Text = DevuelveDesdeBDNew(cAgro, "scartas", "descarta", "codcarta", Text1(131).Text, "N")
    End If
    
    
    
    If Text1(48).Text <> "" Then ' si no hemos indicado la seccion
                                 ' no sabemos a que contabilidad va
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Text1(48).Text) Then
            If vSeccion.AbrirConta Then
                ' cuenta de retencion de facturas de almazara
                If Text1(53).Text <> "" Then
                    Text2(53).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(53), "T")
                End If
                ' cuenta de prevista de banco de facturas de almazara
                If Text1(54).Text <> "" Then
                    Text2(54).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(54), "T")
                End If
                
                ' forma de pago de facturas almazara positivas y negativas
                If Text1(51).Text <> "" Then
                    If vParamAplic.ContabilidadNueva Then
                        Text2(51).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", Text1(51), "N")
                    Else
                        Text2(51).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(51), "N")
                    End If
                End If
                ' forma de pago de facturas anticipos / liquidaciones de socios negativas
                If Text1(52).Text <> "" Then
                    If vParamAplic.ContabilidadNueva Then
                        Text2(52).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", Text1(52), "N")
                    Else
                        Text2(52).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(52), "N")
                    End If
                End If
                
                ' cuenta de ventas de facturas de almazara
                If Text1(49).Text <> "" Then
                    Text2(49).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(49), "T")
                End If
                ' cuenta de gastos de facturas de almazara
                If Text1(50).Text <> "" Then
                    Text2(50).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(50), "T")
                End If
            
            End If
            vSeccion.CerrarConta
        End If
        Set vSeccion = Nothing
        Text2(48).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(48).Text, "N")
    End If
    
    If Text1(56).Text <> "" Then ' si no hemos indicado la seccion
                                 ' no sabemos a que contabilidad va
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Text1(56).Text) Then
            If vSeccion.AbrirConta Then
                ' codigo de iva de facturas internas de adv
                If Text1(114).Text <> "" Then
                    Text2(114).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(114), "N")
                End If
            End If
        End If
    End If
    
    
    ' almacen de adv
    Text2(57).Text = DevuelveDesdeBDNew(cAgro, "salmpr", "nomalmac", "codalmac", Text1(57).Text, "N")
        
    
    
    If Text1(56).Text <> "" Then  ' si no hemos indicado la seccion
                                  ' no sabemos a que contabilidad va
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Text1(56).Text) Then
            If vSeccion.AbrirConta Then
                ' cuenta de prevista de banco de facturas de adv
                If Text1(58).Text <> "" Then
                    Text2(58).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(58), "T")
                End If
            End If
            vSeccion.CerrarConta
        End If
        Set vSeccion = Nothing
    
        Text2(56).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(56).Text, "N")
    End If
    
    ' seccion de suministros
    Text2(60).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(60).Text, "N")
    
    ' seccion de bodega
    
    If Text1(63).Text <> "" Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Text1(63).Text) Then
            If vSeccion.AbrirConta Then
                ' cuenta de banco prevista
                If Text1(59).Text <> "" Then
                    Text2(59).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(59), "T")
                End If
                ' cuenta de ventas de bodega
                If Text1(69).Text <> "" Then
                    Text2(69).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(69), "T")
                End If
            End If
        End If
        
        Text2(63).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(63).Text, "N")
    End If
    
    If Text1(76).Text <> "" Then ' codigo de gasto para el reparto de gasto de liquidacion bodega
        Text2(76).Text = DevuelveDesdeBDNew(cAgro, "rconcepgasto", "nomgasto", "codgasto", Text1(76).Text, "N")
    End If
    
    ' codigo de almacen de gestion de nominas
    If Text1(72).Text <> "" Then
        Text2(72).Text = DevuelveDesdeBDNew(cAgro, "salmpr", "nomalmac", "codalmac", Text1(72).Text, "N")
    End If
    
    ' TRASNPORTE
    
    ' codigo de tarifa de transporte local
    If Text1(93).Text <> "" Then
        Text2(93).Text = DevuelveDesdeBDNew(cAgro, "rtarifatra", "nomtarif", "codtarif", Text1(93).Text, "N")
    End If
    ' codigo de tarifa 2 de transporte local
    If Text1(113).Text <> "" Then
        Text2(113).Text = DevuelveDesdeBDNew(cAgro, "rtarifatra", "nomtarif", "codtarif", Text1(113).Text, "N")
    End If
        
    
    ' concepto de gasto de transporte
    If Text1(94).Text <> "" Then
        Text2(94).Text = DevuelveDesdeBDNew(cAgro, "rconcepgasto", "nomgasto", "codgasto", Text1(94).Text, "N")
    End If
    
    ' concepto de gasto de almazara
    If Text1(112).Text <> "" Then ' codigo de gasto para el reparto de gasto de liquidacion bodega
        Text2(112).Text = DevuelveDesdeBDNew(cAgro, "rconcepgasto", "nomgasto", "codgasto", Text1(112).Text, "N")
    End If
    
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    For i = 1 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
    Next i
    '### a mano
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
Dim i As Byte
Dim vtag As CTag

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
      
    '------------------------------------------------------
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    If CmdCancelar.visible Then
        CmdCancelar.Cancel = True
    Else
        CmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not B
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
'    BloquearCombo Me, Modo
    
    For i = 0 To 29
            Set vtag = New CTag
            vtag.Cargar Me.Combo1(i)
            If vtag.Cargado Then
                If vtag.EsClave And (Modo = 4 Or Modo = 5) Then
                    Me.Combo1(i).Enabled = False
                    Me.Combo1(i).BackColor = &H80000018 'groc
                Else
                    Me.Combo1(i).Enabled = B
                    If B Then
                        Me.Combo1(i).BackColor = vbWhite
                    Else
                        Me.Combo1(i).BackColor = &H80000018 'Amarillo Claro
                    End If
                    If Modo = 3 Then Me.Combo1(i).ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            End If
     Next i

    ' no se pueden modificar la primera y ultima factura de ultima facturaciones
    Frame5.Enabled = False
    Frame15.Enabled = False
    Frame16.Enabled = False
    
    'Bloquear imagen de Busqueda
    For i = 6 To 8
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    For i = 9 To 25
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    For i = 0 To 5
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    For i = 43 To 47
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    For i = 49 To 54
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    For i = 58 To 60
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    For i = 69 To 70
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    For i = 122 To 124
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    For i = 126 To 127
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    For i = 129 To 131
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    
    For i = 134 To 135
        Me.imgBuscar(i).Enabled = (Modo >= 3)
        Me.imgBuscar(i).visible = (Modo >= 3)
    Next i
    
    
'    BloquearImgBuscar Me, Modo
    'Bloquear los checkbox
    BloquearChecks Me, Modo
    
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
    B = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not Encontrado And Not B  'Añadir
    Me.Toolbar1.Buttons(2).Enabled = Encontrado And Not B 'Modificar
    Me.mnAñadir.Enabled = Not Encontrado And Not B
    Me.mnModificar.Enabled = Encontrado And Not B
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
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
    
    ' combo de tipo de transporte
    Combo1(0).AddItem "Portes por Población"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Tarifas de Transporte"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    ' combo de tipo de transporte
    Combo1(29).AddItem "Contador Global"
    Combo1(29).ItemData(Combo1(29).NewIndex) = 0
    Combo1(29).AddItem "Contador por Transportista"
    Combo1(29).ItemData(Combo1(29).NewIndex) = 1
    
    'combos de anticipos
    For i = 1 To 4
        Combo1(i).AddItem "Nro.Factura"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 0
        Combo1(i).AddItem "Cod.Socio"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 1
        Combo1(i).AddItem "Nom.Socio"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 2
        Combo1(i).AddItem "Variedad"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 3
    Next i
    
    'combos de liquidacion
    For i = 5 To 8
        Combo1(i).AddItem "Nro.Factura"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 0
        Combo1(i).AddItem "Cod.Socio"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 1
        Combo1(i).AddItem "Nom.Socio"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 2
        Combo1(i).AddItem "Variedad"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 3
    Next i
    
    'combos de adv
    For i = 9 To 12
        Combo1(i).AddItem "Nro.Factura"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 0
        Combo1(i).AddItem "Cod.Socio"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 1
        Combo1(i).AddItem "Nom.Socio"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 2
        Combo1(i).AddItem "Procedencia"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 3
    Next i
    
    'combos de almazara
    For i = 13 To 16
        Combo1(i).AddItem "Nro.Factura"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 0
        Combo1(i).AddItem "Cod.Socio"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 1
        Combo1(i).AddItem "Nom.Socio"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 2
        Combo1(i).AddItem "Variedad"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 3
    Next i
    
    'combos de bodega
    For i = 17 To 20
        Combo1(i).AddItem "Nro.Factura"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 0
        Combo1(i).AddItem "Cod.Socio"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 1
        Combo1(i).AddItem "Nom.Socio"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 2
        Combo1(i).AddItem "Variedad"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 3
    Next i
    
    'combos de recibos de campo
    For i = 21 To 24
        Combo1(i).AddItem "Cod.Trabajador"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 0
        Combo1(i).AddItem "Nom.Trabajador"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 1
        Combo1(i).AddItem "Procedencia"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 2
    Next i
    
    'combos de bodega
    For i = 25 To 28
        Combo1(i).AddItem "Nro.Factura"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 0
        Combo1(i).AddItem "Cod.Trans"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 1
        Combo1(i).AddItem "Nom.Trans"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 2
        Combo1(i).AddItem "Variedad"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 3
    Next i
    
    
    
End Sub

Private Function CargaPath(Codigo As Integer) As String
Dim Nod As Node
Dim J As Integer
Dim i As Integer
Dim c As String
Dim campo1 As String
Dim padre As String
Dim A As String

    'Primero copiamos la carpeta
    c = "\" & DevuelveDesdeBDNew(cAridoc, "carpetas", "nombre", "codcarpeta", CInt(Codigo), "N")
    campo1 = "nombre"
    padre = DevuelveDesdeBDNew(cAridoc, "carpetas", "padre", "codcarpeta", CStr(Codigo), "N", campo1)
    If CInt(ComprobarCero(padre)) > 0 Then
        c = CargaPath(CInt(padre)) & c
    End If
'
'    If No.Children > 0 Then
'        J = No.Children
'        Set Nod = No.Child
'        For i = 1 To J
'           C = C & CopiaArchivosCarpetaRecursiva(Nod)
'           If i <> J Then Set Nod = Nod.Next
'        Next i
'    End If
    CargaPath = c
End Function


Private Sub AbrirFrmForpaConta(indice1 As Integer)
    Indice = indice1
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = Text1(Indice)
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub
