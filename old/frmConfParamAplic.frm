VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfParamAplic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de la Aplicación"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9435
   Icon            =   "frmConfParamAplic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5580
      Left            =   150
      TabIndex        =   80
      Top             =   600
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   9843
      _Version        =   393216
      Tabs            =   10
      Tab             =   9
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Contabilidad"
      TabPicture(0)   =   "frmConfParamAplic.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Internet"
      TabPicture(1)   =   "frmConfParamAplic.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Entradas"
      TabPicture(2)   =   "frmConfParamAplic.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkRespetarNroNota"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text1(66)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Text1(65)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Text1(64)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkAgruparNotas"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text1(31)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text1(24)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "chkTraza"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chkTaraTractor"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame3"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label22"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label21"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label20"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label19"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label14"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label11"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Aridoc"
      TabPicture(3)   =   "frmConfParamAplic.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(6)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "imgBuscar(6)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label1(7)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "imgBuscar(7)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label1(28)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "imgBuscar(9)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "imgBuscar(8)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label1(47)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Text2(10)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Text1(10)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Text2(11)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Text1(11)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Frame8"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Frame9"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Text1(13)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Text2(13)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Text1(61)"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Text2(61)"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "Frame11"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).ControlCount=   19
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
      Tab(4).Control(11)=   "Text1(25)"
      Tab(4).Control(12)=   "Text1(26)"
      Tab(4).Control(13)=   "Text1(27)"
      Tab(4).Control(14)=   "Text2(27)"
      Tab(4).Control(15)=   "Text1(28)"
      Tab(4).Control(16)=   "Frame5"
      Tab(4).Control(17)=   "Text1(37)"
      Tab(4).Control(18)=   "Text1(38)"
      Tab(4).Control(19)=   "Text1(39)"
      Tab(4).Control(20)=   "Text1(41)"
      Tab(4).ControlCount=   21
      TabCaption(5)   =   "Terceros"
      TabPicture(5)   =   "frmConfParamAplic.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text1(40)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Text2(40)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Text1(42)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Text2(42)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "imgBuscar(1)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label1(5)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "imgBuscar(2)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Label1(13)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).ControlCount=   8
      TabCaption(6)   =   "Almazara"
      TabPicture(6)   =   "frmConfParamAplic.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label1(34)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "imgBuscar(3)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Frame10"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Text2(48)"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Text1(48)"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).ControlCount=   5
      TabCaption(7)   =   "ADV"
      TabPicture(7)   =   "frmConfParamAplic.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label1(36)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "imgBuscar(4)"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "imgBuscar(5)"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Label1(42)"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Label1(44)"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "imgBuscar(58)"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "Text2(56)"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "Text1(56)"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "Text1(57)"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).Control(9)=   "Text2(57)"
      Tab(7).Control(9).Enabled=   0   'False
      Tab(7).Control(10)=   "Text2(58)"
      Tab(7).Control(10).Enabled=   0   'False
      Tab(7).Control(11)=   "Text1(58)"
      Tab(7).Control(11).Enabled=   0   'False
      Tab(7).ControlCount=   12
      TabCaption(8)   =   "Suministros"
      TabPicture(8)   =   "frmConfParamAplic.frx":00EC
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label1(46)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "imgBuscar(60)"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "Label1(52)"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).Control(3)=   "Text2(60)"
      Tab(8).Control(3).Enabled=   0   'False
      Tab(8).Control(4)=   "Text1(60)"
      Tab(8).Control(4).Enabled=   0   'False
      Tab(8).Control(5)=   "Text1(62)"
      Tab(8).Control(5).Enabled=   0   'False
      Tab(8).ControlCount=   6
      TabCaption(9)   =   "Bodega"
      TabPicture(9)   =   "frmConfParamAplic.frx":0108
      Tab(9).ControlEnabled=   -1  'True
      Tab(9).Control(0)=   "imgBuscar(10)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "Label1(53)"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "Label1(45)"
      Tab(9).Control(2).Enabled=   0   'False
      Tab(9).Control(3)=   "imgBuscar(59)"
      Tab(9).Control(3).Enabled=   0   'False
      Tab(9).Control(4)=   "Text1(63)"
      Tab(9).Control(4).Enabled=   0   'False
      Tab(9).Control(5)=   "Text2(63)"
      Tab(9).Control(5).Enabled=   0   'False
      Tab(9).Control(6)=   "ChkContadorManual"
      Tab(9).Control(6).Enabled=   0   'False
      Tab(9).Control(7)=   "Text2(59)"
      Tab(9).Control(7).Enabled=   0   'False
      Tab(9).Control(8)=   "Text1(59)"
      Tab(9).Control(8).Enabled=   0   'False
      Tab(9).ControlCount=   9
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   59
         Left            =   2475
         MaxLength       =   10
         TabIndex        =   198
         Tag             =   "Cta Contable Banco|T|S|||rparam|ctabancobod|||"
         Top             =   2340
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   59
         Left            =   3765
         TabIndex        =   205
         Top             =   2340
         Width           =   3690
      End
      Begin VB.CheckBox ChkContadorManual 
         Caption         =   "Contador de Albarán de Retirada manual "
         Height          =   375
         Left            =   405
         TabIndex        =   197
         Tag             =   "Contador albaran Retirada Manual|N|S|||rparam|albretiradabodman|0||"
         Top             =   1620
         Width           =   3405
      End
      Begin VB.CheckBox chkRespetarNroNota 
         Caption         =   "Se respeta Nro.de Nota"
         Height          =   375
         Left            =   -74520
         TabIndex        =   32
         Tag             =   "Se Respeta Nro.Notas|N|N|0|1|rparam|serespetanota|0||"
         Top             =   4660
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   -67500
         MaxLength       =   6
         TabIndex        =   35
         Tag             =   "Peso Caja Llena|N|S|||rparam|pesocajallena|##0.00||"
         Text            =   "KgCajo"
         Top             =   4020
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   -68850
         MaxLength       =   6
         TabIndex        =   37
         Tag             =   "Kilos Caja Máximo|N|N|||rparam|kiloscajamax|##0.00||"
         Text            =   "kgmax"
         Top             =   4800
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   64
         Left            =   -70110
         MaxLength       =   6
         TabIndex        =   36
         Tag             =   "Kilos Caja Mínimo|N|N|||rparam|kiloscajamin|##0.00||"
         Text            =   "kgmin"
         Top             =   4800
         Width           =   1170
      End
      Begin VB.CheckBox chkAgruparNotas 
         Caption         =   "Se agrupan notas"
         Height          =   375
         Left            =   -74520
         TabIndex        =   31
         Tag             =   "Se Agrupan Notas|N|S|||rparam|agruparnotas|0||"
         Top             =   4310
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   63
         Left            =   3060
         TabIndex        =   199
         Top             =   1050
         Width           =   3690
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   63
         Left            =   2430
         MaxLength       =   3
         TabIndex        =   196
         Tag             =   "Sección Bodega|N|N|||rparam|seccionbodega|000||"
         Top             =   1050
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   62
         Left            =   -72570
         MaxLength       =   10
         TabIndex        =   193
         Tag             =   "BD Ariges|T|S|||rparam|bdariges|||"
         Top             =   1470
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   -72570
         MaxLength       =   3
         TabIndex        =   192
         Tag             =   "Sección Suministros|N|N|||rparam|seccionsumi|000||"
         Top             =   1050
         Width           =   585
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   60
         Left            =   -71940
         TabIndex        =   191
         Top             =   1050
         Width           =   3690
      End
      Begin VB.Frame Frame11 
         Caption         =   "ADV"
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
         Height          =   1050
         Left            =   -74460
         TabIndex        =   186
         Top             =   4320
         Width           =   7710
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   9
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Tag             =   "C1 ADV|N|N|||rparam|c1advaridoc||N|"
            Top             =   600
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   10
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Tag             =   "C2 ADV|N|N|||rparam|c2advaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   11
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Tag             =   "C3 ADV|N|N|||rparam|c3advaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   12
            Left            =   5805
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Tag             =   "C4 ADV|N|N|||rparam|c4advaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 4"
            Height          =   195
            Index           =   51
            Left            =   5805
            TabIndex        =   190
            Top             =   315
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 3"
            Height          =   195
            Index           =   50
            Left            =   3915
            TabIndex        =   189
            Top             =   315
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 2"
            Height          =   195
            Index           =   49
            Left            =   1980
            TabIndex        =   188
            Top             =   315
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 1"
            Height          =   195
            Index           =   48
            Left            =   90
            TabIndex        =   187
            Top             =   315
            Width           =   1620
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   61
         Left            =   -71190
         TabIndex        =   184
         Top             =   1350
         Width           =   4470
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   -72480
         MaxLength       =   10
         TabIndex        =   183
         Tag             =   "Carpeta Facturas|N|N|||rparam|codcarpetaADV|000||"
         Top             =   1350
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   -72525
         MaxLength       =   10
         TabIndex        =   73
         Tag             =   "Cta Contable Banco|T|S|||rparam|ctabancoadv|||"
         Top             =   2340
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   58
         Left            =   -71220
         TabIndex        =   181
         Top             =   2340
         Width           =   3690
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   57
         Left            =   -72210
         TabIndex        =   179
         Top             =   1410
         Width           =   4890
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   57
         Left            =   -72840
         MaxLength       =   10
         TabIndex        =   72
         Tag             =   "Almacen ADV|N|N|||rparam|codalmacadv|000||"
         Top             =   1410
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   56
         Left            =   -72840
         MaxLength       =   10
         TabIndex        =   71
         Tag             =   "Sección ADV|N|N|||rparam|seccionadv|000||"
         Top             =   1020
         Width           =   585
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   56
         Left            =   -72210
         TabIndex        =   177
         Top             =   1020
         Width           =   4890
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   -72480
         MaxLength       =   10
         TabIndex        =   61
         Tag             =   "Cod.Iva Extranjero|N|N|||rparam|codivaintracom|000||"
         Top             =   870
         Width           =   585
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   40
         Left            =   -71850
         TabIndex        =   174
         Top             =   870
         Width           =   4350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   -72480
         MaxLength       =   10
         TabIndex        =   62
         Tag             =   "Cta Contable Retencion|T|S|||rparam|ctaterreten|||"
         Top             =   1230
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   42
         Left            =   -71190
         TabIndex        =   173
         Top             =   1230
         Width           =   3690
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   41
         Left            =   -72360
         MaxLength       =   50
         TabIndex        =   60
         Tag             =   "Path Traza|T|S|||rparam|directoriotraza|||"
         Top             =   5100
         Width           =   5325
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   48
         Left            =   -72480
         MaxLength       =   10
         TabIndex        =   63
         Tag             =   "Sección Almazara|N|N|||rparam|seccionalmaz|000||"
         Top             =   1020
         Width           =   585
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   48
         Left            =   -71850
         TabIndex        =   170
         Top             =   1020
         Width           =   3690
      End
      Begin VB.Frame Frame10 
         Caption         =   "Liquidaciones Socio"
         Height          =   3285
         Left            =   -74760
         TabIndex        =   156
         Top             =   1470
         Width           =   7665
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   49
            Left            =   3570
            TabIndex        =   162
            Top             =   1920
            Width           =   3690
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   69
            Tag             =   "Cta Contable Gastos Almazara|T|S|||rparam|ctagastosalmz|||"
            Top             =   2310
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   50
            Left            =   3570
            TabIndex        =   161
            Top             =   2310
            Width           =   3690
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   68
            Tag             =   "Cta Contable Ventas Almazara|T|S|||rparam|ctaventasalmz|||"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   54
            Left            =   3570
            TabIndex        =   160
            Top             =   1560
            Width           =   3690
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   54
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   67
            Tag             =   "Cta Contable Banco|T|S|||rparam|ctabancoalmz|||"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   64
            Tag             =   "Forma Pago Positivas|N|S|||rparam|codforpaposalmz|000||"
            Top             =   390
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   51
            Left            =   2910
            TabIndex        =   159
            Top             =   390
            Width           =   4350
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   52
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   65
            Tag             =   "Forma de Pago Negativas|N|S|||rparam|codforpanegalmz|000||"
            Top             =   750
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   52
            Left            =   2910
            TabIndex        =   158
            Top             =   750
            Width           =   4350
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   53
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   66
            Tag             =   "Cta Contable Retencion Socio|T|S|||rparam|ctaretenalmz|||"
            Top             =   1140
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   53
            Left            =   3570
            TabIndex        =   157
            Top             =   1140
            Width           =   3690
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   55
            Left            =   2280
            MaxLength       =   1
            TabIndex        =   70
            Tag             =   "Letra Serie Almazara|T|S|||rparam|letraseriealmz|||"
            Top             =   2700
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Ventas"
            Height          =   195
            Index           =   37
            Left            =   300
            TabIndex        =   169
            Top             =   1980
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   50
            Left            =   2010
            ToolTipText     =   "Buscar cuenta"
            Top             =   2340
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Gastos"
            Height          =   195
            Index           =   38
            Left            =   300
            TabIndex        =   168
            Top             =   2340
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   49
            Left            =   2010
            ToolTipText     =   "Buscar cuenta"
            Top             =   1950
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Banco Prevista"
            Height          =   195
            Index           =   39
            Left            =   300
            TabIndex        =   167
            Top             =   1590
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   54
            Left            =   2010
            ToolTipText     =   "Buscar cuenta"
            Top             =   1590
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   51
            Left            =   2010
            ToolTipText     =   "Buscar forma Pago"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago Positivas"
            Height          =   195
            Index           =   40
            Left            =   300
            TabIndex        =   166
            Top             =   450
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   52
            Left            =   2010
            ToolTipText     =   "Buscar forma pago"
            Top             =   780
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago Negativas"
            Height          =   195
            Index           =   41
            Left            =   300
            TabIndex        =   165
            Top             =   810
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   53
            Left            =   2010
            ToolTipText     =   "Buscar cuenta"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Retención"
            Height          =   195
            Index           =   43
            Left            =   300
            TabIndex        =   164
            Top             =   1200
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "Letra Serie Clientes"
            Height          =   195
            Index           =   35
            Left            =   300
            TabIndex        =   163
            Top             =   2730
            Width           =   1650
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Anticipos / Liquidaciones Socio"
         Height          =   2625
         Left            =   -74520
         TabIndex        =   145
         Top             =   2520
         Width           =   7665
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   47
            Left            =   3570
            TabIndex        =   154
            Top             =   1920
            Width           =   3690
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   47
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta Contable Banco|T|S|||rparam|ctabanco|||"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   4
            Tag             =   "Forma Pago Positivas|N|S|||rparam|codforpaposi|000||"
            Top             =   390
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   43
            Left            =   2910
            TabIndex        =   152
            Top             =   390
            Width           =   4350
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   5
            Tag             =   "Forma de Pago Negativas|N|S|||rparam|codforpanega|000||"
            Top             =   750
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   44
            Left            =   2910
            TabIndex        =   149
            Top             =   750
            Width           =   4350
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   46
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Cta Contable Aportacion|T|S|||rparam|ctaaportasoc|||"
            Top             =   1530
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   46
            Left            =   3570
            TabIndex        =   148
            Top             =   1530
            Width           =   3690
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta Contable Retencion Socio|T|S|||rparam|ctaretensoc|||"
            Top             =   1140
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   45
            Left            =   3570
            TabIndex        =   146
            Top             =   1140
            Width           =   3690
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Banco Prevista"
            Height          =   195
            Index           =   33
            Left            =   300
            TabIndex        =   155
            Top             =   1950
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   47
            Left            =   2010
            ToolTipText     =   "Buscar cuenta"
            Top             =   1950
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   43
            Left            =   2010
            ToolTipText     =   "Buscar forma Pago"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago Positivas"
            Height          =   195
            Index           =   29
            Left            =   300
            TabIndex        =   153
            Top             =   450
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   44
            Left            =   2010
            ToolTipText     =   "Buscar forma pago"
            Top             =   780
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago Negativas"
            Height          =   195
            Index           =   32
            Left            =   300
            TabIndex        =   151
            Top             =   810
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   46
            Left            =   2010
            ToolTipText     =   "Buscar cuenta"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Aportación"
            Height          =   195
            Index           =   30
            Left            =   300
            TabIndex        =   150
            Top             =   1560
            Width           =   1650
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   45
            Left            =   2010
            ToolTipText     =   "Buscar cuenta"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Retención"
            Height          =   195
            Index           =   24
            Left            =   300
            TabIndex        =   147
            Top             =   1200
            Width           =   1650
         End
      End
      Begin VB.TextBox Text1 
         Height          =   585
         Index           =   39
         Left            =   -74580
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Tag             =   "Texto Pie Toma Datos|T|S|||rparam|pietomadatos|||"
         Top             =   4440
         Width           =   7545
      End
      Begin VB.TextBox Text1 
         Height          =   1125
         Index           =   38
         Left            =   -74580
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Tag             =   "Texto Toma Datos|T|S|||rparam|texttomadatos|||"
         Top             =   3030
         Width           =   7545
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   57
         Tag             =   "Porcentaje AFO|N|S|||rparam|porcenafo|##0.00||"
         Top             =   2370
         Width           =   585
      End
      Begin VB.Frame Frame5 
         Caption         =   "Última Facturación"
         Height          =   1725
         Left            =   -71490
         TabIndex        =   134
         Top             =   1230
         Width           =   4845
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   12
            Left            =   3570
            MaxLength       =   10
            TabIndex        =   127
            Tag             =   "Ult.Fact.Liquidación VC|N|S|||rparam|ultfactliqvc|0000000||"
            Top             =   1380
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   36
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   126
            Tag             =   "Prim.Fact.Liquidación VC|N|S|||rparam|primfactliqvc|0000000||"
            Top             =   1380
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3570
            MaxLength       =   10
            TabIndex        =   125
            Tag             =   "Ult.Fact.Liquidación|N|S|||rparam|ultfactliq|0000000||"
            Top             =   1050
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   124
            Tag             =   "Prim.Fact.Liquidación|N|S|||rparam|primfactliq|0000000||"
            Top             =   1050
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   3570
            MaxLength       =   10
            TabIndex        =   123
            Tag             =   "Ult.Fact.Anticipo VC|N|S|||rparam|ultfactantvc|0000000||"
            Top             =   720
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   122
            Tag             =   "Prim.Fact.Anticipo VC|N|S|||rparam|primfactantvc|0000000||"
            Top             =   720
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   29
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   120
            Tag             =   "Prim.Fact.Anticipo|N|S|||rparam|primfactant|0000000||"
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   3570
            MaxLength       =   10
            TabIndex        =   121
            Tag             =   "Ult.Fact.Anticipo|N|S|||rparam|ultfactant|0000000||"
            Top             =   390
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Liquidación Ventas Campo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   330
            TabIndex        =   140
            Top             =   1380
            Width           =   1980
         End
         Begin VB.Label Label1 
            Caption         =   "Liquidación:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   330
            TabIndex        =   139
            Top             =   1050
            Width           =   1830
         End
         Begin VB.Label Label1 
            Caption         =   "Anticipos Ventas Campo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   330
            TabIndex        =   138
            Top             =   750
            Width           =   1830
         End
         Begin VB.Label Label1 
            Caption         =   "Anticipos:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   137
            Top             =   450
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Desde"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   2430
            TabIndex        =   136
            Top             =   150
            Width           =   630
         End
         Begin VB.Label Label1 
            Caption         =   "Hasta"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   3570
            TabIndex        =   135
            Top             =   150
            Width           =   630
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   31
         Left            =   -72720
         MaxLength       =   50
         TabIndex        =   29
         Tag             =   "Impresora Entradas|T|N|||rparam|impresoraentradas|||"
         Top             =   3600
         Width           =   4995
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   -72600
         MaxLength       =   10
         TabIndex        =   56
         Tag             =   "Porcentaje Retención|N|S|||rparam|porcretenfacsoc||##0.00|"
         Top             =   1995
         Width           =   585
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   -71880
         TabIndex        =   130
         Top             =   780
         Width           =   3690
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   -72510
         MaxLength       =   10
         TabIndex        =   53
         Tag             =   "Sección Hortofrutícola|N|N|||rparam|seccionhorto|000||"
         Top             =   780
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   -73170
         MaxLength       =   8
         TabIndex        =   55
         Tag             =   "Coste seg.soc|N|N|||rparam|costesegso|0.0000||"
         Text            =   "cost.s"
         Top             =   1620
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   -73170
         MaxLength       =   8
         TabIndex        =   54
         Tag             =   "Coste Horas|N|N|||rparam|costehora|0.0000||"
         Text            =   "cost.h"
         Top             =   1290
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   -70110
         MaxLength       =   6
         TabIndex        =   34
         Tag             =   "Cajas por Palet|N|N|||rparam|cajasporpalet|###,##0||"
         Text            =   "ncajas"
         Top             =   4020
         Width           =   1170
      End
      Begin VB.CheckBox chkTraza 
         Caption         =   "Hay Trazabilidad"
         Height          =   375
         Left            =   -74520
         TabIndex        =   33
         Tag             =   "Hay Trazabilidad|N|S|||rparam|haytraza|0||"
         Top             =   5010
         Width           =   2145
      End
      Begin VB.CheckBox chkTaraTractor 
         Caption         =   "Se tara tractor de entrada"
         Height          =   375
         Left            =   -74520
         TabIndex        =   30
         Tag             =   "Se Tara Tractor|N|S|||rparam|setaratractor|0||"
         Top             =   3960
         Width           =   2775
      End
      Begin VB.Frame Frame3 
         Height          =   2775
         Left            =   -74865
         TabIndex        =   111
         Top             =   750
         Width           =   8340
         Begin VB.CheckBox ChkCajas 
            Height          =   225
            Index           =   4
            Left            =   7530
            TabIndex        =   28
            Tag             =   "Son Cajas 5|N|S|||rparam|escaja5|||"
            Top             =   2340
            Width           =   285
         End
         Begin VB.CheckBox ChkCajas 
            Height          =   225
            Index           =   3
            Left            =   7530
            TabIndex        =   25
            Tag             =   "Son Cajas 4|N|S|||rparam|escaja4|||"
            Top             =   1950
            Width           =   285
         End
         Begin VB.CheckBox ChkCajas 
            Height          =   225
            Index           =   2
            Left            =   7530
            TabIndex        =   22
            Tag             =   "Son Cajas 3|N|S|||rparam|escaja3|||"
            Top             =   1530
            Width           =   285
         End
         Begin VB.CheckBox ChkCajas 
            Height          =   225
            Index           =   1
            Left            =   7530
            TabIndex        =   19
            Tag             =   "Son Cajas 2|N|S|||rparam|escaja2|||"
            Top             =   1110
            Width           =   285
         End
         Begin VB.CheckBox ChkCajas 
            Height          =   225
            Index           =   0
            Left            =   7530
            TabIndex        =   16
            Tag             =   "Son Cajas 1|N|S|||rparam|escaja1|||"
            Top             =   690
            Width           =   285
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   6030
            MaxLength       =   6
            TabIndex        =   27
            Tag             =   "Peso Caja 5|N|S|||rparam|pesocaja5|##0.00||"
            Text            =   "peso 5"
            Top             =   2295
            Width           =   1170
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   6030
            MaxLength       =   6
            TabIndex        =   24
            Tag             =   "Peso Caja 4|N|S|||rparam|pesocaja4|##0.00||"
            Text            =   "peso 4"
            Top             =   1890
            Width           =   1170
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   6030
            MaxLength       =   6
            TabIndex        =   21
            Tag             =   "Peso Caja 3|N|S|||rparam|pesocaja3|##0.00||"
            Text            =   "peso 3"
            Top             =   1485
            Width           =   1170
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   6030
            MaxLength       =   6
            TabIndex        =   18
            Tag             =   "Peso Caja 2|N|S|||rparam|pesocaja2|##0.00||"
            Text            =   "peso 2"
            Top             =   1080
            Width           =   1170
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   6030
            MaxLength       =   6
            TabIndex        =   15
            Tag             =   "Peso Caja 1|N|S|||rparam|pesocaja1|##0.00||"
            Text            =   "peso 1"
            Top             =   675
            Width           =   1170
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   18
            Left            =   720
            MaxLength       =   20
            TabIndex        =   26
            Tag             =   "Tipo Caja 5|T|S|||rparam|tipocaja5|||"
            Text            =   "tipo 5"
            Top             =   2295
            Width           =   5175
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   17
            Left            =   720
            MaxLength       =   20
            TabIndex        =   23
            Tag             =   "Tipo Caja 4|T|S|||rparam|tipocaja4|||"
            Text            =   "tipo 4"
            Top             =   1890
            Width           =   5175
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   16
            Left            =   720
            MaxLength       =   20
            TabIndex        =   20
            Tag             =   "Tipo Caja 3|T|S|||rparam|tipocaja3|||"
            Text            =   "tipo 3"
            Top             =   1485
            Width           =   5175
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   15
            Left            =   720
            MaxLength       =   20
            TabIndex        =   17
            Tag             =   "Tipo Caja 2|T|S|||rparam|tipocaja2|||"
            Text            =   "tipo 2"
            Top             =   1080
            Width           =   5175
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   720
            MaxLength       =   20
            TabIndex        =   14
            Tag             =   "Tipo Caja 1|T|S|||rparam|tipocaja1|||"
            Text            =   "tipo 1"
            Top             =   675
            Width           =   5175
         End
         Begin VB.Label Label18 
            Caption         =   "Son Cajas"
            Height          =   285
            Left            =   7320
            TabIndex        =   144
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Caja "
            Height          =   285
            Left            =   720
            TabIndex        =   118
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "5.-"
            Height          =   285
            Left            =   405
            TabIndex        =   117
            Top             =   2295
            Width           =   240
         End
         Begin VB.Label Label9 
            Caption         =   "4.-"
            Height          =   285
            Left            =   405
            TabIndex        =   116
            Top             =   1890
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "3.-"
            Height          =   285
            Left            =   405
            TabIndex        =   115
            Top             =   1485
            Width           =   285
         End
         Begin VB.Label Label6 
            Caption         =   "2.-"
            Height          =   285
            Left            =   405
            TabIndex        =   114
            Top             =   1080
            Width           =   285
         End
         Begin VB.Label Label5 
            Caption         =   "1.-"
            Height          =   285
            Left            =   405
            TabIndex        =   113
            Top             =   675
            Width           =   285
         End
         Begin VB.Label Label4 
            Caption         =   "Peso de Caja"
            Height          =   285
            Left            =   6030
            TabIndex        =   112
            Top             =   315
            Width           =   1590
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   -71190
         TabIndex        =   110
         Top             =   1680
         Width           =   4470
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   -72480
         MaxLength       =   10
         TabIndex        =   40
         Tag             =   "Extension|N|N|||rparam|codextension|000||"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Frame Frame9 
         Caption         =   "Liquidaciones"
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
         Height          =   1050
         Left            =   -74460
         TabIndex        =   104
         Top             =   3180
         Width           =   7710
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   8
            Left            =   5805
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Tag             =   "C4 Liquidación|N|N|||rparam|c4liquaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   7
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Tag             =   "C3 Liquidación|N|N|||rparam|c3liquaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   6
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Tag             =   "C2 Liquidación|N|N|||rparam|c2liquaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   5
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Tag             =   "C1 Liquidacion|N|N|||rparam|c1liquaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 1"
            Height          =   195
            Index           =   26
            Left            =   90
            TabIndex        =   108
            Top             =   315
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 2"
            Height          =   195
            Index           =   25
            Left            =   1980
            TabIndex        =   107
            Top             =   315
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 3"
            Height          =   195
            Index           =   16
            Left            =   3915
            TabIndex        =   106
            Top             =   315
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 4"
            Height          =   195
            Index           =   14
            Left            =   5805
            TabIndex        =   105
            Top             =   315
            Width           =   1305
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Anticipos"
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
         Height          =   1050
         Left            =   -74460
         TabIndex        =   99
         Top             =   2040
         Width           =   7710
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   4
            Left            =   5805
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Tag             =   "C4 Anticipo|N|N|||rparam|c4antiaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Tag             =   "C3 Anticipo|N|N|||rparam|c3antiaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Tag             =   "C1 Anticipo|N|N|||rparam|c1antiaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Tag             =   "C2 Anticipo|N|N|||rparam|c2antiaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 4"
            Height          =   195
            Index           =   12
            Left            =   5805
            TabIndex        =   103
            Top             =   315
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 3"
            Height          =   195
            Index           =   10
            Left            =   3915
            TabIndex        =   102
            Top             =   315
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 2"
            Height          =   195
            Index           =   8
            Left            =   1980
            TabIndex        =   101
            Top             =   315
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 1"
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   100
            Top             =   315
            Width           =   1665
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   -72465
         MaxLength       =   10
         TabIndex        =   39
         Tag             =   "Carpeta Facturas|N|N|||rparam|codcarpetaliqu|000||"
         Top             =   1035
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   -71190
         TabIndex        =   97
         Top             =   1035
         Width           =   4470
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   -72465
         MaxLength       =   10
         TabIndex        =   38
         Tag             =   "Carpeta Albaranes|N|N|||rparam|codcarpetaanti|000||"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   -71190
         TabIndex        =   95
         Top             =   720
         Width           =   4470
      End
      Begin VB.Frame Frame7 
         Height          =   1815
         Left            =   -74595
         TabIndex        =   89
         Top             =   900
         Width           =   8010
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   9
            Tag             =   "Direccion e-mail|T|S|||rparam|diremail|||"
            Text            =   "3"
            Top             =   450
            Width           =   6210
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   6
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   10
            Tag             =   "Servidor SMTP|T|S|||rparam|smtpHost|||"
            Text            =   "3"
            Top             =   900
            Width           =   6210
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   7
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   11
            Tag             =   "Usuario SMTP|T|S|||rparam|smtpUser|||"
            Text            =   "3"
            Top             =   1440
            Width           =   3090
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   8
            Left            =   5250
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   12
            Tag             =   "Password SMTP|T|S|||rparam|smtpPass|||"
            Text            =   "3"
            Top             =   1440
            Width           =   2220
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   94
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   93
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   92
            Top             =   1500
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   23
            Left            =   4440
            TabIndex        =   91
            Top             =   1500
            Width           =   840
         End
         Begin VB.Label Label8 
            Caption         =   "Envio E-Mail"
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
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   90
            Top             =   0
            Width           =   1320
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Soporte"
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
         Height          =   1035
         Left            =   -74595
         TabIndex        =   87
         Top             =   2940
         Width           =   8025
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   1350
            MaxLength       =   100
            TabIndex        =   13
            Tag             =   "Web Soporte|T|S|||rparam|websoporte|||"
            Top             =   360
            Width           =   6135
         End
         Begin VB.Label Label2 
            Caption         =   "Web soporte"
            Height          =   255
            Left            =   180
            TabIndex        =   88
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   -74520
         TabIndex        =   81
         Top             =   720
         Width           =   7665
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   2235
            MaxLength       =   20
            TabIndex        =   0
            Tag             =   "Servidor Contabilidad|T|S|||rparam|serconta|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   210
            Width           =   4875
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   4230
            MaxLength       =   15
            TabIndex        =   86
            Tag             =   "Código Parámetros Aplic|N|N|||sparam|codparam||S|"
            Text            =   "1"
            Top             =   240
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   2235
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   2
            Tag             =   "Password Contabilidad|T|S|||rparam|pasconta|||"
            Text            =   "3"
            Top             =   840
            Width           =   4875
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   2235
            MaxLength       =   20
            TabIndex        =   1
            Tag             =   "Usuario Contabilidad|T|S|||rparam|usuconta|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   525
            Width           =   4875
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2235
            MaxLength       =   2
            TabIndex        =   3
            Tag             =   "Nº Contabilidad|N|S|||rparam|numconta|||"
            Text            =   "3"
            Top             =   1185
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   15
            Left            =   300
            TabIndex        =   85
            Top             =   870
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   17
            Left            =   300
            TabIndex        =   84
            Top             =   570
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
            Height          =   195
            Index           =   18
            Left            =   300
            TabIndex        =   83
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
            Height          =   195
            Index           =   19
            Left            =   300
            TabIndex        =   82
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   59
         Left            =   2205
         ToolTipText     =   "Buscar cuenta"
         Top             =   2370
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Banco Prevista"
         Height          =   195
         Index           =   45
         Left            =   495
         TabIndex        =   206
         Top             =   2400
         Width           =   1650
      End
      Begin VB.Label Label22 
         Caption         =   "Peso Caja Llena"
         Height          =   285
         Left            =   -68820
         TabIndex        =   204
         Top             =   4050
         Width           =   1320
      End
      Begin VB.Label Label21 
         Caption         =   "Mínimo"
         Height          =   285
         Left            =   -70110
         TabIndex        =   203
         Top             =   4590
         Width           =   570
      End
      Begin VB.Label Label20 
         Caption         =   "Máximo"
         Height          =   285
         Left            =   -68850
         TabIndex        =   202
         Top             =   4590
         Width           =   570
      End
      Begin VB.Label Label19 
         Caption         =   "Límites Kilos Caja"
         Height          =   285
         Left            =   -71610
         TabIndex        =   201
         Top             =   4830
         Width           =   1950
      End
      Begin VB.Label Label1 
         Caption         =   "Sección Bodega"
         Height          =   195
         Index           =   53
         Left            =   450
         TabIndex        =   200
         Top             =   1050
         Width           =   1650
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   2160
         ToolTipText     =   "Buscar Sección"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Base Datos Ariges"
         Height          =   195
         Index           =   52
         Left            =   -74550
         TabIndex        =   195
         Top             =   1470
         Width           =   1650
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   60
         Left            =   -72840
         ToolTipText     =   "Buscar Sección"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sección Suministros"
         Height          =   195
         Index           =   46
         Left            =   -74550
         TabIndex        =   194
         Top             =   1050
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Carpeta ADV"
         Height          =   195
         Index           =   47
         Left            =   -74310
         TabIndex        =   185
         Top             =   1395
         Width           =   1200
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   -72780
         ToolTipText     =   "Buscar Carpeta"
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   58
         Left            =   -72780
         ToolTipText     =   "Buscar cuenta"
         Top             =   2370
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Banco Prevista"
         Height          =   195
         Index           =   44
         Left            =   -74490
         TabIndex        =   182
         Top             =   2370
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Almacén"
         Height          =   195
         Index           =   42
         Left            =   -74460
         TabIndex        =   180
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   -73110
         ToolTipText     =   "Buscar Almacén"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   -73110
         ToolTipText     =   "Buscar Sección"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sección ADV"
         Height          =   195
         Index           =   36
         Left            =   -74460
         TabIndex        =   178
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   -72750
         ToolTipText     =   "Buscar Iva"
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.IVA Extranjero"
         Height          =   195
         Index           =   5
         Left            =   -74520
         TabIndex        =   176
         Top             =   930
         Width           =   1650
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   -72750
         ToolTipText     =   "Buscar cuenta"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Retención"
         Height          =   195
         Index           =   13
         Left            =   -74520
         TabIndex        =   175
         Top             =   1290
         Width           =   1650
      End
      Begin VB.Label Label17 
         Caption         =   "Path Ficheros clasificación"
         Height          =   285
         Left            =   -74550
         TabIndex        =   172
         Top             =   5100
         Width           =   1890
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   -72750
         ToolTipText     =   "Buscar Sección"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sección Almazara"
         Height          =   195
         Index           =   34
         Left            =   -74490
         TabIndex        =   171
         Top             =   1050
         Width           =   1650
      End
      Begin VB.Label Label16 
         Caption         =   "Texto Pie de Toma de Datos"
         Height          =   225
         Left            =   -74580
         TabIndex        =   143
         Top             =   4200
         Width           =   2565
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   1
         Left            =   -71940
         ToolTipText     =   "Zoom descripción"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   -71970
         ToolTipText     =   "Zoom descripción"
         Top             =   2700
         Width           =   240
      End
      Begin VB.Label Label15 
         Caption         =   "Texto Cabecera de Toma de Datos"
         Height          =   225
         Left            =   -74580
         TabIndex        =   142
         Top             =   2730
         Width           =   2565
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje AFO"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   -74550
         TabIndex        =   141
         Top             =   2400
         Width           =   2040
      End
      Begin VB.Label Label14 
         Caption         =   "Impresora de Entradas"
         Height          =   285
         Left            =   -74520
         TabIndex        =   133
         Top             =   3630
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje Retención"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   58
         Left            =   -74550
         TabIndex        =   132
         Top             =   2040
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Sección Hortofrutícola"
         Height          =   195
         Index           =   0
         Left            =   -74550
         TabIndex        =   131
         Top             =   840
         Width           =   1650
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   -72780
         ToolTipText     =   "Buscar Sección"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Coste Seg.Social"
         Height          =   285
         Left            =   -74550
         TabIndex        =   129
         Top             =   1650
         Width           =   1590
      End
      Begin VB.Label Label12 
         Caption         =   "Coste Horas"
         Height          =   285
         Left            =   -74550
         TabIndex        =   128
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Cajas por Palet"
         Height          =   285
         Left            =   -71610
         TabIndex        =   119
         Top             =   4050
         Width           =   1590
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -72780
         ToolTipText     =   "Buscar Extensión"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Extensión"
         Height          =   195
         Index           =   28
         Left            =   -74310
         TabIndex        =   109
         Top             =   1725
         Width           =   1380
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   -72780
         ToolTipText     =   "Buscar Carpeta"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Carpeta Liquidacion"
         Height          =   195
         Index           =   7
         Left            =   -74310
         TabIndex        =   98
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   -72780
         ToolTipText     =   "Buscar Carpeta"
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Carpeta Anticipos"
         Height          =   195
         Index           =   6
         Left            =   -74310
         TabIndex        =   96
         Top             =   810
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8145
      TabIndex        =   75
      Top             =   6330
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   240
      TabIndex        =   78
      Top             =   6225
      Width           =   3000
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
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   210
         Width           =   2760
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6870
      TabIndex        =   74
      Top             =   6330
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   76
      Top             =   6330
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Añadir"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
Private WithEvents frmAlm As frmComercial
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Dim indice As Byte
Dim Encontrado As Boolean
Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: Añadir
'4: Modificar

Dim vSeccion As CSeccion

Private Sub chkAgruparNotas_Click()
    If (chkAgruparNotas.Value = 1) Then
        Me.chkRespetarNroNota.Enabled = False
        Me.chkRespetarNroNota.Value = 0
    Else
        Me.chkRespetarNroNota.Enabled = True
    End If
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
                
                ' entradas de almacen
                vParamAplic.SeTaraTractor = Me.chkTaraTractor.Value
                vParamAplic.HayTraza = Me.chkTraza.Value
                vParamAplic.CajasporPalet = ComprobarCero(Text1(24).Text)
                vParamAplic.SeAgrupanNotas = Me.chkAgruparNotas.Value
                vParamAplic.SeRespetaNota = Me.chkRespetarNroNota.Value
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
                vParamAplic.EsCaja1 = Me.ChkCajas(0).Value
                vParamAplic.EsCaja2 = Me.ChkCajas(1).Value
                vParamAplic.EsCaja3 = Me.ChkCajas(2).Value
                vParamAplic.EsCaja4 = Me.ChkCajas(3).Value
                vParamAplic.EsCaja5 = Me.ChkCajas(4).Value
                vParamAplic.KilosCajaMin = ComprobarCero(Text1(64).Text)
                vParamAplic.KilosCajaMax = ComprobarCero(Text1(65).Text)
                vParamAplic.PesoCajaLLena = ComprobarCero(Text1(66).Text)
                
                vParamAplic.ImpresoraEntradas = Replace(Text1(31).Text, "\", "\\")
                'aridoc
                vParamAplic.CarpetaAnt = Text1(10)
                vParamAplic.CarpetaLiq = Text1(11)
                vParamAplic.CarpetaADV = Text1(61)
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
                
                vParamAplic.CosteHora = ComprobarCero(Text1(25).Text)
                vParamAplic.CosteSegSo = ComprobarCero(Text1(26).Text)
                vParamAplic.SeccionHorto = ComprobarCero(Text1(27).Text)
                vParamAplic.SeccionAlmaz = ComprobarCero(Text1(48).Text)
                vParamAplic.SeccionADV = ComprobarCero(Text1(56).Text)
                vParamAplic.PorcreteFacSoc = Text1(28).Text
                
                vParamAplic.PrimFactAnt = Text1(29).Text
                vParamAplic.UltFactAnt = Text1(30).Text
                vParamAplic.PrimFactAntVC = Text1(32).Text
                vParamAplic.UltFactAntVC = Text1(33).Text
                vParamAplic.PrimFactLiq = Text1(34).Text
                vParamAplic.UltFactLiq = Text1(35).Text
                vParamAplic.PrimFactLiqVC = Text1(36).Text
                vParamAplic.UltFactLiqVC = Text1(12).Text
                
                vParamAplic.PorcenAFO = Text1(37).Text
                vParamAplic.TTomaDatos = Text1(38).Text
                vParamAplic.PieTomaDatos = Text1(39).Text
                vParamAplic.CodIvaIntra = Text1(40).Text
                vParamAplic.CtaTerReten = Text1(42).Text
                
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
                
                ' ADV
                vParamAplic.AlmacenADV = Text1(57).Text
                vParamAplic.CtaBancoADV = Text1(58).Text
                
                ' Suministros
                vParamAplic.SeccionSumi = ComprobarCero(Text1(60).Text)
                vParamAplic.BDAriges = Text1(62).Text
                
                ' Bodega
                vParamAplic.SeccionBodega = ComprobarCero(Text1(63).Text)
                vParamAplic.AlbRetiradaManual = Me.ChkContadorManual.Value
                vParamAplic.CtaBancoBOD = Text1(59).Text
                
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
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
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
Dim I As Byte
'    'Icono del formulario
'    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(5).Image = 11  'Salir
    End With
    
    'IMAGES para zoom
    For I = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next I
 
    LimpiarCampos   'Limpia los campos TextBox
   
   'cargar IMAGES de busqueda
    For I = 0 To 5
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 6 To 10
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 43 To 47
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 49 To 50
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    For I = 51 To 54
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 58 To 60
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    

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
        Me.SSTab1.TabsPerRow = 6
        AbrirConexionAridoc "root", "aritel"
    Else
        Me.SSTab1.TabsPerRow = 5
    End If
    
    PonerModo 0

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

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmDoc_DatoSeleccionado(CadenaSeleccion As String)
'Carpetas de Aridoc
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre carpeta
End Sub

Private Sub frmExt_DatoSeleccionado(CadenaSeleccion As String)
'Extension de Aridoc
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmAri_DatoSeleccionado(CadenaSeleccion As String)
Dim cad As String
    cad = RecuperaValor(CadenaSeleccion, 1)
    Text1(indice).Text = Mid(cad, 2, Len(cad))
    Text1(indice).Text = Format(Text1(indice).Text, "000")
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(indice).Text = Format(Text1(indice).Text, "000")
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de iva de la contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 3) 'Porceiva
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim NumNivel As Byte

TerminaBloquear
    
    If vParamAplic.NumeroConta = 0 Then Exit Sub
    
    Select Case Index
        Case 0, 3, 4, 60, 10
            Select Case Index
                Case 0 ' Seccion hortofrutícola
                    indice = Index + 27
                Case 3 ' seccion de Almazara
                    indice = 48
                Case 4 ' seccion de Adv
                    indice = 56
                Case 60 ' seccion de suministros
                    indice = Index
                Case 10 ' seccion de bodega
                    indice = 63
            End Select
            
            Set frmSec = New frmManSeccion
            frmSec.DatosADevolverBusqueda = "0|1|2|"
            frmSec.CodigoActual = Text1(indice).Text
            frmSec.Show vbModal
            Set frmSec = Nothing
            PonerFoco Text1(indice)
        
        Case 1  'Porcentaje iva de factura de terceros de extranjero
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    indice = Index + 39
                    Set frmIva = New frmTipIVAConta
                    frmIva.DatosADevolverBusqueda = "0|1|2|"
                    frmIva.CodigoActual = Text1(indice).Text
                    frmIva.Show vbModal
                    Set frmIva = Nothing
                    PonerFoco Text1(indice)
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
                    indice = Index + 40
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            

         Case 6, 7, 8 'carpetas de aridoc
            If Index = 8 Then
                indice = 61
            Else
                indice = Index + 4
            End If
            
            Set frmAri = New frmCarpAridoc
            frmAri.Opcion = 20
            frmAri.Show vbModal
            Set frmAri = Nothing
            PonerFoco Text1(indice)
        
         Case 9 'extesion de fichero de aridoc
            indice = Index + 4
            Set frmExt = New frmExtAridoc
            frmExt.DatosADevolverBusqueda = "0|1|"
            frmExt.CodigoActual = Text1(indice).Text
            frmExt.Show vbModal
            Set frmExt = Nothing
            PonerFoco Text1(indice)
                
        Case 43, 44 ' forma de pago de facturas de anticipos / liquidaciones socios
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    AbrirFrmForpaConta (Index)
                End If
            End If
        
        Case 45, 46, 47 ' cuenta de retencion y de aportacion de facturas anti / liqui de socios
                        ' 47 cta de banco prevista
            If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(27).Text) Then
                If vSeccion.AbrirConta Then
                    indice = Index
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(indice)
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
                    indice = Index
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(indice)
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
            
        Case 5 ' alamacen de adv
            Set frmAlm = New frmComercial
            
            AyudaAlmacenCom frmAlm, Text1(57).Text
            
            Set frmAlm = Nothing

        '59,58 cta de retencion de almazara y cta banco adv
        Case 59, 58
            If Text1(48).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(56).Text) Then
                If vSeccion.AbrirConta Then
                    indice = Index
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = Text1(indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco Text1(indice)
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing


    End Select

    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.Data1, 1

End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            indice = 38
            frmZ.pTitulo = "Texto para Cabecera de Toma de Datos"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
        
            frmZ.Show vbModal
            Set frmZ = Nothing
                
            PonerFoco Text1(indice)
        Case 1
            indice = 39
            frmZ.pTitulo = "Texto para Pie de Toma de Datos"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
        
            frmZ.Show vbModal
            Set frmZ = Nothing
                
            PonerFoco Text1(indice)
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
            Case 59: KEYBusqueda KeyAscii, 59 'cuenta de retencion adv
            Case 58: KEYBusqueda KeyAscii, 58 'cuenta de banco prevista adv
        
            Case 60: KEYBusqueda KeyAscii, 60 'seccion de suministros
            
            Case 63: KEYBusqueda KeyAscii, 10 'seccion de bodega
        
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim cad As String

    If Text1(Index).Text = "" Then Exit Sub

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
            
            
        Case 10, 11, 12
            If Text1(Index).Text = "" Then Exit Sub
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            cad = CargaPath(Text1(Index))
            Text2(Index).Text = Mid(cad, 2, Len(cad))
        
        
        Case 13
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "extension", "descripcion", "codext", "N", cAridoc)
        
        Case 14, 15, 16, 17, 18
            If Text1(Index).Text = "" Then Exit Sub
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        
        Case 19, 20, 21, 22, 23
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoDecimal Text1(Index), 4
       
        Case 64, 65, 66 ' limite inferior y superior de kilos caja
                        ' 66 peso caja llena
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoDecimal Text1(Index), 4
       
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
                        Text2(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(Index), "N")
                    Else
                        Text2(Index).Text = ""
                    End If
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing
            
            
        Case 45, 46, 47 ' cuentas contables de retencion aportacion y banco
                        ' para contabilizacion de facturas de socio
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
            
    ' ***********ALMAZARA*********
        Case 51, 52 ' forma de pago en positivo y en negativo
            If Text1(48).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(Text1(48).Text) Then
                If vSeccion.AbrirConta Then
                    If PonerFormatoEntero(Text1(Index)) Then
                        Text2(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(Index), "N")
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
    ' ***********END ALMAZARA*********
        
        
        
    ' ***********ADV*********
       Case 57 ' almacen de adv
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "salmpr", "nomalmac", "codalmac", "N", cAgro)
        
        Case 58, 59
            ' 58 cuenta contable de banco adv
            ' 59 cuenta contable retencion adv
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
    
    ' ***********END ADV*********
        
       Case 60 ' seccion de suministros
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rseccion", "nomsecci", "codsecci", "N", cAgro)
        
    
       Case 63 ' seccion de bodega
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rseccion", "nomsecci", "codsecci", "N", cAgro)
    
    
    End Select
End Sub


Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 6, 7
            If Text1(Index).Text <> "" Then
                If Not EsNumerico(Text1(Index).Text) Then
                    Cancel = True
                    ConseguirFoco Text1(Index), Modo
                End If
            End If
    End Select
    
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
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me)
    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
'    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
Dim I As Byte
Dim cad As String


On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    ' ************* si hay aridoc **************
    If vParamAplic.HayAridoc = 1 Then
         cad = CargaPath(Text1(10))
         Text2(10).Text = Mid(cad, 2, Len(cad))
         cad = CargaPath(Text1(11))
         Text2(11).Text = Mid(cad, 2, Len(cad))
         cad = CargaPath(Text1(61))
         Text2(61).Text = Mid(cad, 2, Len(cad))

         Text2(13).Text = DevuelveDesdeBDNew(cAridoc, "extension", "descripcion", "codext", Text1(13).Text, "N")
    End If
    
    ' ************* configurar els camps de les descripcions de les comptes *************
    If Text1(27).Text = "" Then Exit Sub ' si no hemos indicado la seccion
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
                Text2(43).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(43), "N")
            End If
            ' forma de pago de facturas anticipos / liquidaciones de socios negativas
            If Text1(44).Text <> "" Then
                Text2(44).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(44), "N")
            End If
            
            
        End If
        vSeccion.CerrarConta
    End If
    Set vSeccion = Nothing
    
    Text2(27).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(27).Text, "N")
    Text2(48).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(48).Text, "N")
    Text2(56).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(56).Text, "N")
    
    
    If Text1(48).Text = "" Then Exit Sub ' si no hemos indicado la seccion
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
                Text2(51).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(51), "N")
            End If
            ' forma de pago de facturas anticipos / liquidaciones de socios negativas
            If Text1(52).Text <> "" Then
                Text2(52).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", Text1(52), "N")
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
    
    ' almacen de adv
    Text2(57).Text = DevuelveDesdeBDNew(cAgro, "salmpr", "nomalmac", "codalmac", Text1(57).Text, "N")
    
    If Text1(56).Text = "" Then Exit Sub ' si no hemos indicado la seccion
                                         ' no sabemos a que contabilidad va
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(Text1(56).Text) Then
        If vSeccion.AbrirConta Then
            ' cuenta de retencion de adv
            If Text1(59).Text <> "" Then
                Text2(59).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(59), "T")
            End If
            ' cuenta de prevista de banco de facturas de adv
            If Text1(58).Text <> "" Then
                Text2(58).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Text1(58), "T")
            End If
        End If
        vSeccion.CerrarConta
    End If
    Set vSeccion = Nothing
    
    
    ' seccion de suministros
    Text2(60).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(60).Text, "N")
    
    ' seccion de bodega
    Text2(63).Text = DevuelveDesdeBDNew(cAgro, "rseccion", "nomsecci", "codsecci", Text1(63).Text, "N")
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
Dim I As Integer

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    For I = 1 To Combo1.Count - 1
        Combo1(I).ListIndex = -1
    Next I
    '### a mano
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim I As Byte
Dim vtag As CTag

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
      
    '------------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not b
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
'    BloquearCombo Me, Modo
    
    For I = 1 To 12
            Set vtag = New CTag
            vtag.Cargar Me.Combo1(I)
            If vtag.Cargado Then
                If vtag.EsClave And (Modo = 4 Or Modo = 5) Then
                    Me.Combo1(I).Enabled = False
                    Me.Combo1(I).BackColor = &H80000018 'groc
                Else
                    Me.Combo1(I).Enabled = b
                    If b Then
                        Me.Combo1(I).BackColor = vbWhite
                    Else
                        Me.Combo1(I).BackColor = &H80000018 'Amarillo Claro
                    End If
                    If Modo = 3 Then Me.Combo1(I).ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            End If
     Next I

    ' no se pueden modificar la primera y ultima factura de ultima facturaciones
    Frame5.Enabled = False
    
    'Bloquear imagen de Busqueda
    For I = 6 To 8
        Me.imgBuscar(I).Enabled = (Modo >= 3)
        Me.imgBuscar(I).visible = (Modo >= 3)
    Next I
    For I = 9 To 10
        Me.imgBuscar(I).Enabled = (Modo >= 3)
        Me.imgBuscar(I).visible = (Modo >= 3)
    Next I
    For I = 0 To 5
        Me.imgBuscar(I).Enabled = (Modo >= 3)
        Me.imgBuscar(I).visible = (Modo >= 3)
    Next I
    For I = 43 To 47
        Me.imgBuscar(I).Enabled = (Modo >= 3)
        Me.imgBuscar(I).visible = (Modo >= 3)
    Next I
    For I = 49 To 54
        Me.imgBuscar(I).Enabled = (Modo >= 3)
        Me.imgBuscar(I).visible = (Modo >= 3)
    Next I
    For I = 58 To 60
        Me.imgBuscar(I).Enabled = (Modo >= 3)
        Me.imgBuscar(I).visible = (Modo >= 3)
    Next I
    
'    BloquearImgBuscar Me, Modo
    'Bloquear los checkbox
    BloquearChecks Me, Modo
    
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not Encontrado And Not b  'Añadir
    Me.Toolbar1.Buttons(2).Enabled = Encontrado And Not b 'Modificar
    Me.mnAñadir.Enabled = Not Encontrado And Not b
    Me.mnModificar.Enabled = Encontrado And Not b
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
End Sub


' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 1 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    'combos de anticipos
    For I = 1 To 4
        Combo1(I).AddItem "Nro.Factura"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 0
        Combo1(I).AddItem "Cod.Socio"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 1
        Combo1(I).AddItem "Nom.Socio"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 2
        Combo1(I).AddItem "Variedad"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 3
    Next I
    
    'combos de liquidacion
    For I = 5 To 8
        Combo1(I).AddItem "Nro.Factura"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 0
        Combo1(I).AddItem "Cod.Socio"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 1
        Combo1(I).AddItem "Nom.Socio"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 2
        Combo1(I).AddItem "Variedad"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 3
    Next I
    
    'combos de adv
    For I = 9 To 12
        Combo1(I).AddItem "Nro.Factura"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 0
        Combo1(I).AddItem "Cod.Socio"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 1
        Combo1(I).AddItem "Nom.Socio"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 2
        Combo1(I).AddItem "Procedencia"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 3
    Next I
    
    
    
End Sub

Private Function CargaPath(Codigo As Integer) As String
Dim Nod As Node
Dim J As Integer
Dim I As Integer
Dim C As String
Dim campo1 As String
Dim padre As String
Dim A As String

    'Primero copiamos la carpeta
    C = "\" & DevuelveDesdeBDNew(cAridoc, "carpetas", "nombre", "codcarpeta", CInt(Codigo), "N")
    campo1 = "nombre"
    padre = DevuelveDesdeBDNew(cAridoc, "carpetas", "padre", "codcarpeta", CStr(Codigo), "N", campo1)
    If CInt(padre) > 0 Then
        C = CargaPath(CInt(padre)) & C
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
    CargaPath = C
End Function


Private Sub AbrirFrmForpaConta(indice1 As Integer)
    indice = indice1
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = Text1(indice)
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub

