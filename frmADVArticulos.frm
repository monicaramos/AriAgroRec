VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmADVArticulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos ADV"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "frmADVArticulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   34
      Top             =   480
      Width           =   11295
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   990
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "Código de articulo|T|N|||advartic|codartic||S|"
         Text            =   "1234567890123456"
         Top             =   240
         Width           =   1530
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||advartic|nomartic|||"
         Top             =   240
         Width           =   4140
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre "
         Height          =   255
         Left            =   2745
         TabIndex        =   36
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   6840
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
         TabIndex        =   23
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10410
      TabIndex        =   20
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Top             =   6960
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   240
      TabIndex        =   33
      Top             =   1380
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9631
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
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmADVArticulos.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(11)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSumaStocks"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgFec(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(16)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(9)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgBuscar(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "imgBuscar(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgBuscar(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(5)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(6)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(8)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(17)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "imgBuscar(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(19)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(20)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chkCtrStock(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtSumaStock"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(10)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(8)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(6)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text2(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text2(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text2(7)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(7)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(3)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text2(5)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "FrameDatosAlmacen"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(4)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(9)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(11)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(12)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(14)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(16)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "combo1(0)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(21)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(22)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "Stocks Almacenes"
      TabPicture(1)   =   "frmADVArticulos.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Materias Activas"
      TabPicture(2)   =   "frmADVArticulos.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux1"
      Tab(2).ControlCount=   1
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4485
         Left            =   -74850
         TabIndex        =   75
         Top             =   390
         Width           =   10245
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   12
            Left            =   3930
            MaxLength       =   10
            TabIndex        =   81
            Tag             =   "Cantidad|N|S|||advartic_matactiva|cantidad|#,##0.0000|N|"
            Text            =   "Text3"
            Top             =   3990
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.TextBox Text3 
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   10
            Left            =   180
            MaxLength       =   16
            TabIndex        =   80
            Tag             =   "Código Articulo|T|N|||advartic_matactiva|codartic||S|"
            Text            =   "Text3"
            Top             =   3990
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.TextBox Text3 
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   11
            Left            =   900
            MaxLength       =   8
            TabIndex        =   79
            Tag             =   "Codigo Materia Act.|N|N|||advartic_matactiva|codmatact|000000|S|"
            Text            =   "Text3"
            Top             =   3990
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.TextBox txtAux2 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   77
            Text            =   "Text2"
            Top             =   4005
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   2
            Left            =   1575
            MaskColor       =   &H00000000&
            TabIndex        =   76
            ToolTipText     =   "Buscar materia activa"
            Top             =   4005
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   90
            TabIndex        =   78
            Top             =   90
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
            Height          =   330
            Index           =   1
            Left            =   4260
            Top             =   120
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
            Bindings        =   "frmADVArticulos.frx":0060
            Height          =   3825
            Index           =   1
            Left            =   90
            TabIndex        =   82
            Top             =   570
            Width           =   10830
            _ExtentX        =   19103
            _ExtentY        =   6747
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
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   22
         Left            =   4620
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "Nro.Serie|T|S|||advartic|numserie||N|"
         Text            =   "T"
         Top             =   2745
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   21
         Left            =   1950
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "ADR|T|S|||advartic|adr||N|"
         Text            =   "T"
         Top             =   2745
         Width           =   1635
      End
      Begin VB.ComboBox combo1 
         Height          =   315
         Index           =   0
         Left            =   1950
         TabIndex        =   6
         Tag             =   "Cod. Tipo Artículo|N|N|0|2|advartic|tipoprod||N|"
         Text            =   "Combo2"
         Top             =   2310
         Width           =   1665
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   16
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   67
         Text            =   "Text1"
         Top             =   4410
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   14
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   4050
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   12
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   3690
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   11
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   3330
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   2970
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   8910
         MaxLength       =   6
         TabIndex        =   17
         Tag             =   "Unidades por Caja|N|N|||advartic|unicajas|####0|N|"
         Text            =   "Text1"
         Top             =   3570
         Width           =   765
      End
      Begin VB.Frame FrameDatosAlmacen 
         Caption         =   "Datos Relacionados con Almacen"
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
         Height          =   1860
         Left            =   6270
         TabIndex        =   48
         Top             =   1620
         Width           =   3630
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   18
            Left            =   2115
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "Fecha última compra|F|S|||advartic|ultfecco|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   1440
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   2115
            MaxLength       =   12
            TabIndex        =   15
            Tag             =   "Precio Venta al público|N|N|0|999999.0000|advartic|preciove|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   1080
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   2115
            MaxLength       =   12
            TabIndex        =   14
            Tag             =   "Precio Ultima Compra|N|S|0|999999.0000|advartic|preciouc|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   720
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   2100
            MaxLength       =   12
            TabIndex        =   13
            Tag             =   "Precio Medio Ponderado|N|S|0|999999.0000|advartic|preciomp|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   1320
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1845
            Picture         =   "frmADVArticulos.frx":0078
            ToolTipText     =   "Buscar fecha"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Últ. fecha compra"
            Height          =   255
            Index           =   15
            Left            =   270
            TabIndex        =   52
            Top             =   1485
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Precio Venta Público"
            Height          =   255
            Index           =   14
            Left            =   270
            TabIndex        =   51
            Top             =   1125
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Precio última compra"
            Height          =   255
            Index           =   12
            Left            =   255
            TabIndex        =   50
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Index           =   10
            Left            =   255
            TabIndex        =   49
            Top             =   420
            Width           =   1815
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   1575
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   1950
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "Cod. Tipo Unidad|N|N|0|99|advartic|codunida|00|N|"
         Text            =   "Text1"
         Top             =   1575
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Cod. Proveedor|N|N|0|999999|advartic|codprove|000000|N|"
         Text            =   "Text1"
         Top             =   855
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1950
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "Cod. Familia|N|N|0|9999|advartic|codfamia|0000|N|"
         Text            =   "Text1"
         Top             =   1215
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   5
         Tag             =   "Tipo de IVA|N|N|||advartic|codigiva|##0|N|"
         Text            =   "Ttt"
         Top             =   1935
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   7
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   46
         Text            =   "Text2"
         Top             =   1935
         Width           =   3285
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   45
         Text            =   "Text2"
         Top             =   1215
         Width           =   3285
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   855
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   10080
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "Te"
         Top             =   4770
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   8340
         MaxLength       =   13
         TabIndex        =   11
         Tag             =   "Código de Barras|T|S|||advartic|codigoea||N|"
         Text            =   "1234567890123"
         Top             =   855
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   8340
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Fecha de Alta|F|N|||advartic|fecaltas|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   1260
         Width           =   1335
      End
      Begin VB.TextBox txtSumaStock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   8070
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   4485
         Width           =   1590
      End
      Begin VB.CheckBox chkCtrStock 
         Caption         =   "¿Control de stock?"
         Height          =   315
         Index           =   0
         Left            =   8160
         TabIndex        =   18
         Tag             =   "Control de stock|N|N|0|1|advartic|ctrstock||N|"
         Top             =   4020
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Index           =   20
         Left            =   285
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Tag             =   "Texto para compras|T|S|||advartic|textocom|||"
         Top             =   4440
         Width           =   5865
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Index           =   19
         Left            =   285
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "Texto para Ventas|T|S|||advartic|textoven|||"
         Top             =   3420
         Width           =   5865
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   4440
         Left            =   -74760
         TabIndex        =   40
         Top             =   480
         Width           =   10920
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   9
            Left            =   9210
            MaxLength       =   8
            TabIndex        =   71
            Text            =   "Text3"
            Top             =   3930
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox Text3 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   90
            MaxLength       =   16
            TabIndex        =   70
            Tag             =   "Código Articulo|T|N|||advartic_salmac|codartic||S|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.TextBox Text3 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   810
            MaxLength       =   8
            TabIndex        =   24
            Tag             =   "Código Almacen|N|N|||advartic_salmac|codalmac|000|S|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.ComboBox cmbAux 
            Height          =   315
            Index           =   0
            Left            =   10125
            TabIndex        =   31
            Tag             =   "Status inventario|N|N|0|1|advartic_salmac|statusin|0|S|"
            Text            =   "Combo2"
            Top             =   3915
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   5
            Left            =   6615
            MaxLength       =   16
            TabIndex        =   28
            Tag             =   "Stock Máximo|N|S|||advartic_salmac|stockmax|#,###,###,##0.000|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   7
            Left            =   8370
            MaxLength       =   10
            TabIndex        =   30
            Tag             =   "Fecha inventario|F|S|||advartic_salmac|fechainv|dd/mm/yyyy|N|"
            Text            =   "Text3"
            Top             =   3930
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   6
            Left            =   7470
            MaxLength       =   16
            TabIndex        =   29
            Tag             =   "Stock inventario|N|S|||advartic_salmac|stockinv|#,###,###,##0.000|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   5760
            MaxLength       =   16
            TabIndex        =   27
            Tag             =   "Punto de Pedido|N|S|||advartic_salmac|puntoped|#,###,###,##0.000|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   4815
            MaxLength       =   16
            TabIndex        =   26
            Tag             =   "Stock Mínimo|N|S|||advartic_salmac|stockmin|#,###,###,##0.000|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   3735
            MaxLength       =   16
            TabIndex        =   25
            Tag             =   "Cantidad Stock|N|N|||advartic_salmac|canstock|#,###,###,##0.000|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.TextBox txtAux2 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   1710
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   69
            Text            =   "Text2"
            Top             =   3915
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   0
            Left            =   1485
            MaskColor       =   &H00000000&
            TabIndex        =   32
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   3915
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   1
            Left            =   9000
            MaskColor       =   &H00000000&
            TabIndex        =   68
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   3915
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   0
            TabIndex        =   41
            Top             =   0
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
            Top             =   480
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
            Bindings        =   "frmADVArticulos.frx":0103
            Height          =   3825
            Index           =   0
            Left            =   0
            TabIndex        =   42
            Top             =   480
            Width           =   10830
            _ExtentX        =   19103
            _ExtentY        =   6747
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
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   9210
            MaxLength       =   40
            TabIndex        =   72
            Tag             =   "Hora Inventario|FH|S|||advartic_salmac|horainve|yyyy-mm-dd hh:mm:ss|N|"
            Text            =   "Text3"
            Top             =   3660
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Unidades por Caja"
         Height          =   255
         Index           =   4
         Left            =   7350
         TabIndex        =   83
         Top             =   3630
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Serie"
         Height          =   255
         Index           =   3
         Left            =   3750
         TabIndex        =   74
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ADR"
         Height          =   255
         Index           =   1
         Left            =   270
         TabIndex        =   73
         Top             =   2730
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1650
         ToolTipText     =   "Buscar tipo unidad"
         Top             =   1575
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Tipo Unidad"
         Height          =   255
         Index           =   17
         Left            =   270
         TabIndex        =   62
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de I.V.A."
         Height          =   255
         Index           =   8
         Left            =   270
         TabIndex        =   61
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Familia"
         Height          =   255
         Index           =   6
         Left            =   270
         TabIndex        =   60
         Top             =   1215
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.  Proveedor"
         Height          =   255
         Index           =   5
         Left            =   270
         TabIndex        =   59
         Top             =   855
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1650
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1650
         ToolTipText     =   "Buscar familia"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1650
         ToolTipText     =   "Buscar tipo IVA"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Artículo"
         Height          =   255
         Index           =   9
         Left            =   270
         TabIndex        =   58
         Top             =   2325
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Asociado"
         Height          =   255
         Index           =   2
         Left            =   6510
         TabIndex        =   57
         Top             =   855
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Alta"
         Height          =   255
         Index           =   16
         Left            =   6510
         TabIndex        =   56
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   8070
         Picture         =   "frmADVArticulos.frx":011B
         ToolTipText     =   "Buscar fecha"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label lblSumaStocks 
         Caption         =   "Suma Stock Almacenes"
         Height          =   375
         Left            =   6990
         TabIndex        =   55
         Top             =   4455
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Texto para Compras"
         Height          =   240
         Index           =   2
         Left            =   285
         TabIndex        =   54
         Top             =   4245
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Texto para Ventas"
         Height          =   240
         Index           =   11
         Left            =   285
         TabIndex        =   53
         Top             =   3225
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4200
      Top             =   6960
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
      TabIndex        =   38
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
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
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
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
         Left            =   8520
         TabIndex        =   39
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10440
      TabIndex        =   37
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
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
Attribute VB_Name = "frmADVArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: ARTICULOS                 -+-+
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
Private WithEvents frmFam As frmComercial ' ayuda familias de comercial
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmPro As frmComercial ' ayuda proveedores de comercial
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmTUn As frmComercial ' ayuda tipos de unidad de comercial
Attribute frmTUn.VB_VarHelpID = -1
Private WithEvents frmMat As frmADVMatActivas ' Materias activas
Attribute frmMat.VB_VarHelpID = -1

'Private WithEvents frmA As frmManAlmProp  'almacenes


'Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Private WithEvents frmIva As frmTipIVAConta  'Tipos de IVA de la contabilidad
Attribute frmIva.VB_VarHelpID = -1
'' *****************************************************
'
'
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

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim PVPAnt As String

Dim vSeccion As CSeccion

Private Sub btnBuscar_Click(Index As Integer)
    Select Case Index
        Case 0 'Código de Almacen
'            Set frmA = New frmManAlmProp
'            frmA.DatosADevolverBusqueda = "0|1|"
'            frmA.Show vbModal
'            Set frmA = Nothing
            
        Case 2 ' codigo de materia activa
            Set frmMat = New frmADVMatActivas
            frmMat.DatosADevolverBusqueda = "0|1|"
            frmMat.Show vbModal
            Set frmMat = Nothing
        
            
            
        Case 1
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
            Set frmC = New frmCal
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
              While btnBuscar(Index).Parent.Name <> obj.Name
                    esq = esq + obj.Left
                    dalt = dalt + obj.Top
                    Set obj = obj.Container
              Wend
            
            menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
            frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            imgFec(Index).Tag = 7 '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If Text3(7) <> "" Then frmC.NovaData = Text3(7).Text
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco Text3(7) '<===
            ' ********************************************
                
    End Select

End Sub

Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String
Dim bol As Boolean
Dim Codigo As String

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                    Codigo = Text1(0).Text
                    InsetarArticulosPorAlmacen
                    TerminaBloquear
                    PosicionarData Codigo
                    CargaGrid 0, True
                    CargaGrid 1, True
                    
                    PonerModo 2
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    '[Monica]22/05/2012: si es moixent, preguntamos si queremos cambiar precios de albaranes y recalcular
                    '                    los albaranes que tengan este articulo
                    If vParamAplic.Cooperativa = 3 And Text1(17).Text <> PVPAnt Then
                        If MsgBox("¿ Quiere modificar los precios de este artículo en los albaranes ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            ModificarPreciosAlbaranes
                        End If
                    End If
                    TerminaBloquear
                    PosicionarData Text1(0).Text
                End If
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
'            If InsertarModificarLinea Then
''                DesBloqueaRegistroForm Text1(0)
'                TerminaBloquear
'                cad = "codalmac = " & Text3(0).Text & ""
'                If SituarData(Data4, cad, Indicador) Then
'                    ModificaLineas = 0
'                    lblIndicador.Caption = Indicador
'                    PonerModoFrame 0
'                    PonerSumaStocks
'                End If
'            End If
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    ModificarLinea
                    PosicionarData Data1.Recordset!codArtic
            End Select
            PonerSumaStocks
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
                Text1(0).BackColor = vbYellow 'codartic
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
Dim I As Integer

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
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    
    'cargar IMAGES de busqueda
    For I = 0 To imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "advartic"
    Ordenacion = " ORDER BY codartic"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codartic=-1"
    Data1.Refresh
    
    ModoLineas = 0
    CargaCombo 0
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'búsqueda
'        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
'        Text1(0).BackColor = vbYellow 'codartic
'        ' ****************************************************************************
'    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    chkCtrStock(0).Value = 0
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
Dim I As Integer, NumReg As Byte
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
    BloquearCombo Me, Modo
    
    
    BloquearTxt Text1(13), Modo >= 2
    

    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgFec Me, 0, Modo
    BloquearImgFec Me, 1, Modo
    BloquearChk chkCtrStock(0), (Modo = 0 Or Modo = 2 Or Modo = 5)
    
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
      
    ' ****** si n'hi han combos a la capçalera ***********************
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
Dim I As Byte
    
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
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = b
    Me.mnImprimir.Enabled = b
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Adoaux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
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
               
        Case 0 'stocks en almacenes
            Sql = "SELECT codartic,advartic_salmac.codalmac,salmpr.nomalmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin,CASE statusin WHEN 0 THEN ""No"" WHEN 1 THEN ""Sí"" END "
            Sql = Sql & " FROM advartic_salmac, salmpr "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE advartic_salmac.codartic = '-1'"
            End If
            Sql = Sql & " and advartic_salmac.codalmac = salmpr.codalmac "
            Sql = Sql & " ORDER BY advartic_salmac.codalmac"
        
        Case 1 'materias activas
            Sql = "SELECT advartic_matactiva.codartic,advartic_matactiva.codmatact,advmatactiva.nommatact ,advartic_matactiva.cantidad" ', advartic_matactiva.plazoent "
            Sql = Sql & " FROM advartic_matactiva, advmatactiva "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE advartic_matactiva.codartic = '-1'"
            End If
            Sql = Sql & " and advartic_matactiva.codmatact = advmatactiva.codmatact "
            Sql = Sql & " ORDER BY advartic_matactiva.codmatact "
            
    End Select
    
    MontaSQLCarga = Sql
End Function

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Colectivos
    Text3(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codalmacen
    FormateaCampo Text3(1)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomalmacen
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
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmMat_DatoSeleccionado(CadenaSeleccion As String)
'materias activas
    Text3(11).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    FormateaCampo Text3(11)
    Me.txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
'proveedores
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    FormateaCampo Text1(2)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de articxulos
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    FormateaCampo Text1(6)
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de IVA (de la Contabilidad)
    Text1(7).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo Text1(7)
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2) '% iva
End Sub

Private Sub frmTUn_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de unidad
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    FormateaCampo Text1(5)
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

' *** si n'hi ha buscar data, posar a les <=== el menor index de les imagens de buscar data ***
' NOTA: ha de coincidir l'index de la image en el del camp a on va a parar el valor
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
            imgFec(1).Tag = 10 '<===
        
            ' *** repasar si el camp es Text3 o Text1 ***
            If Text1(10).Text <> "" Then frmC.NovaData = Text1(10).Text
            ' ********************************************
        
        Case 1
            imgFec(1).Tag = 18 '<===
            
            ' *** repasar si el camp es Text3 o Text1 ***
            If Text1(18).Text <> "" Then frmC.NovaData = Text1(18).Text
            ' ********************************************
    End Select
    

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es Text3 o Text1 ***
    PonerFoco Text1(CByte(imgFec(1).Tag)) '<===
    ' ********************************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es Text3 o Text1 ***
    Select Case imgFec(1).Tag
        Case 10
            Text1(10).Text = Format(vFecha, "dd/mm/yyyy") '<===
        Case 18
            Text1(18).Text = Format(vFecha, "dd/mm/yyyy") '<===
        Case 7
            Text3(7).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End Select
    ' ********************************************
End Sub
' *****************************************************

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    Screen.MousePointer = vbHourglass
    frmADVListArticulos.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
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


Private Sub Text3_LostFocus(Index As Integer)
Dim cadMen As String

    
     If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 1 'Codigo Almacen
'             txtAux2(0).Text = PonerNombreDeCod(Text3(Index), "salmpr", "nomalmac")
'             If txtAux2(0).Text = "" Then
'                cadMen = "No existe el Almacén: " & Text3(Index).Text & vbCrLf
'                cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
'                If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                    Set frmA = New frmManAlmProp
'
'                    frmA.DatosADevolverBusqueda = "0|1|"
'                    frmA.NuevoCodigo = Text3(Index).Text
'                    Text3(Index).Text = ""
'                    TerminaBloquear
'                    frmA.Show vbModal
'                    Set frmA = Nothing
'                    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                Else
'                    Text3(Index).Text = ""
'                End If
'                PonerFoco Text3(Index)
'             End If
             
        Case 2, 3, 4, 5, 6 'Stocks, Punto Pedido
                'Formato tipo 1: Decimal(12,2)
                If Trim(Text3(Index)) <> "" Then PonerFormatoDecimal Text3(Index), 1
        
        Case 7  'Fecha Inventario
            If Text3(Index).Text <> "" Then PonerFormatoFecha Text3(Index)

        Case 9  'Hora Inventario
            If Trim(Text3(Index).Text) <> "" Then PonerFormatoHora Text3(Index)
            
        Case 11 ' materia activa
            If Text3(Index).Text = "" Then Exit Sub
            txtAux2(1).Text = PonerNombreDeCod(Text3(Index), "advmatactiva", "nommatact")
            If txtAux2(1).Text = "" Then
                cadMen = "No existe la Materia Activa: " & Text3(Index).Text & vbCrLf
                cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                    Set frmMat = New frmADVMatActivas

                    frmMat.DatosADevolverBusqueda = "0|1|"
                    frmMat.NuevoCodigo = Text3(Index).Text
                    Text3(Index).Text = ""
                    TerminaBloquear
                    frmMat.Show vbModal
                    Set frmMat = Nothing
                    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                Else
                    Text3(Index).Text = ""
                End If
                PonerFoco Text3(Index)
            End If
        
        Case 12 ' cantidad de materia activa
            If PonerFormatoDecimal(Text3(Index), 7) Then PonerFocoBtn cmdAceptar
            
    End Select
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
        Case 12 'Imprimir
            mnImprimir_Click
'            printNou
        Case 13    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer
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

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, , 1)
    
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
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(Text1(0), 25, "Código")
    cad = cad & ParaGrid(Text1(1), 50, "Nombre")
    cad = cad & ParaGrid(Text1(2), 25, "EAN")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Articulos" ' ***** repasa açò: títol de BuscaGrid *****
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
Dim cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            cad = cad & Text1(J).Text & "|"
        End If
    Loop Until I = 0
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
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("sartic", "codartic")
'    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions
    Text1(17).Text = 0
    Combo1(0).ListIndex = 0

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
End Sub

Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    
    '[Monica]22/05/2012: si me modifican el precio de venta, y es Moixent, preguntare si quieren modificar albaranes
    PVPAnt = Text1(17).Text
    
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
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
    cad = "¿Seguro que desea eliminar el Artículo?"
    cad = cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not eliminar Then
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
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Articulo", Err.Description
End Sub

Private Sub PonerCampos()
Dim I As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For I = 0 To 1
            CargaGrid I, True
            If Not Adoaux(I).Recordset.EOF Then _
                PonerCamposForma2 Me, Adoaux(I), 2, "FrameAux" & I
    Next I

    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    Text2(2).Text = PonerNombreDeCod(Text1(2), "proveedor", "nomprove")
    Text2(3).Text = PonerNombreDeCod(Text1(3), "advfamia", "nomfamia")
    Text2(5).Text = PonerNombreDeCod(Text1(5), "sunida", "nomunida")
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.SeccionADV) Then
        If vSeccion.AbrirConta Then
            ' porcentaje de iva de terceros
            If PonerFormatoEntero(Text1(7)) Then
                Text2(7).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(7), "N")
            Else
                Text2(7).Text = ""
            End If
        End If
        vSeccion.CerrarConta
    End If
    Set vSeccion = Nothing

    
    PonerSumaStocks 'Poner la suma total de stocks de los almacenes donde esta el artic

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
End Sub

Private Sub PonerSumaStocks()
Dim rst As ADODB.Recordset
Dim Sql As String
    
    Sql = DevuelveDesdeBDNew(cAgro, "advartic_salmac", "codartic", "codartic", Text1(0).Text, "T")
    If Sql <> "" Then
        Sql = "select sum(canstock) from advartic_salmac where codartic=" & DBSet(Text1(0).Text, "T")
        Set rst = New ADODB.Recordset
        rst.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not rst.EOF Then
            Me.txtSumaStock.Text = rst.Fields(0).Value
        End If
        rst.Close
        Set rst = Nothing
    Else
        Me.txtSumaStock.Text = 0
    End If
End Sub

Private Sub cmdCancelar_Click()
Dim I As Integer
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
                        LLamaLineas NumTabMto, ModoLineas 'ocultar Text3
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'Text32(2).text = ""

                    End If
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto + 1)
                    
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        Adoaux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar Text3
                    PonerModo 2 '4
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

            End Select
            
            PosicionarData Data1.Recordset!codArtic
             
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
Dim Mens As String


    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    If Not b Then Exit Function
    ' ************************************************************************************
    
    ' comprobamos que si hay contabilidad las cuentas contables existan
'    If Modo = 3 Or Modo = 4 Then
'        If vParamAplic.NumeroConta <> 0 Then
'                 Text2(4).Text = PonerNombreCuenta(Text1(4), Modo)
'                 Text2(17).Text = PonerNombreCuenta(Text1(17), Modo)
                 
                 
'            If text1(4).Text <> "" Then
'                text2(4).Text = PonerNombreCuenta(text1(4), Modo)
'                If text2(4).Text = "" Then b = False
'            Else
'                MsgBox "Debe poner una Cuenta Contable Socio existente. Reintroduzca.", vbExclamation
'                b = False
'            End If
'            If text1(17).Text <> "" Then
'                text2(17).Text = PonerNombreCuenta(text1(17), Modo)
'                If text2(17).Text = "" Then b = False
'            Else
'                MsgBox "Debe poner una Cuenta Contable Cliente existente. Reintroduzca.", vbExclamation
'                b = False
'            End If
'        End If
'    End If
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData(Codigo As String)
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codartic='" & Trim(Codigo) & "')" 'DBSet(Text1(0).Text, "T") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador, False) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codartic=" & DBSet(Data1.Recordset!codArtic, "T")
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM advartic_salmac " & vWhere
        
    ' ***** elimina les llínies de materias activas****
    conn.Execute "DELETE FROM advartic_matactiva " & vWhere
        
    'Eliminar la CAPÇALERA
    vWhere = " WHERE codartic=" & DBSet(Data1.Recordset!codArtic, "T")
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



    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************

    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Codigo Artículo
            'Comprobar si ya existe el cod de articulo en la tabla
            If Modo = 3 Then 'Insertar
                If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
            End If

        Case 2 'Codigo de Proveedor
            If Modo = 1 Then Exit Sub
        
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "proveedor", "nomprove")
                If Text2(Index).Text = "" Then
                    MsgBox "No existe el Proveedor de Comercial " & Text1(Index).Text & ". Reintroduzca.", vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 3 'Código de Familia
            If Modo = 1 Then Exit Sub
        
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "advfamia", "nomfamia")
                If Text2(Index).Text = "" Then
                    MsgBox "No existe la Familia de ADV " & Text1(Index).Text & ". Reintroduzca.", vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If

            
        Case 5 'Código Tipo Unidad
            If Modo = 1 Then Exit Sub
            
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "sunida", "nomunida")
                If Text2(Index).Text = "" Then
                    MsgBox "No existe el Tipo de Unidad de Comercial " & Text1(Index).Text & ". Reintroduzca.", vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 7 'Tipo de IVA
             If vParamAplic.SeccionADV = "" Then Exit Sub ' si no hemos indicado la seccion
                                                 ' no sabemos a que contabilidad va
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.SeccionADV) Then
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

        Case 10, 18 'Fecha alta, Fecha última compra
            If Modo = 1 Then Exit Sub
            PonerFormatoFecha Text1(Index)

        Case 13, 15, 17 'Precios
            'Formato tipo 2: Decimal(10,4)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 7
           
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'proveedor
                Case 3: KEYBusqueda KeyAscii, 1 'familia
                Case 5: KEYBusqueda KeyAscii, 3 'tipo de unidad
                Case 6: KEYBusqueda KeyAscii, 4 'tipo de articulo
                Case 7: KEYBusqueda KeyAscii, 2 'tipo de iva
                Case 10: KEYFecha KeyAscii, 0 'fecha de alta
                Case 18: KEYFecha KeyAscii, 1 'fecha de ultima compra
            End Select
        End If
    Else
        If (Index <> 19 Or (Index = 19 And Text1(19).Text = "")) And _
           (Index <> 20 Or (Index = 20 And Text1(20).Text = "")) Then KEYpress KeyAscii
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

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYBusquedaLin(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
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
Dim eliminar As Boolean

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
    eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'articulo en almacen
            Sql = "Seguro que desea eliminar de la BD el registro:"
            Sql = Sql & vbCrLf & "Cod. Artículo: " & Adoaux(Index).Recordset.Fields(0)
            Sql = Sql & vbCrLf & "Cod. Almacen: " & Adoaux(Index).Recordset.Fields(1)

            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM advartic_salmac"
                Sql = Sql & vWhere & " AND codalmac= " & Adoaux(Index).Recordset!codAlmac
            End If
    
        Case 1 'materia activa
            Sql = "Seguro que desea eliminar de la BD el registro:"
            Sql = Sql & vbCrLf & "Cod. Artículo: " & Adoaux(Index).Recordset.Fields(0)
            Sql = Sql & vbCrLf & "Cod. Materia Activa: " & Adoaux(Index).Recordset.Fields(1)

            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM advartic_matactiva"
                Sql = Sql & vWhere & " AND codmatact= " & Adoaux(Index).Recordset!codmatact
            End If
    
    
    End Select

    If eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        
        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
            PonerCampos
            
        End If
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto + 1)
    End If
    
    ModoLineas = 0
    PosicionarData Data1.Recordset!codArtic
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text3(1), False
    BloquearBtn btnBuscar(0), False

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vTabla = "advartic_salmac"
        Case 1: vTabla = "advartic_matactiva"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***1
'            If Index = 1 Then
'                NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
'            End If

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
                Case 0 'stocks en almacenes
                    Text3(0).Text = Text1(0).Text 'codartic
                    Text3(1).Text = ""
                    For I = 1 To 9 'Text3.Count - 1
                        Text3(I).Text = ""
                    Next I
                    txtAux2(0).Text = ""
                    PonerFoco Text3(1)
            End Select
        
        Case 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***1
'            If Index = 1 Then
'                NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
'            End If

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                Case 1 ' materia activa
                    Text3(10).Text = Text1(0).Text 'codartic
                    Text3(11).Text = ""
                    For I = 12 To 12 '13
                        Text3(I).Text = ""
                    Next I
                    BloquearTxt Text3(11), False
                    txtAux2(1).Text = ""
                    PonerFoco Text3(11)
            End Select
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
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
    BloquearTxt Text3(0), True
    BloquearTxt Text3(1), True
    BloquearTxt txtAux2(0), True
    BloquearBtn btnBuscar(0), True
    Select Case Index
        Case 0 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
    
        Case 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
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
        Case 0 'stocks en almacenes
        
            For J = 0 To 1
                Text3(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            txtAux2(0).Text = DataGridAux(Index).Columns(2).Text
            For J = 3 To 8
                Text3(J - 1).Text = DataGridAux(Index).Columns(J).Text
            Next J
            
            Text3(9).Text = Mid(Text3(8).Text, 12, 8)
            
            For I = 0 To 0
                BloquearTxt Text3(I), False
            Next I
            
            PosicionarCombo cmbAux(0), Adoaux(Index).Recordset!statusin
            
        Case 1
            For J = 10 To 11
                Text3(J).Text = DataGridAux(Index).Columns(J - 10).Text
            Next J
            txtAux2(1).Text = DataGridAux(Index).Columns(2).Text
            For J = 12 To 12 '13
                Text3(J).Text = DataGridAux(Index).Columns(J - 9).Text
            Next J
            
            For I = 11 To 11
                BloquearTxt Text3(I), True
            Next I
            
        
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'stocks en almacenes
            PonerFoco Text3(2)
        Case 1
            PonerFoco Text3(12)
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
        Case 0 'stocks
             For jj = 1 To 9
                Text3(jj).visible = b
                Text3(jj).Top = alto
             Next jj
             Text3(9).Left = Text3(8).Left
             txtAux2(0).visible = b
             txtAux2(0).Top = alto
             For jj = 0 To 1
                btnBuscar(jj).visible = b
                btnBuscar(jj).Top = alto
             Next jj
             Me.cmbAux(0).visible = b
             Me.cmbAux(0).Top = alto
            
            
        Case 1 ' materias activas
             For jj = 11 To 12 ' 13
                Text3(jj).visible = b
                Text3(jj).Top = alto
             Next jj
             txtAux2(1).visible = b
             txtAux2(1).Top = alto
             For jj = 2 To 2
                btnBuscar(jj).visible = b
                btnBuscar(jj).Top = alto
                btnBuscar(jj).Enabled = (xModo = 1)
             Next jj
            
    End Select
End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo(Index As Integer)
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    cmbAux(0).Clear

    cmbAux(0).AddItem "No"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 0
    cmbAux(0).AddItem "Sí"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 1

    Combo1(0).Clear

    Combo1(0).AddItem "Producto"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Trabajo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Varios"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2

End Sub


Private Sub Text3_GotFocus(Index As Integer)
   If Not Text3(Index).MultiLine Then ConseguirFocoLin Text3(Index)
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not Text3(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Text3(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim devuelve As String

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    
    Select Case nomframe
        Case "FrameAux0"
                
            'Campo de cantidad de Stock (Son decimales)
            If Trim(Text3(2).Text) = "" Or IsNull(Text3(2).Text) Then
                MsgBox "El campo Cantidad Stock no puede ser nulo", vbExclamation, "Artículos"
                b = False
            End If
            If Not b Then Exit Function
            
            ' ******************************************************************************
            'Comprobamos  si existe
            devuelve = DevuelveDesdeBDNew(cAgro, "advartic_salmac", "codartic", "codartic", Text1(0).Text, "T", , "codalmac", Text3(1).Text, "N")
            If ModoLineas = 1 And devuelve <> "" Then
                b = False
                devuelve = "Ya existe el Artículo en el Almacen: " & vbCrLf
                devuelve = devuelve & "Codigo: " & Text3(0).Text & vbCrLf
                devuelve = devuelve & "Descripción: " & txtAux2(0).Text
                MsgBox devuelve, vbExclamation, "Artículos"
            End If
    
        Case "FrameAux1"
            PonerFormatoDecimal Text3(12), 7
            
    End Select
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
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Codigo Proveedor
            Set frmPro = New frmComercial
            AyudaProveedoresCom frmPro, Text1(2).Text
            Set frmPro = Nothing
            PonerFoco Text1(Index + 2)
        
        Case 1  'Cod. Familia
            Set frmFam = New frmComercial
            AyudaFamiliasADV frmFam, Text1(3).Text
            Set frmFam = Nothing
            PonerFoco Text1(Index + 2)
        
        Case 3  'Cod. Tipo Unidad
            Set frmTUn = New frmComercial
            AyudaTUnidadesCom frmTUn, Text1(5).Text
            Set frmTUn = Nothing
            PonerFoco Text1(Index + 2)
        
        Case 2  'Tipos de IVA. Tabla de la BD Contabilidad
            If vParamAplic.SeccionADV = "" Then Exit Sub  ' si no hemos indicado la seccion
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.SeccionADV) Then
                If vSeccion.AbrirConta Then
                    indice = 7
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
            
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Colectivos
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codfamia
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomfamia
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'tarifas
                If DataGridAux(Index).Columns.Count > 2 Then
'                    Text3(11).Text = DataGridAux(Index).Columns("direccio").Text
'                    Text3(12).Text = DataGridAux(Index).Columns("observac").Text
                End If
                
            Case 1 'bonificaiones
                If DataGridAux(Index).Columns.Count > 2 Then
'                    Text3(21).Text = DataGridAux(Index).Columns(5).Text
'                    Text3(22).Text = DataGridAux(Index).Columns(6).Text
'                    Text3(23).Text = DataGridAux(Index).Columns(8).Text
'                    Text3(24).Text = DataGridAux(Index).Columns(15).Text
'                    Text32(22).Text = DataGridAux(Index).Columns(7).Text
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'cuentas bancarias
'                Text3(11).Text = ""
'                Text3(12).Text = ""
            Case 1 'departamentos
                For I = 21 To 24
'                   Text3(i).Text = ""
                Next I
'               Text32(22).Text = ""
            Case 2 'Tarjetas
'               Text3(50).Text = ""
'               Text3(51).Text = ""
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    SSTab1.Tab = numTab
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************

Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
Dim tip As Integer
Dim I As Byte

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
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'stocks en almacenes
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;" 'codartic
            tots = tots & "S|Text3(1)|T|Alm.|650|;" 'almacen
            tots = tots & "S|btnBuscar(0)|B|||;S|txtAux2(0)|T|Denominación|1700|;"
            tots = tots & "S|Text3(2)|T|Cant.Stock|1000|;"
            tots = tots & "S|Text3(3)|T|Stock Min.|1000|;"
            tots = tots & "S|Text3(4)|T|Punto Ped.|1000|;"
            tots = tots & "S|Text3(5)|T|Stock Max.|1000|;"
            tots = tots & "S|Text3(6)|T|Stock Inv.|1000|;"
            tots = tots & "S|Text3(7)|T|Fecha Inv.|1200|;"
            tots = tots & "S|btnBuscar(1)|B|||;"
            tots = tots & "S|Text3(8)|T|Hora Inv.|1000|;"
            tots = tots & "N||||0|;" 'inventariandose
            tots = tots & "S|cmbAux(0)|C|Inv.|600|;"
            
            Text3(8).Tag = "Hora Inventario|FH|S|||advartic_salmac|horainve|hh:mm:ss|N|"
            arregla tots, DataGridAux(Index), Me
            Text3(8).Tag = "Hora Inventario|FH|S|||advartic_salmac|horainve|yyyy-mm-dd hh:mm:ss|N|"
        
            DataGridAux(0).Columns(1).Alignment = dbgLeft
'            DataGridAux(0).Columns(6).Alignment = dbgRight
'            DataGridAux(0).Columns(7).Alignment = dbgRight
'            DataGridAux(0).Columns(8).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            
        Case 1 'materias activas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;" 'codartic
            tots = tots & "S|Text3(11)|T|Código|850|;" 'codigo de materia activa
            tots = tots & "S|btnBuscar(2)|B|||;S|txtAux2(1)|T|Denominación|3700|;S|Text3(12)|T|Cantidad|1250|;"
'            tots = tots & "S|Text3(13)|T|Plazo Seguridad|2000|;"
'            tots = tots & "S|Text3(12)|T|Plazo Reentrega|2000|;"
            
            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(Index).ScrollBars = dbgAutomatic
            
            DataGridAux(0).Columns(1).Alignment = dbgLeft
'            DataGridAux(0).Columns(6).Alignment = dbgRight
'            DataGridAux(0).Columns(7).Alignment = dbgRight
'            DataGridAux(0).Columns(8).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not Adoaux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
'        LimpiarCamposFrame Index
    End If
      
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
        Case 0: nomframe = "FrameAux0" 'stocks en almacenes
        Case 1: nomframe = "FrameAux1" 'materias activas
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        If Text3(9).Text = "" Then Text3(9).Text = "00:00:00"
        Text3(8).Text = Format(Text3(7).Text, FormatoFecha) & " " & Text3(9).Text

        If InsertarDesdeForm2(Me, 2, nomframe) Then
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
                Case 1
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
           
            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'stocks en almacenes
        Case 1: nomframe = "FrameAux1" 'materias activas de los articulos
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        If Text3(9).Text = "" Then Text3(9).Text = "00:00:00"
        Text3(8).Text = Format(Text3(7).Text, FormatoFecha) & " " & Text3(9).Text
        
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            
            V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
            LLamaLineas NumTabMto, 0
        End If
    End If
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codartic='" & Trim(Text1(0).Text) & "'"
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
'    Select Case Index
'        Case 0 'Cuentas Bancarias
'            Text3(11).Text = ""
'            Text3(12).Text = ""
'        Case 1 'Departamentos
'            Text3(21).Text = ""
'            Text3(22).Text = ""
'            Text32(22).Text = ""
'            Text3(23).Text = ""
'            Text3(24).Text = ""
'        Case 2 'Tarjetas
'            Text3(50).Text = ""
'            Text3(51).Text = ""
'        Case 4 'comisiones
'            Text32(2).Text = ""
'    End Select
'
    If Err.Number <> 0 Then Err.Clear
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

Private Sub InsetarArticulosPorAlmacen()
'Inserta en la tabla salmac una fila del artículo que se esta insertando
'para cada uno de los almacenes que existen en la tabla salmpr
Dim vCodArtic As String, vcodalmac As Integer
Dim rsAlmPr As ADODB.Recordset
Dim cad As String
    
    On Error GoTo EInsEnAlm

    vCodArtic = Text1(0).Text
    Set rsAlmPr = New ADODB.Recordset
    cad = "Select codalmac from salmpr order by codalmac;"
    rsAlmPr.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not rsAlmPr.EOF
        vcodalmac = rsAlmPr.Fields(0).Value
        cad = "INSERT INTO advartic_salmac (codartic,codalmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
        cad = cad & " VALUES (" & DBSet(vCodArtic, "T") & "," & vcodalmac & ",0,0,0,0,0,NULL,NULL,0)"
        conn.Execute cad
        rsAlmPr.MoveNext
    Wend
        
    rsAlmPr.Close
    Set rsAlmPr = Nothing
EInsEnAlm:
    If Err.Number <> 0 Then MuestraError Err.Number, "Insertando Artículo en Almacenes.", Err.Description
End Sub

Private Function ModificarPreciosAlbaranes() As Boolean
Dim Sql As String, Sql2 As String
Dim Rs As ADODB.Recordset
Dim Importe As Currency
    On Error GoTo eModificarPreciosAlbaranes

    ModificarPreciosAlbaranes = False


    Sql = "select * from advpartes_lineas where codartic = " & DBSet(Text1(0).Text, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Importe = Round2(DBLet(Rs!cantidad, "N") * TransformaPuntosComas(ImporteSinFormato(Text1(17).Text)), 2)
        
        Sql2 = "update advpartes_lineas set preciove = " & DBSet(ImporteSinFormato(Text1(17).Text), "N")
        Sql2 = Sql2 & ", importel = " & DBSet(Importe, "N")
        Sql2 = Sql2 & " where numparte = " & DBSet(Rs!Numparte, "N")
        Sql2 = Sql2 & " and numlinea = " & DBSet(Rs!numlinea, "N")
        
        conn.Execute Sql2
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    ModificarPreciosAlbaranes = True
    Exit Function

eModificarPreciosAlbaranes:
    MuestraError Err.Number, "Modificar Precios Albaranes", Err.Description
End Function
