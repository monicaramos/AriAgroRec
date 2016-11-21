VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmManHcoFruta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Fruta Clasificada"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   11970
   Icon            =   "frmManHcoFruta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2370
      Left            =   90
      TabIndex        =   26
      Top             =   540
      Width           =   11730
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   10740
         MaxLength       =   8
         TabIndex        =   18
         Tag             =   "Contrato|T|S|||rhisfruta|contrato|||"
         Text            =   "Text1"
         Top             =   1890
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   4980
         MaxLength       =   4
         TabIndex        =   97
         Top             =   1920
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   10110
         MaxLength       =   7
         TabIndex        =   13
         Tag             =   "Peso Trans|N|N|||rhisfruta|kilostra|###,##0||"
         Top             =   1035
         Width           =   1350
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   8700
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "Transportado por|N|N|0|1|rhisfruta|transportadopor||N|"
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   10110
         MaxLength       =   12
         TabIndex        =   9
         Tag             =   "Pr.Estimado|N|S|||rhisfruta|prestimado|###,##0.0000||"
         Text            =   "123"
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   47
         Top             =   1920
         Width           =   3210
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   46
         Top             =   1920
         Width           =   2685
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   6450
         MaxLength       =   6
         TabIndex        =   45
         Top             =   1920
         Width           =   795
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   44
         Top             =   1590
         Width           =   2685
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2250
         MaxLength       =   4
         TabIndex        =   43
         Top             =   1590
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod.Variedad|N|N|0|999999|rhisfruta|codvarie|000000||"
         Text            =   "Text1"
         Top             =   855
         Width           =   1050
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   225
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Albarán|N|S|||rhisfruta|numalbar|0000000|S|"
         Text            =   "Text1"
         Top             =   450
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   5925
         MaxLength       =   8
         TabIndex        =   14
         Tag             =   "Importe Transporte|N|S|||rhisfruta|imptrans|#,##0.00||"
         Text            =   "123"
         Top             =   1590
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   7320
         MaxLength       =   8
         TabIndex        =   15
         Tag             =   "Importe Acarreo|N|S|||rhisfruta|impacarr|#,##0.00||"
         Text            =   "123"
         Top             =   1590
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   8715
         MaxLength       =   8
         TabIndex        =   16
         Tag             =   "Importe Recolección|N|S|||rhisfruta|imprecol|#,##0.00||"
         Text            =   "123"
         Top             =   1590
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   10110
         MaxLength       =   8
         TabIndex        =   17
         Tag             =   "Importe Penalización|N|S|||rhisfruta|imppenal|#,##0.00||"
         Text            =   "123"
         Top             =   1590
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   5925
         MaxLength       =   7
         TabIndex        =   10
         Tag             =   "Peso Bruto|N|N|||rhisfruta|kilosbru|###,##0||"
         Top             =   1050
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   8715
         MaxLength       =   7
         TabIndex        =   12
         Tag             =   "Peso Neto|N|N|||rhisfruta|kilosnet|###,##0||"
         Top             =   1050
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   7320
         MaxLength       =   7
         TabIndex        =   11
         Tag             =   "Nro.Cajas|N|N|||rhisfruta|numcajon|###,##0||"
         Top             =   1050
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "Recolectado|N|N|0|1|rhisfruta|recolect||N|"
         Top             =   450
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   5910
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Tipo Entrada|N|N|0|3|rhisfruta|tipoentr||N|"
         Top             =   450
         Width           =   1380
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Impreso"
         Height          =   195
         Index           =   0
         Left            =   4650
         TabIndex        =   2
         Tag             =   "Impreso|N|N|0|1|rhisfruta|impreso|0||"
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   1170
         MaxLength       =   8
         TabIndex        =   5
         Tag             =   "Campo|N|N|0|99999999|rhisfruta|codcampo|00000000||"
         Text            =   "Text1"
         Top             =   1575
         Width           =   1050
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Socio|N|N|0|999999|rhisfruta|codsocio|000000||"
         Text            =   "Text1"
         Top             =   1215
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2250
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   1215
         Width           =   3570
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Albarán|F|N|||rhisfruta|fecalbar|dd/mm/yyyy||"
         Top             =   450
         Width           =   1065
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2250
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   855
         Width           =   3570
      End
      Begin VB.Label Label1 
         Caption         =   "Contrato"
         Enabled         =   0   'False
         Height          =   255
         Index           =   11
         Left            =   10080
         TabIndex        =   99
         ToolTipText     =   "Buscar Campo"
         Top             =   1920
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   7020
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Orden"
         Height          =   255
         Index           =   6
         Left            =   4470
         TabIndex        =   98
         ToolTipText     =   "Buscar Campo"
         Top             =   1950
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Transp."
         Height          =   255
         Index           =   5
         Left            =   10155
         TabIndex        =   96
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Transportado por"
         Height          =   255
         Index           =   0
         Left            =   8700
         TabIndex        =   95
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Precio Estimado"
         Height          =   255
         Left            =   10110
         TabIndex        =   50
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   49
         ToolTipText     =   "Buscar Campo"
         Top             =   1950
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Zona"
         Height          =   255
         Index           =   3
         Left            =   5940
         TabIndex        =   48
         ToolTipText     =   "Buscar Campo"
         Top             =   1950
         Width           =   810
      End
      Begin VB.Label Label18 
         Caption         =   "Imp.Transporte"
         Height          =   255
         Left            =   5925
         TabIndex        =   42
         Top             =   1365
         Width           =   1275
      End
      Begin VB.Label Label6 
         Caption         =   "Imp.Acarreo"
         Height          =   255
         Left            =   7320
         TabIndex        =   41
         Top             =   1365
         Width           =   1140
      End
      Begin VB.Label Label4 
         Caption         =   "Imp.Recolección"
         Height          =   255
         Left            =   8715
         TabIndex        =   40
         Top             =   1365
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Imp.Penalización"
         Height          =   255
         Left            =   10110
         TabIndex        =   39
         Top             =   1365
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Bruto"
         Height          =   255
         Index           =   10
         Left            =   5925
         TabIndex        =   38
         Top             =   825
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Neto"
         Height          =   255
         Index           =   9
         Left            =   8760
         TabIndex        =   37
         Top             =   825
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cajas"
         Height          =   255
         Index           =   7
         Left            =   7365
         TabIndex        =   36
         Top             =   825
         Width           =   1185
      End
      Begin VB.Label Label11 
         Caption         =   "Recolectado"
         Height          =   255
         Index           =   3
         Left            =   7320
         TabIndex        =   35
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Entrada"
         Height          =   255
         Index           =   1
         Left            =   5910
         TabIndex        =   34
         Top             =   210
         Width           =   945
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   900
         ToolTipText     =   "Buscar T.Mercado"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Campo"
         Height          =   255
         Index           =   14
         Left            =   225
         TabIndex        =   33
         ToolTipText     =   "Buscar Campo"
         Top             =   1620
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   900
         ToolTipText     =   "Buscar Socio"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         Height          =   255
         Index           =   8
         Left            =   225
         TabIndex        =   32
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Alb"
         Height          =   255
         Index           =   29
         Left            =   1170
         TabIndex        =   30
         Top             =   180
         Width           =   780
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1980
         Picture         =   "frmManHcoFruta.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Variedad"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   28
         Top             =   900
         Width           =   660
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   900
         ToolTipText     =   "Buscar Variedad"
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Albarán"
         Height          =   255
         Index           =   28
         Left            =   225
         TabIndex        =   27
         Top             =   180
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5625
      Left            =   90
      TabIndex        =   51
      Top             =   2940
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   9922
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   8
      TabHeight       =   520
      ForeColor       =   9771019
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Entradas"
      TabPicture(0)   =   "frmManHcoFruta.frx":0097
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtAux3(11)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FrameAux0"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Gastos"
      TabPicture(1)   =   "frmManHcoFruta.frx":00B3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux1"
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
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
         Height          =   5025
         Left            =   -74880
         TabIndex        =   84
         Top             =   330
         Width           =   10650
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   5
            Left            =   6120
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   94
            Text            =   "Importe total"
            Top             =   4590
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   8
            Left            =   4545
            MaxLength       =   10
            TabIndex        =   90
            Tag             =   "Importe|N|S|||rhisfruta_gastos|importe|###,##0.00|N|"
            Text            =   "Importe"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   9
            Left            =   1830
            MaxLength       =   2
            TabIndex        =   89
            Tag             =   "Cod.Gasto|N|N|0|99|rhisfruta_gastos|codgasto|00||"
            Text            =   "Ga"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   1
            Left            =   2565
            MaskColor       =   &H00000000&
            TabIndex        =   88
            ToolTipText     =   "Buscar Gasto"
            Top             =   1665
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   2790
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   87
            Text            =   "Nombre gasto"
            Top             =   1665
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   10
            Left            =   180
            MaxLength       =   7
            TabIndex        =   86
            Tag             =   "Num.Albaran|N|N|||rhisfruta_gastos|numalbar||S|"
            Text            =   "numalbar"
            Top             =   1665
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   11
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   85
            Tag             =   "Linea|N|N|0|999999|rhisfruta_gastos|numlinea|000000|S|"
            Text            =   "linea"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   2
            Left            =   30
            TabIndex        =   91
            Top             =   210
            Width           =   1110
            _ExtentX        =   1958
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
         Begin MSDataGridLib.DataGrid DataGrid4 
            Bindings        =   "frmManHcoFruta.frx":00CF
            Height          =   3810
            Left            =   30
            TabIndex        =   92
            Top             =   660
            Width           =   8475
            _ExtentX        =   14949
            _ExtentY        =   6720
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adoaux 
            Height          =   330
            Index           =   1
            Left            =   1485
            Top             =   240
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
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
         Begin VB.Label Label3 
            Caption         =   "TOTAL  GASTOS:"
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
            Height          =   285
            Left            =   4140
            TabIndex        =   93
            Top             =   4620
            Width           =   1875
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
         Height          =   2715
         Left            =   5640
         TabIndex        =   74
         Top             =   2805
         Width           =   5700
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   5
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   80
            Tag             =   "Variedad|N|N|0|999999|rhisfruta_clasif|codvarie|000000|S|"
            Text            =   "varie"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   180
            MaxLength       =   7
            TabIndex        =   79
            Tag             =   "Num.Albaran|N|N|||rhisfruta_clasif|numalbar||S|"
            Text            =   "numalbar"
            Top             =   1665
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   6
            Left            =   2790
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   78
            Text            =   "Nombre Calidad"
            Top             =   1665
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   0
            Left            =   2565
            MaskColor       =   &H00000000&
            TabIndex        =   77
            ToolTipText     =   "Buscar Calidad"
            Top             =   1665
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   6
            Left            =   1845
            MaxLength       =   2
            TabIndex        =   75
            Tag             =   "Calidad|N|N|0|99|rhisfruta_clasif|codcalid|00|S|"
            Text            =   "Calidad"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   7
            Left            =   4545
            MaxLength       =   7
            TabIndex        =   76
            Tag             =   "Kilos Netos|N|S|||rhisfruta_clasif|kilosnet|###,##0|N|"
            Text            =   "Kilos Neto"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   135
            TabIndex        =   81
            Top             =   315
            Width           =   1110
            _ExtentX        =   1958
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
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "frmManHcoFruta.frx":00E4
            Height          =   1710
            Left            =   135
            TabIndex        =   82
            Top             =   720
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   3016
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adoaux 
            Height          =   330
            Index           =   0
            Left            =   1485
            Top             =   360
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
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
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
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
         Height          =   2490
         Left            =   90
         TabIndex        =   59
         Top             =   330
         Width           =   11415
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   6
            Left            =   4410
            MaxLength       =   7
            TabIndex        =   65
            Tag             =   "Kilos Neto|N|N|||rhisfruta_entradas|kilosnet|###,##0||"
            Text            =   "KilosNet"
            Top             =   1155
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   10
            Left            =   7860
            MaxLength       =   30
            TabIndex        =   71
            Tag             =   "Imp.Penal|N|S|||rhisfruta_entradas|imppenal|##,##0.00|N|"
            Text            =   "Imp.Penal"
            Top             =   1125
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   9
            Left            =   7050
            MaxLength       =   30
            TabIndex        =   70
            Tag             =   "Imp.Recol|N|S|||rhisfruta_entradas|imprecol|##,##0.00||"
            Text            =   "Imp.Recol"
            Top             =   1125
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   8
            Left            =   6330
            MaxLength       =   30
            TabIndex        =   69
            Tag             =   "Imp.Acarreo|N|S|||rhisfruta_entradas|impacarr|##,##0.00||"
            Text            =   "Imp.Acar"
            Top             =   1140
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   7
            Left            =   5745
            MaxLength       =   30
            TabIndex        =   68
            Tag             =   "Imp.Tranporte|N|S|||rhisfruta_entradas|imptrans|#,##0.00||"
            Text            =   "Imp.Trans"
            Top             =   1140
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   5
            Left            =   3825
            MaxLength       =   7
            TabIndex        =   64
            Tag             =   "Num.Cajon|N|N|||rhisfruta_entradas|numcajon|###,##0|N|"
            Text            =   "numCaj"
            Top             =   1155
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   3105
            MaxLength       =   7
            TabIndex        =   63
            Tag             =   "Kilos Brutos|N|N|||rhisfruta_entradas|kilosbru|###,##0||"
            Text            =   "KilosBru"
            Top             =   1155
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   2430
            MaxLength       =   30
            TabIndex        =   62
            Tag             =   "Hora Entrada|FHH|N|||rhisfruta_entradas|horaentr|hh:mm:ss|N|"
            Text            =   "HoraEnt"
            Top             =   1155
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   495
            MaxLength       =   7
            TabIndex        =   66
            Tag             =   "Num.Albaran|N|N|||rhisfruta_entradas|numalbar|0000000|S|"
            Text            =   "numalba"
            Top             =   1155
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   1170
            MaxLength       =   7
            TabIndex        =   60
            Tag             =   "Num.Nota|N|N|||rhisfruta_entradas|numnotac|0000000|S|"
            Text            =   "nota"
            Top             =   1155
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   61
            Tag             =   "Fecha Ent.|F|N|||rhisfruta_entradas|fechaent|dd/mm/yyyy|N|"
            Text            =   "FecEnt"
            Top             =   1155
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   12
            Left            =   4950
            MaxLength       =   12
            TabIndex        =   67
            Tag             =   "Pr.Estimado|N|S|||rhisfruta_entradas|prestimado|###,##0.0000||"
            Text            =   "Pr.Estimado"
            Top             =   1140
            Visible         =   0   'False
            Width           =   765
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   60
            TabIndex        =   72
            Top             =   210
            Width           =   1110
            _ExtentX        =   1958
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
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "frmManHcoFruta.frx":00F9
            Height          =   1725
            Left            =   60
            TabIndex        =   73
            Top             =   675
            Width           =   11190
            _ExtentX        =   19738
            _ExtentY        =   3043
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   1995
         Left            =   90
         TabIndex        =   53
         Top             =   3525
         Width           =   5475
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   300
            MaxLength       =   7
            TabIndex        =   57
            Tag             =   "Num.Albaran|N|N|||rhisfruta_incidencia|numalbar|0000000|S|"
            Text            =   "numalbar"
            Top             =   1425
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   1260
            MaxLength       =   7
            TabIndex        =   56
            Tag             =   "Num.Nota|N|N|||rhisfruta_incidencia|numnotac|0000000|N|"
            Text            =   "numnota"
            Top             =   1425
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   2070
            MaxLength       =   4
            TabIndex        =   55
            Tag             =   "Cod.Incidencia|N|N|||rhisfruta_incidencia|codincid|0000|N|"
            Text            =   "codincid"
            Top             =   1425
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   3150
            MaxLength       =   12
            TabIndex        =   54
            Text            =   "nomincid"
            Top             =   1425
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmManHcoFruta.frx":010E
            Height          =   1395
            Left            =   60
            TabIndex        =   58
            Top             =   300
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   2461
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtAux3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   11
         Left            =   1305
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Tag             =   "Observaciones|T|S|||rhisfruta_entradas|observac|||"
         Text            =   "frmManHcoFruta.frx":0123
         Top             =   2850
         Visible         =   0   'False
         Width           =   4260
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   83
         Top             =   2850
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   90
      TabIndex        =   22
      Top             =   8580
      Width           =   2175
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
         Left            =   240
         TabIndex        =   23
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10710
      TabIndex        =   20
      Top             =   8625
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9540
      TabIndex        =   19
      Top             =   8640
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Añadir"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar fichero Excel"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Importar fichero Excel"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Liquidacion Directa"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   25
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10680
      TabIndex        =   21
      Top             =   8610
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   810
      Top             =   7875
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   630
      Top             =   7920
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2340
      Top             =   8610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "doc"
   End
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   225
      Left            =   2820
      TabIndex        =   101
      Top             =   8610
      Visible         =   0   'False
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.Label lblProgres 
      Height          =   195
      Index           =   0
      Left            =   2850
      TabIndex        =   100
      Top             =   8850
      Width           =   6195
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
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         HelpContextID   =   2
         Shortcut        =   ^I
      End
      Begin VB.Menu mnExportar 
         Caption         =   "E&xportar Excel "
         Shortcut        =   ^X
      End
      Begin VB.Menu mnImportar 
         Caption         =   "Impor&tar Excel"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnLiquidacion 
         Caption         =   "&Liquidación Directa"
         Shortcut        =   ^L
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
Attribute VB_Name = "frmManHcoFruta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public NroAlbaran As String  ' venimos de mantenimineto de socios

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmLHco As frmManLinHcoFruta 'Lineas de entradas de albaranes
Attribute frmLHco.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden 'Form Mto de incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes ' mensajes para sacar campos
Attribute frmMens.VB_VarHelpID = -1

Private WithEvents frmVar As frmComVar 'Form Mto de variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Form Mto de socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'Form Mto de calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmGas As frmManConcepGasto 'Form Mto de conceptos de gastos
Attribute frmGas.VB_VarHelpID = -1
Private WithEvents frmCamp As frmManCampos 'Form Mto de campos
Attribute frmCamp.VB_VarHelpID = -1
Private WithEvents frmMens2 As frmMensajes ' mensajes para password
Attribute frmMens2.VB_VarHelpID = -1
Private WithEvents frmMens3 As frmMensajes ' mensajes para password
Attribute frmMens3.VB_VarHelpID = -1

Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec
Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

' indicamos que las variedades que vamos a mostrar en el formulario no sean del grupo 6
Private CadB1 As String


Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim Clave As String


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim indice As Byte
Dim Facturas As String

' para la opcion modificar de cabecera
Dim SocioAnt As String
Dim VarieAnt As String
Dim CampoAnt As String

Dim Cliente As String
Private BuscaChekc As String


Dim NumNota As String
Dim Variedad As String
Dim Socio As String
Dim Bruto As String
Dim Neto As String
Dim NIF As String
Dim NomSocio As String
Dim FechaEnt As String
Dim HoraEnt As String


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim SociosNoExisten As String
Dim VariedadesNoExisten As String
Dim CalidadesNoExisten As String

Dim Continuar As Boolean



Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Calidades
            Set frmCal = New frmManCalidades
            frmCal.DatosADevolverBusqueda = "2|3|"
            frmCal.ParamVariedad = txtAux(5).Text
            frmCal.CodigoActual = txtAux(6).Text
            frmCal.Show vbModal
            Set frmCal = Nothing
            PonerFoco txtAux(6)
        Case 1 'Conceptos de gastos
            Set frmGas = New frmManConcepGasto
            frmGas.DatosADevolverBusqueda = "0|1|"
            frmGas.CodigoActual = txtAux(9).Text
            frmGas.Show vbModal
            Set frmGas = Nothing
            PonerFoco txtAux(9)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3  'AÑADIR
            If DatosOk Then InsertarCabecera
        

        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaCabecera Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
'                    FormatoDatosTotales
'                    i = Data3.Recordset.AbsolutePosition

' [Monica] quitadas estas dos lineas
'                        PonerCampos
'                        PonerCamposLineas
                   
'                    SituarDataPosicion Data3, CLng(i), ""
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            Select Case ModificaLineas
                Case 1 'afegir llínia
                    InsertarLinea NumTabMto
                Case 2 'modificar llínies
                    ModificarLinea
                    If NumTabMto = 0 Then ComprobarClasificacion
                    PosicionarData
            End Select
            If NumTabMto = 1 Then CalcularTotalGastos
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            
            Select Case NumTabMto
                Case 0
                    ComprobarClasificacion
                            
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid3.AllowAddNew = False
                        If Not Adoaux(0).Recordset.EOF Then Adoaux(0).Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid3"
                    PonerModo 2
                    DataGrid3.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid3.Enabled = True
                    PonerFocoGrid DataGrid3
                Case 1
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid4.AllowAddNew = False
                        If Not Adoaux(1).Recordset.EOF Then Adoaux(1).Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid4"
                    PonerModo 2
                    DataGrid4.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid4.Enabled = True
                    PonerFocoGrid DataGrid4
            End Select
            
    End Select
End Sub

Private Sub BotonLiquidacion()

    Screen.MousePointer = vbHourglass
    
    frmListAnticipos.NumCod = Text1(0).Text
    frmListAnticipos.OpcionListado = 19
    frmListAnticipos.Show vbModal
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("palets", "numpalet")
'    FormateaCampo Text1(0)
    Combo1(0).ListIndex = -1
    Combo1(1).ListIndex = -1
    Combo1(2).ListIndex = -1
    
    If vParamAplic.Cooperativa = 3 Then
        Combo1(0).ListIndex = 0
        Combo1(1).ListIndex = 1
        Combo1(2).ListIndex = 1
    End If
    
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    
    ' los pesos y cajas se ponen a cero
    Text1(5).Text = 0
    Text1(6).Text = 0
    Text1(7).Text = 0
    Text1(13).Text = 0
    
    LimpiarDataGrids
    
    ' el campo de total de gastos tiene que estar limpio
    Text2(5).Text = ""
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGrid2.Top
        If DataGrid2.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid2"
        
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB1
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select rhisfruta.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & " where " & CadB1
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean

'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
    VarieAnt = Text1(2).Text
    SocioAnt = Text1(3).Text
    CampoAnt = Text1(4).Text
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
        
End Sub


Private Sub BotonModificarLinea(Index As Integer)
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte

    On Error GoTo eModificarLinea


'     'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then
'        TerminaBloquear
'        Exit Sub
'    End If

'    'Si no estaba modificando lineas salimos
'    'Es decir, si estaba insertando linea no podemos hacer otra cosa
'    If ModificaLineas = 1 Then  '1= Insertar
'        TerminaBloquear
'        Exit Sub
'    End If
    
    Select Case NumTabMto
        Case 0
            If Adoaux(0).Recordset.EOF Then
                TerminaBloquear
                Exit Sub
            End If
    
        Case 1
            If Adoaux(1).Recordset.EOF Then
                TerminaBloquear
                Exit Sub
            End If
    End Select
       
    ModificaLineas = 2
    
'    NumTabMto = Index
'    If Index = 2 Then NumTabMto = 3
    PonerModo 5, Index
 

    
    Select Case NumTabMto
        Case 0 ' rhisfruta_clasificacion
            vWhere = ObtenerWhereCP(False)
            If Not BloqueaRegistro("rhisfruta_clasif", vWhere) Then
                TerminaBloquear
                Exit Sub
            End If
            If DataGrid3.Bookmark < DataGrid3.FirstRow Or DataGrid3.Bookmark > (DataGrid3.FirstRow + DataGrid3.VisibleRows - 1) Then
                J = DataGrid3.Bookmark - DataGrid3.FirstRow
                DataGrid3.Scroll 0, J
                DataGrid3.Refresh
            End If
            
        '    anc = ObtenerAlto(Me.DataGrid1)
            anc = DataGrid3.Top
            If DataGrid3.Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 10
            End If
        
            For J = 4 To 6
                txtAux(J).Text = DataGrid3.Columns(J - 4).Text
            Next J
            Text2(6).Text = DataGrid3.Columns(3).Text
            
            txtAux(7).Text = DataGrid3.Columns(4).Text
            
            ModificaLineas = 2 'Modificar
            LLamaLineas ModificaLineas, anc, "DataGrid3"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid3.Enabled = True
            
'            PonerBotonCabecera False
            PonerFoco txtAux(7)
            Me.DataGrid3.Enabled = False


        Case 1 ' rhisfruta_gastos
            vWhere = ObtenerWhereCP(False)
            If Not BloqueaRegistro("rhisfruta_gastos", vWhere) Then
                TerminaBloquear
                Exit Sub
            End If
            If DataGrid4.Bookmark < DataGrid4.FirstRow Or DataGrid4.Bookmark > (DataGrid4.FirstRow + DataGrid4.VisibleRows - 1) Then
                J = DataGrid4.Bookmark - DataGrid4.FirstRow
                DataGrid4.Scroll 0, J
                DataGrid4.Refresh
            End If
            
        '    anc = ObtenerAlto(Me.DataGrid1)
            anc = DataGrid4.Top
            If DataGrid4.Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGrid4.RowTop(DataGrid4.Row) + 10
            End If
        
            txtAux(10).Text = DataGrid4.Columns(0).Text
            txtAux(11).Text = DataGrid4.Columns(1).Text
            txtAux(9).Text = DataGrid4.Columns(2).Text
            
            Text2(1).Text = DataGrid4.Columns(3).Text
            
            txtAux(8).Text = DataGrid4.Columns(4).Text
            
            ModificaLineas = 2 'Modificar
            LLamaLineas ModificaLineas, anc, "DataGrid4"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid4.Enabled = True
            
'            PonerBotonCabecera False
            PonerFoco txtAux(8)
            Me.DataGrid4.Enabled = False

    End Select
    
eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean
    
    Select Case grid
        Case "DataGrid1"
            DeseleccionaGrid Me.DataGrid1
            'PonerModo xModo + 1
    
            b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    
            For jj = 0 To txtAux.Count - 1
                If jj = 4 Or jj = 6 Or jj = 7 Or jj = 8 Then
                    txtAux(jj).Height = DataGrid1.RowHeight
                    txtAux(jj).Top = alto
                    txtAux(jj).visible = b
                End If
            Next jj
            
        Case "DataGrid2"
            DeseleccionaGrid Me.DataGrid2
            b = (xModo = 1)
             For jj = 1 To 10 '09/09/2010
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto - 210 '200
                txtAux3(jj).visible = b
            Next jj
            
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            b = (xModo = 1 Or xModo = 2)
             For jj = 6 To 7
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = b
            Next jj
            btnBuscar(0).Height = DataGrid3.RowHeight - 10
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = b
            Text2(6).Height = DataGrid3.RowHeight - 10
            Text2(6).Top = alto + 5
            Text2(6).visible = b
            
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = b
    
        Case "DataGrid4"
            DeseleccionaGrid Me.DataGrid4
            b = (xModo = 1 Or xModo = 2)
             For jj = 8 To 9
                txtAux(jj).Height = DataGrid4.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = b
            Next jj
            btnBuscar(1).Height = DataGrid4.RowHeight - 10
            btnBuscar(1).Top = alto + 5
            btnBuscar(1).visible = b
            Text2(1).Height = DataGrid4.RowHeight - 10
            Text2(1).Top = alto + 5
            Text2(1).visible = b
   
    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    '[Monica]02/12/2014: en el caso de Picassent quieren una clave de control cuando vayan a modificar o a eliminar
    If vParamAplic.Cooperativa = 2 Then
        Clave = ""
        
        Set frmMens2 = New frmMensajes
        
        frmMens2.OpcionMensaje = 59
        frmMens2.Caption = "Clave de Acceso"
        frmMens2.Show vbModal
    
        If Clave <> vParamAplic.ClaveAcceso Then
            MsgBox "Clave incorrecta.", vbExclamation
            Exit Sub
        End If
        Set frmMens2 = Nothing
        
        Clave = ""
    End If
    
    
    
    If Not ContinuarSiAlbaranImpreso(Data1.Recordset!numalbar) Then Exit Sub
    
    
    Cad = "Cabecera de Albaranes." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Albarán:            "
    Cad = Cad & vbCrLf & "Nº Albarán:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
'        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
'        NumPedElim = Data1.Recordset.Fields(1).Value
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
        
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        vTipoMov.DevolverContador CodTipoMov, NumPedElim
'        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid2.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

'    If LastCol = -1 Then Exit Sub

    'Datos de la tabla albaran_calibres
    If Not Data3.Recordset.EOF Then
        'Datos de la tabla rhisfruta_incidencia
        CargaGrid DataGrid1, Data2, True
        txtAux3(11).Text = DBLet(Data3.Recordset!Observac, "T")
    Else
        'Datos de la tabla rhisfruta_incidencia
        CargaGrid DataGrid1, Data2, False
        txtAux3(11).Text = ""
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If NroAlbaran <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    
'    PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 15
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(4).Image = 3   'Insertar
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(8).Image = 10  'Impresión de albaran
        .Buttons(9).Image = 34  ' Exportar a excel
        .Buttons(10).Image = 35  ' Importar a excel
        .Buttons(11).Image = 13  ' Liquidacion directa
        .Buttons(12).Image = 11  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For kCampo = 0 To ToolAux.Count - 1
        With Me.ToolAux(kCampo)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next kCampo
   ' ***********************************
   'IMAGES para zoom
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i
   
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    
    CodTipoMov = "ALF" 'hcoCodTipoM
    VieneDeBuscar = False
    
        
    '## A mano
    NombreTabla = "rhisfruta"
    NomTablaLineas = "rhisfruta_entradas" 'Tabla de entradas
    Ordenacion = " ORDER BY rhisfruta.numalbar"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    ' registros de variedades que no sean del grupo 6
    CadB1 = "rhisfruta.codvarie in (select codvarie from variedades, productos, grupopro where grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 "
    CadB1 = CadB1 & " and variedades.codprodu = productos.codprodu and productos.codgrupo = grupopro.codgrupo ) "
    
    '[Monica]03/11/2011: Si es quatretonda permitimos que sean variedades de almazara / bodega
    If vParamAplic.Cooperativa = 7 Then
        CadB1 = "rhisfruta.codvarie in (select codvarie from variedades, productos, grupopro where grupopro.codgrupo <> 6 "
        CadB1 = CadB1 & " and variedades.codprodu = productos.codprodu and productos.codgrupo = grupopro.codgrupo ) "
    End If
    
    
    CadenaConsulta = "select * from rhisfruta "
    If NroAlbaran <> "" Then
        CadenaConsulta = CadenaConsulta & " where numalbar = " & DBSet(NroAlbaran, "N")
    Else
        CadenaConsulta = CadenaConsulta & " where numalbar = -1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
'        CargaGrid DataGrid1, Data2, False
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    SSTab1.Tab = 0
    
    If DatosADevolverBusqueda <> "" Then
        Text1(0).Text = DatosADevolverBusqueda
        HacerBusqueda
        SSTab1.Tab = 1
    Else
        PonerModo 0
    End If
    
    If vParamAplic.Cooperativa = 16 Then
        Text1(14).Enabled = True
        Text1(14).visible = True
        Label1(11).visible = True
    Else
        Text4(3).Width = 4125
    End If
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    Me.Combo1(2).ListIndex = -1
    Me.Check1(0).Value = 0
'    Label2(2).Caption = ""
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If imgFec(0).Tag < 2 Then
        Text1(CByte(imgFec(0).Tag) + 1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Else
        Text1(CByte(imgFec(0).Tag) + 8).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub

' devolvemos la linea del datagrid en donde estabamos
Private Sub frmLAlb_DatoSeleccionado(CadenaSeleccion As String)
Dim vWhere As String
             
   PonerCamposLineas
   
   If CadenaSeleccion = "" Then Exit Sub
             
   vWhere = "(numalbar = " & RecuperaValor(CadenaSeleccion, 1) & " and numlinea = " & RecuperaValor(CadenaSeleccion, 2) & ")"
   SituarDataMULTI Data3, vWhere, "" ', Indicador
   
   PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
   PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
   

End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Calidades
    txtAux(6).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'Cod Calidad
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Calidad
End Sub

Private Sub frmCamp_DatoSeleccionado(CadenaSeleccion As String)
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de campo
    FormateaCampo Text1(4)
'    If EstaCampoDeAlta(Text1(4).Text) Then
'        PonerDatosCampo Text1(4).Text
'    Else
'        MsgBox "El campo no está dado de alta. Reintroduzca.", vbExclamation
'        Text1(4).Text = ""
'        PonerFoco Text1(4)
'    End If
End Sub

Private Sub frmGas_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Conceptos de gastos
    txtAux(9).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'Cod concepto de gasto
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmLHco_DatoSeleccionado(CadenaSeleccion As String)
Dim vWhere As String
             
   PonerCamposLineas
   
   If CadenaSeleccion = "" Then Exit Sub
             
   vWhere = "(numalbar = " & RecuperaValor(CadenaSeleccion, 1) & " and numlinea = " & RecuperaValor(CadenaSeleccion, 2) & ")"
   SituarDataMULTI Data3, vWhere, "" ', Indicador
   
   PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
   PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
   
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    PonerDatosCampo Text1(4)
End Sub

Private Sub frmMens2_DatoSeleccionado(CadenaSeleccion As String)
    Clave = CadenaSeleccion
End Sub

Private Sub frmMens3_DatoSeleccionado(CadenaSeleccion As String)
    Continuar = (CadenaSeleccion <> "")
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Variedades
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Variedad
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Variedad
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Socios
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Socios
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           CargaCadenaAyuda vCadena, Index
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Variedad
            indice = 2
            PonerFoco Text1(indice)
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(indice)
        
        Case 1 'Socios
            indice = 3
            PonerFoco Text1(indice)
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(indice)
            
        Case 2 'campos
            indice = 4
            Set frmCamp = New frmManCampos
            frmCamp.DatosADevolverBusqueda = "0|"
            frmCamp.Show vbModal
            Set frmCamp = Nothing
            PonerFoco Text1(indice)
            
    End Select
    
    Screen.MousePointer = vbDefault
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

    If Index < 2 Then
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 1).Text <> "" Then frmC.NovaData = Text1(Index + 1).Text
    Else
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 8).Text <> "" Then frmC.NovaData = Text1(Index + 8).Text
    End If
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    If Index < 2 Then
        PonerFoco Text1(CByte(imgFec(0).Tag) + 1) '<===
    Else
        PonerFoco Text1(CByte(imgFec(0).Tag) + 8) '<===
    End If
    ' ********************************************
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 15
        frmZ.pTitulo = "Observaciones del Albarán"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub



Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    
    
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub

Private Sub mnGenerarAlb_Click()
'Generar Albarán
    
    If Data1.Recordset.EOF Then Exit Sub
    
'    BotonImprimir

End Sub

Private Sub mnExportar_Click()
    If Not CargarCondicion Then Exit Sub
    
    Shell App.Path & "\clasificacion.exe /E|" & vUsu.CadenaConexion & "|", vbNormalFocus
    
End Sub

Private Function CargarCondicion() As Boolean
Dim Sql As String
Dim NFic As Integer

    On Error GoTo eCargarCondicion

    CargarCondicion = False

    If Dir(App.Path & "\condicionsql.txt", vbArchive) <> "" Then Kill (App.Path & "\condicionsql.txt")
        
    If Data1.Recordset.RecordCount = 0 Then
        CargarCondicion = True
        Exit Function
    End If
        
    NFic = FreeFile
    
    Open App.Path & "\condicionsql.txt" For Output As #NFic
    
    Print #NFic, Replace(Data1.RecordSource, "select rhisfruta.* ", "select rhisfruta.numalbar ")
        
    Close #NFic
    
    CargarCondicion = True
    
    Exit Function
    
eCargarCondicion:
    MuestraError Err.Number, "Cargando condición", Err.Description
End Function


Private Sub mnImportar_Click()
    '[Monica]21/11/2016: para el caso de bolbaite se introducen las entradas como las de almazara en ABN
    If vParamAplic.Cooperativa = 14 Then
        lblProgres(0).visible = False
        Pb1.visible = False
        
        ImportacionEntradas
        
        lblProgres(0).visible = False
        Pb1.visible = False
    Else
        Shell App.Path & "\clasificacion.exe /I|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
    End If
End Sub

Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
        
    If Not ContinuarSiAlbaranImpreso(Data1.Recordset!numalbar) Then Exit Sub
    BotonImprimir
End Sub

Private Sub mnGenerarFactura_Click()
'Generacion de factura a partir del albaran aprovechando los precios provisionales
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If BLOQUEADesdeFormulario(Me) Then
        BotonGenerarFactura Data1.Recordset.Fields(0).Value
        TerminaBloquear
    End If
End Sub

Private Sub mnLiquidacion_Click()
Dim Sql As String

    If Data1.Recordset.EOF Then Exit Sub

'    If Combo1(0).ListIndex = 1 Then
'        MsgBox "No se permite liquidar entradas de Venta Vampo." & vbCrLf & vbCrLf & "Vaya por el punto correspondiente.", vbExclamation
'        Exit Sub
'    End If

    Sql = "select count(*) from rfactsoc_albaran where numalbar = " & DBSet(Text1(0).Text, "N")
    If DevuelveValor(Sql) <> 0 Then
        If MsgBox("Este albarán ya está facturado. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
    End If

    BotonLiquidacion
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnModificar_Click()

    '[Monica]02/12/2014: en el caso de Picassent quieren una clave de control cuando vayan a modificar o a eliminar
    If vParamAplic.Cooperativa = 2 Then
        Clave = ""
        
        Set frmMens2 = New frmMensajes
        
        frmMens2.OpcionMensaje = 59
        frmMens2.Caption = "Clave de Acceso"
        frmMens2.Show vbModal
    
        If Clave <> vParamAplic.ClaveAcceso Then
            MsgBox "Clave incorrecta.", vbExclamation
            Exit Sub
        End If
        Set frmMens2 = Nothing
        
        Clave = ""
    End If

    If Not ContinuarSiAlbaranImpreso(Data1.Recordset!numalbar) Then Exit Sub
    
    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then BotonModificarLinea (NumTabMto)
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim Sql As String

    On Error GoTo EBloqueaAlb
    
    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM scafac1 "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea todas las lineas de la factura
Dim Sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM slifac "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
End Function

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
'    If Index = 9 Then HaCambiadoCP = False 'CPostal
'    If Index = 1 And Modo = 1 Then
'        SendKeys "{tab}"
'        Exit Sub
'    End If
    If Index = 3 Then 'codigo de cliente
        Cliente = Text1(Index).Text
    End If
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
Dim cadMen As String
Dim Sql As String

        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha albaran
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index), True
            
        Case 2 'Variedad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "variedades", "nomvarie")
                If Modo = 1 Then Exit Sub
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmComVar
                        frmVar.DatosADevolverBusqueda = "0|1|"
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
                    If (Modo = 3 Or Modo = 4) Then
                        If EsVariedadGrupo6(Text1(Index).Text) Then
                            MsgBox "Esta variedad es del Grupo de Bodega. Revise.", vbExclamation
                            PonerFoco Text1(Index)
                        End If
                        If EsVariedadGrupo5(Text1(Index).Text) Then
                            MsgBox "Esta variedad es del Grupo de Almazara. Revise.", vbExclamation
                            PonerFoco Text1(Index)
                        End If
                    End If
                End If
            End If
    
        Case 3 'Socio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                If Modo = 1 Then Exit Sub
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
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
                    If Not EstaSocioDeAlta(Text1(Index)) Then
                        MsgBox "El socio está dado de baja. Reintroduzca.", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    Else
                        PonerCamposSocioVariedad
                    End If
                End If
            End If
            
        Case 4 'Campo
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then Exit Sub
                Sql = ""
                Sql = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", Text1(Index).Text, "N")
                If Sql = "" Then
                    cadMen = "No existe el Campo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCamp = New frmManCampos
                        frmCamp.DatosADevolverBusqueda = "0|1|"
'                        frmCamp.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCamp.Show vbModal
                        Set frmCamp = Nothing
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
        
    End Select
    If (Index = 2 Or Index = 3 Or Index = 4) And (Modo = 3 Or Modo = 4) Then
        If Not EsCampoSocioVariedad(Text1(4).Text, Text1(3).Text, Text1(2).Text) Then
            MsgBox "El campo no es del socio o de la variedad introducida. Revise.", vbExclamation
        End If
    End If

    
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim cadAux As String
    
'    '--- Laura 12/01/2007
'    cadAux = Text1(5).Text
'    If Text1(4).Text <> "" Then Text1(5).Text = ""
'    '---
    
'    '--- Laura 12/01/2007
'    Text1(5).Text = cadAux
'    '---
'    CadB = ObtenerBusqueda(Me)
    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB & " and " & CadB1
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select rhisfruta.* from " & NombreTabla & " LEFT JOIN rhisfruta_entradas ON rhisfruta.numalbar=rhisfruta_entradas.numalbar "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " and " & CadB1 & " GROUP BY rhisfruta.numalbar " & Ordenacion
'        CadenaConsulta = "select palets.* from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY palets.numpalet " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
    Cad = Cad & "Albaran|rhisfruta.numalbar|N||11·"
    Cad = Cad & "Fecha|rhisfruta.fecalbar|F||14·"
    
    Cad = Cad & "Cod|rhisfruta.codvarie|N||7·" 'ParaGrid(Text1(3), 10, "Cliente")
    Cad = Cad & "Nombre|variedades.nomvarie|N||20·"
    Cad = Cad & "Socio|rhisfruta.codsocio|N||10·" 'ParaGrid(Text1(3), 10, "Cliente")
    Cad = Cad & "Nombre|rsocios.nomsocio|N||28·"
    Cad = Cad & "Campo|rhisfruta.codcampo|N||10·"
    
    Tabla = NombreTabla & " INNER JOIN variedades ON rhisfruta.codvarie=variedades.codvarie "
    Tabla = "(" & Tabla & ") INNER JOIN rsocios ON rhisfruta.codsocio=rsocios.codsocio "
    
    Titulo = "Histórico de Entradas"
    devuelve = "0|"
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vtabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
'        If EsCabecera Then
'            PonerCadenaBusqueda
'            Text1(0).Text = Format(Text1(0).Text, "0000000")
'        End If
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        LLamaLineas Modo, 0, "DataGrid2"
        PonerCampos
    End If


    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCamposLineas()
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafac1 de la factura seleccionada
Dim b As Boolean
Dim b2 As Boolean
Dim i As Integer


    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    For i = 1 To 3
        Select Case i
            Case 1
                CargaGrid DataGrid2, Data3, True
                '++monica
                If Data3.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid1, Data2, True
                Else
                    CargaGrid DataGrid1, Data2, False
                End If
                '++
            Case 2  ' clasificacion
                If Data1.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid3, Adoaux(0), True
                Else
                    CargaGrid DataGrid3, Adoaux(0), False
                End If
            Case 3  ' gastos
                If Data1.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid4, Adoaux(1), True
                Else
                    CargaGrid DataGrid4, Adoaux(1), False
                End If
        End Select
    Next i
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single
Dim b As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    b = PonerCamposForma2(Me, Data1, 2, "Frame2")
    
'    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    
    Text2(2).Text = PonerNombreDeCod(Text1(2), "variedades", "nomvarie", "codvarie", "N") 'variedades
    Text2(3).Text = PonerNombreDeCod(Text1(3), "rsocios", "nomsocio", "codsocio", "N") 'socios
    
    VisualizarDatosCampo Data1.Recordset!Codcampo
    
'    MostrarCadena Text1(3), Text1(4)
    
    Modo = 2
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
'    ActualizarToolbar Modo, Kmodo
    BuscaChekc = ""

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or NroAlbaran <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    For i = 5 To 13
        BloquearTxt Text1(i), Not (Modo = 1)
        Text1(i).Enabled = (Modo = 1)
    Next i
    Me.Check1(0).Enabled = (Modo = 1)
    
    b = (Modo <> 1)
    'Campos Nº Pedido bloqueado y en azul
    BloquearTxt Text1(0), b, True
'    BloquearTxt Text1(3), b 'referencia
    
    
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i
    For i = 0 To btnBuscar.Count - 1
        btnBuscar(i).visible = False
        btnBuscar(i).Enabled = True
    Next i
    Text2(6).visible = False
    Text2(6).Enabled = True
    For i = 1 To 10 '09/09/2010
        BloquearTxt txtAux3(i), True
        txtAux3(i).Enabled = False
    Next i
    For i = 1 To 10
        BloquearTxt txtAux3(i), (Modo <> 1)
        txtAux3(i).Enabled = (Modo = 1)
    Next i
    
    txtAux3(11).visible = True
    txtAux3(11).Enabled = (Modo = 1)
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
    
    Text1(2).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    imgBuscar(0).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    imgBuscar(0).visible = (Modo = 1 Or Modo = 3 Or Modo = 4)
    
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
    Select Case NumTabMto
        Case 0
            BloquearFrameAux Me, "FrameAux0", Modo, NumTabMto
        Case 1
            BloquearFrameAux Me, "FrameAux1", Modo, NumTabMto
    End Select
    
    If indFrame = 1 Then
        txtAux(6).Enabled = (ModificaLineas = 1) And (NumTabMto = 0)
        txtAux(6).visible = (ModificaLineas = 1) And (NumTabMto = 0)
        btnBuscar(0).Enabled = (ModificaLineas = 1) And (NumTabMto = 0)
        btnBuscar(0).visible = (ModificaLineas = 1) And (NumTabMto = 0)
    End If
        
    For i = 8 To 9
        txtAux(i).Enabled = (ModificaLineas = 1) And (NumTabMto = 1)
        txtAux(i).visible = (ModificaLineas = 1) And (NumTabMto = 1)
    Next i
    btnBuscar(1).Enabled = (ModificaLineas = 1) And (NumTabMto = 1)
    btnBuscar(1).visible = (ModificaLineas = 1) And (NumTabMto = 1)
        
    txtAux(8).Enabled = (ModificaLineas = 1 Or ModificaLineas = 2) And (NumTabMto = 1)
    txtAux(8).visible = (ModificaLineas = 1 Or ModificaLineas = 2) And (NumTabMto = 1)
        
    Text2(5).visible = True
    Text2(5).Enabled = True
        
        
    ' ***************************
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scafac
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    '[Monica]24/09/2013: en montifrut no hay campos
    If vParamAplic.Cooperativa <> 12 Then
        If Not EsCampoSocioVariedad(Text1(4).Text, Text1(3).Text, Text1(2).Text) Then
            MsgBox "El campo no es del socio o de la variedad introducida. Revise.", vbExclamation
            PonerFoco Text1(2)
            b = False
        End If
        If Not b Then Exit Function
        
        '[Monica]14/06/2010: el campo esta dado de baja tiene que estar de alta
        If Not EstaCampoDeAlta(Text1(4).Text) Then
            MsgBox "El campo está dado de baja. Revise.", vbExclamation
            PonerFoco Text1(4)
            b = False
        End If
        If Not b Then Exit Function
    End If
    If Modo = 4 Then
        ' si estamos modificando y me han cambiado la variedad vemos si la nueva variedad
        ' tiene las mismas calidades que las lineas de rhisfruta_clasif
        If Text1(2).Text <> VarieAnt Then
            Sql = "select codcalid from rhisfruta_clasif where numalbar = " & Text1(0).Text
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF And b
                Sql2 = "select count(*) from rcalidad where codvarie = " & DBSet(Text1(2).Text, "N")
                Sql2 = Sql2 & " and codcalid = " & DBSet(Rs.Fields(0).Value, "N")
                
                If TotalRegistros(Sql2) = 0 Then b = False
            
                Rs.MoveNext
            Wend
            
            Set Rs = Nothing
            
            If Not b Then
                MsgBox "La variedad introducida no tiene las mismas calidades que la anterior. Revise.", vbExclamation
                b = False
            End If
        End If
    End If
    If Not b Then Exit Function
    
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim i As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For i = 0 To txtAux.Count - 1
        If i = 4 Or i = 6 Or i = 7 Then
            If txtAux(i).Text = "" Then
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
            
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    '[Monica]02/12/2014: en el caso de Picassent quieren una clave de control cuando vayan a modificar o a eliminar
    If vParamAplic.Cooperativa = 2 Then
            Clave = ""
            
            Set frmMens2 = New frmMensajes
            
            frmMens2.OpcionMensaje = 59
            frmMens2.Caption = "Clave de Acceso"
            frmMens2.Show vbModal
        
            If Clave <> vParamAplic.ClaveAcceso Then
                MsgBox "Clave incorrecta.", vbExclamation
                Exit Sub
            End If
            Set frmMens2 = Nothing
            
            Clave = ""
    End If
    
    If Not ContinuarSiAlbaranImpreso(Data1.Recordset!numalbar) Then Exit Sub
    
    
    If BloqueaRegistro(NombreTabla, "numalbar = " & Data1.Recordset!numalbar) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Index
            Case 0 'rhisfruta_entradas
                Select Case Button.Index
                    Case 1 'añadir entrada
                        Set frmLHco = New frmManLinHcoFruta
                        
                        frmLHco.ModoExt = 3
                        frmLHco.Albaran = Data1.Recordset.Fields(0).Value
                        frmLHco.Show vbModal
                    
                        Set frmLHco = Nothing
                        
                    Case 2 'modificar entrada
                        Set frmLHco = New frmManLinHcoFruta
                        
                        frmLHco.ModoExt = 4
                        frmLHco.Albaran = Data3.Recordset.Fields(0).Value
                        frmLHco.Nota = Data3.Recordset.Fields(1).Value
                        frmLHco.Show vbModal
                        
                        Set frmLHco = Nothing
                        
                    Case 3 ' boton eliminar linea de variedades
                        BotonEliminarLinea 0
                    Case Else
                End Select
                CalcularTotales
                PonerCampos
                TerminaBloquear
                
            Case 1 'clasificacion
                NumTabMto = 0
                Select Case Button.Index
                    Case 1
                        BotonAnyadirLinea Index
                    Case 2
                        BotonModificarLinea Index
                    Case 3
                        BotonEliminarLinea Index
                    Case Else
                End Select
            
            Case 2 'gastos
                NumTabMto = 1
                Select Case Button.Index
                    Case 1
                        BotonAnyadirLinea Index
                    Case 2
                        BotonModificarLinea Index
                    Case 3
                        BotonEliminarLinea Index
                    Case Else
                End Select
                
        End Select
        
    End If

End Sub


Private Sub BotonEliminarLinea(Index As Integer)
Dim Cad As String
Dim Sql As String

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    Select Case Index
        Case 0 'entrada
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar la Entrada?"
            Cad = Cad & vbCrLf & "Albarán: " & Data3.Recordset.Fields(0)
            Cad = Cad & vbCrLf & "Nota: " & Data3.Recordset.Fields(1)
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Data3.Recordset.AbsolutePosition
                
                If Not EliminarLinea Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    If SituarDataTrasEliminar(Data3, NumRegElim) Then
                        PonerCampos
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
            Screen.MousePointer = vbDefault
       Case 1 'clasificacion
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar la Calidad?"
            Cad = Cad & vbCrLf & "Albarán: " & Adoaux(0).Recordset.Fields(0)
            Cad = Cad & vbCrLf & "Calidad: " & Adoaux(0).Recordset.Fields(2) & "-" & Adoaux(0).Recordset.Fields(3)
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Adoaux(0).Recordset.AbsolutePosition
                TerminaBloquear
                Sql = "delete from rhisfruta_clasif where numalbar = " & Adoaux(0).Recordset.Fields(0)
                Sql = Sql & " and codvarie = " & Adoaux(0).Recordset.Fields(1)
                Sql = Sql & " and codcalid = " & Adoaux(0).Recordset.Fields(2)
                conn.Execute Sql
                
                SituarDataTrasEliminar Adoaux(0), NumRegElim
                
                CargaGrid DataGrid3, Adoaux(0), True
'                SSTab1.Tab = 1

                ComprobarClasificacion
            End If
            Screen.MousePointer = vbDefault
       
       Case 2 'gastos
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar el Gasto?"
            Cad = Cad & vbCrLf & "Albarán: " & Adoaux(1).Recordset.Fields(0)
            Cad = Cad & vbCrLf & "Código: " & Adoaux(1).Recordset.Fields(2) & "-" & Adoaux(1).Recordset.Fields(3)
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Adoaux(0).Recordset.AbsolutePosition
                TerminaBloquear
                Sql = "delete from rhisfruta_gastos where numalbar = " & Adoaux(1).Recordset.Fields(0)
                Sql = Sql & " and numlinea = " & Adoaux(1).Recordset.Fields(1)
                Sql = Sql & " and codgasto = " & Adoaux(1).Recordset.Fields(2)
                conn.Execute Sql
                
                SituarDataTrasEliminar Adoaux(1), NumRegElim
                
                CargaGrid DataGrid4, Adoaux(1), True
'                SSTab1.Tab = 1

            End If
            Screen.MousePointer = vbDefault
       
       
    End Select
       
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Linea de Albarán", Err.Description

End Sub



'Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    KEYdown KeyCode
'End Sub
'
'Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
'
'Private Sub Text3_LostFocus(Index As Integer)
'    Select Case Index
'        Case 0, 1, 2 'trabajador
'            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
'        Case 3 'cod. envio
'            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "senvio", "nomenvio", "codenvio", "Cod. Envio", "N")
'            If Screen.ActiveControl.TabIndex <> 27 Then PonerFocoBtn Me.cmdAceptar
'        Case 13 'observa 5
'            PonerFocoBtn Me.cmdAceptar
'    End Select
'End Sub
'

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos
        Case 4  'Añadir
            mnNuevo_Click
        Case 5  'Modificar
            mnModificar_Click
        Case 6  'Borrar
            mnEliminar_Click
        Case 8  ' Impresion de albaran
            mnImprimir_Click
        Case 9  'Exportar
            mnExportar_Click
        Case 10 'Importar
            mnImportar_Click
        Case 11 'Liquidacion
            mnLiquidacion_Click
        Case 12   'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub ActualizarToolbar(Modo As Byte, Kmodo As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner

    If (Modo = 5) And (Kmodo <> 5) Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 3
'        Toolbar1.Buttons(5).ToolTipText = "Nuevo Albaran"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 4
        Toolbar1.Buttons(5).ToolTipText = "Modificar Factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
    End If
    If Kmodo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 12
'        Toolbar1.Buttons(5).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 13
        Toolbar1.Buttons(5).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 14
        Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
    End If
End Sub
    
    
'Private Function ModificarLinea() As Boolean
''Modifica un registro en la tabla de lineas de Albaran: slialb
'Dim SQL As String
'Dim vWhere As String
'Dim b As Boolean
'
'    On Error GoTo EModificarLinea
'
'    ModificarLinea = False
'    If Data2.Recordset.EOF Then Exit Function
'
'    vWhere = ObtenerWhereCP(True)
'    vWhere = vWhere & " AND codtipoa='" & Data3.Recordset.Fields!codtipoa & "' "
'    vWhere = vWhere & " AND numalbar=" & Data3.Recordset.Fields!numalbar
'    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
'
'    If DatosOkLinea() Then
'        SQL = "UPDATE slifac SET "
'        SQL = SQL & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
'        SQL = SQL & "precioar = " & DBSet(txtAux(4).Text, "N") & ", "
'        SQL = SQL & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
'        SQL = SQL & "importel = " & DBSet(txtAux(8).Text, "N") & ", "
'        SQL = SQL & "origpre='" & txtAux(5) & "'"
'        SQL = SQL & vWhere
'    End If
'
'    If SQL <> "" Then
'        'actualizar la factura y vencimientos
'        b = ModificarFactura(SQL)
'
'        ModificarLinea = b
'    End If
'
'EModificarLinea:
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description
'        b = False
'    End If
'    ModificarLinea = b
'End Function


Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
    On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not b
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGRid

    b = DataGrid1.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid1"
            Opcion = 1
        Case "DataGrid2"
            Opcion = 2
        Case "DataGrid3" 'clasificacion
            Opcion = 3
        Case "DataGrid4" 'gastos
            Opcion = 4
            
    End Select
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    
    If Opcion = 4 Then CalcularTotalGastos
   
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGRid

    Select Case vDataGrid.Name
        Case "DataGrid1" 'rhisfruta_incidencia
'           SQL = "SELECT numalbar, numnotac, codincid, nomincid
            tots = "N||||0|;N||||0|;"
            tots = tots & "S|txtAux(2)|T|Codigo|800|;"
            tots = tots & "S|txtAux(3)|T|Nombre Incidencia|3910|;"
            arregla tots, DataGrid1, Me
         
         Case "DataGrid2" 'rhisfruta_entradas
'           SQL = "SELECT numalbar, numnotac, fechaent, horaentr, kilosbru, numcajon, kilosnet, prestimado, observac,
'           sql= sql & "imptrans, impacarr, imprecol, imppenal
            tots = "N||||0|;"
            tots = tots & "S|txtAux3(1)|T|Nota|750|;"
            tots = tots & "S|txtAux3(2)|T|Fecha Ent|1000|;S|txtAux3(3)|T|Hora|750|;S|txtAux3(4)|T|Bruto|1000|;S|txtAux3(5)|T|Cajones|800|;"
            tots = tots & "S|txtAux3(6)|T|Neto|1000|;S|txtAux3(12)|T|Pr.Estim|1000|;N||||0|;S|txtAux3(7)|T|Transporte|1050|;S|txtAux3(8)|T|Acarreo|900|;S|txtAux3(9)|T|Recolección|1150|;S|txtAux3(10)|T|Penalización|1150|;"
            arregla tots, DataGrid2, Me
            
         Case "DataGrid3" 'rhisfruta_clasif
'       SQL = SELECT albaran_envase.numalbar, numlinea, albaran_envase.codartic, sartic.nomartic, sartic.codtipar, stipar.nomtipar, "
'             albaran_envase.tipomovi, CASE albaran_envase.tipomovi WHEN 0 THEN ""Salida"" WHEN 1 THEN ""Entrada"" END, albaran_envase.cantidad "
            tots = "N||||0|;N||||0|;"
            tots = tots & "S|txtAux(6)|T|Calidad|800|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|Text2(6)|T|Nombre|3000|;"
            tots = tots & "S|txtAux(7)|T|Kilos Neto|1100|;"
            arregla tots, DataGrid3, Me
    
         Case "DataGrid4" 'rhisfruta_gastos
'       SQL = SELECT rhisfruta_gastos.numalbar, rhisfruta_gastos.numlinea, rhisfruta_gastos.codgasto, rconcepgasto.nomgasto, rhisfruta_gastos.importe "
            tots = "N||||0|;N||||0|;"
            tots = tots & "S|txtAux(9)|T|Código|800|;S|btnBuscar(1)|B|||;"
            tots = tots & "S|Text2(1)|T|Descripción|5000|;"
            tots = tots & "S|txtAux(8)|T|Importe|2100|;"
            arregla tots, DataGrid4, Me
            
            DataGrid4.Columns(2).Alignment = dbgLeft
            DataGrid4.Columns(3).Alignment = dbgLeft
            DataGrid4.Columns(4).Alignment = dbgRight
    
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub




Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub

    Select Case Index
        Case 6 ' calidad
            If txtAux(Index) <> "" Then
                Text2(6) = DevuelveDesdeBDNew(cAgro, "rcalidad", "nomcalid", "codvarie", txtAux(5), "N", , "codcalid", txtAux(6).Text, "N")
                If Text2(6).Text = "" Then
                    cadMen = "No existe la Calidad: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCal = New frmManCalidades
                        frmCal.DatosADevolverBusqueda = "2|3|"
                        frmCal.ParamVariedad = txtAux(5).Text
                        frmCal.NuevoCodigo = txtAux(6).Text
                        TerminaBloquear
                        frmCal.Show vbModal
                        Set frmCal = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        PonerFoco txtAux(6)
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                Text2(6).Text = ""
            End If

        Case 7 'peso neto
            If txtAux(Index) <> "" Then
                If PonerFormatoEntero(txtAux(Index)) Then cmdAceptar.SetFocus
            End If

    
        Case 9 ' nombre de gastos
            If txtAux(Index) <> "" Then
                Text2(1) = DevuelveDesdeBDNew(cAgro, "rconcepgasto", "nomgasto", "codgasto", txtAux(9), "N")
                If Text2(1).Text = "" Then
                    cadMen = "No existe el Concepto de Gasto: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmGas = New frmManConcepGasto
                        frmGas.DatosADevolverBusqueda = "0|1|"
                        frmGas.NuevoCodigo = txtAux(6).Text
                        TerminaBloquear
                        frmGas.Show vbModal
                        Set frmGas = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        PonerFoco txtAux(Index)
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    If EsGastodeFactura(txtAux(9).Text) = True Then
                        MsgBox "Este concepto de gasto es de factura. Reintroduzca.", vbExclamation
                        PonerFoco txtAux(Index)
                    End If
                End If
            Else
                Text2(1).Text = ""
            End If
    
        Case 8 ' importe
            If txtAux(Index) <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 3) Then cmdAceptar.SetFocus
            End If
        
    
    
    End Select
    
End Sub




Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en tablas de cabecera de albaran
    '------------------------------------------
    Sql = " " & ObtenerWhereCP(True)
    
    'Lineas de clasificacion (rhisfruta_clasif)
    conn.Execute "Delete from rhisfruta_clasif " & Sql
    
    'Lineas de incidencias de notas (rhisfruta_incidencia)
    conn.Execute "Delete from rhisfruta_incidencia " & Sql
    
    'Lineas de entradas (rhisfruta_entradas)
    conn.Execute "Delete from rhisfruta_entradas " & Sql

    'Lineas de gastos si los hay
    conn.Execute "Delete from rhisfruta_gastos " & Sql
    
    'Cabecera de albaran (rhisfruta)
    conn.Execute "Delete from " & NombreTabla & Sql
    
    'Decrementar contador si borramos el ult. palet
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador "ALF", Val(Text1(0).Text)
    Set vTipoMov = Nothing
    
    b = True
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Albarán", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Function EliminarLinea() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data3.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    'Eliminar en tablas de rhisfruta_incidencia y rhisfruta_entradas
    '------------------------------------------
    Sql = " where numalbar = " & Data3.Recordset.Fields(0)
    Sql = Sql & " and numnotac = " & Data3.Recordset.Fields(1)

    'Lineas de incidencia (rhisfruta_incidencia)
    conn.Execute "Delete from rhisfruta_incidencia " & Sql

    'Lineas de entradas
    conn.Execute "Delete from rhisfruta_entradas " & Sql
    
    '[Monica]10/09/2012: Mogente si solo hay una calidad actualizamos la clasificacion
    '                    para optimizar la entrada
    ActualizarClasificacionHco Text1(0).Text, Text1(2).Text

    
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Entrada del Albarán ", Err.Description & " " & Mens
        conn.RollbackTrans
        EliminarLinea = False
    Else
        conn.CommitTrans
        EliminarLinea = True
    End If
End Function

Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid2, Data3, False 'entradas e incidencias
    CargaGrid DataGrid1, Data2, False
    CargaGrid DataGrid3, Me.Adoaux(0), False 'clasificacion
    CargaGrid DataGrid4, Me.Adoaux(1), False 'gastos
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = " numalbar= " & Text1(0).Text 'Data1.Recordset!numalbar 'Text1(0).Text
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    Select Case Opcion
    Case 1  ' incidencias
        Sql = "SELECT rhisfruta_incidencia.numalbar, rhisfruta_incidencia.numnotac, rhisfruta_incidencia.codincid, rincidencia.nomincid "
        Sql = Sql & " FROM rhisfruta_incidencia, rincidencia WHERE rhisfruta_incidencia.codincid = rincidencia.codincid "
    Case 2  'entradas
        Sql = "SELECT rhisfruta_entradas.numalbar, rhisfruta_entradas.numnotac, rhisfruta_entradas.fechaent, rhisfruta_entradas.horaentr, "
        Sql = Sql & " rhisfruta_entradas.kilosbru, rhisfruta_entradas.numcajon , rhisfruta_entradas.kilosnet,rhisfruta_entradas.prestimado, rhisfruta_entradas.observac, "
        Sql = Sql & " rhisfruta_entradas.imptrans, rhisfruta_entradas.impacarr, rhisfruta_entradas.imprecol, rhisfruta_entradas.imppenal "
        Sql = Sql & " FROM rhisfruta_entradas " 'lineas de entradas del albaran
        Sql = Sql & " WHERE 1=1 "
    Case 3  'clasificacion
        Sql = "SELECT rhisfruta_clasif.numalbar, rhisfruta_clasif.codvarie, rhisfruta_clasif.codcalid, rcalidad.nomcalid, rhisfruta_clasif.kilosnet "
        Sql = Sql & " FROM rhisfruta_clasif, rcalidad "
        Sql = Sql & " WHERE rhisfruta_clasif.codcalid = rcalidad.codcalid and  "
        Sql = Sql & " rhisfruta_clasif.codvarie = rcalidad.codvarie "
    Case 4  'gastos
        Sql = "SELECT rhisfruta_gastos.numalbar, rhisfruta_gastos.numlinea, rhisfruta_gastos.codgasto, rconcepgasto.nomgasto, rhisfruta_gastos.importe "
        Sql = Sql & " FROM rhisfruta_gastos, rconcepgasto "
        Sql = Sql & " WHERE rhisfruta_gastos.codgasto = rconcepgasto.codgasto "
    End Select
    
    If enlaza Then
        Sql = Sql & " and " & ObtenerWhereCP(False)
        If Opcion = 1 Then Sql = Sql & " AND numnotac=" & Data3.Recordset.Fields!Numnotac
    Else
        Sql = Sql & " and numalbar = -1"
    End If
    Sql = Sql & " ORDER BY numalbar"
    If Opcion = 1 Then Sql = Sql & ", numnotac "
    If Opcion = 4 Then Sql = Sql & ", numlinea "
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

        b = ((Modo = 2) Or (Modo = 0)) And (NroAlbaran = "") 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(1).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(2).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(4).Enabled = b
        Me.mnNuevo.Enabled = b
        
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (NroAlbaran = "") 'And Not (Check1(0).Value = 1)
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnEliminar.Enabled = b
        'Impresión de albaran
        Toolbar1.Buttons(8).Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
        Me.mnImprimir.Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
        
        ' exportar a excel
        Toolbar1.Buttons(9).Enabled = ((Modo = 2) Or (Modo = 0)) And (NroAlbaran = "") And UCase(Dir(App.Path & "\controlclas.cfg")) = UCase("controlclas.cfg")
        Me.mnExportar.Enabled = ((Modo = 2) Or (Modo = 0)) And (NroAlbaran = "") And UCase(Dir(App.Path & "\controlclas.cfg")) = UCase("controlclas.cfg")
        
        ' importar a excel
        Toolbar1.Buttons(10).Enabled = ((Modo = 2) Or (Modo = 0)) And (NroAlbaran = "") And UCase(Dir(App.Path & "\controlclas.cfg")) = UCase("controlclas.cfg")
        Me.mnImportar.Enabled = ((Modo = 2) Or (Modo = 0)) And (NroAlbaran = "") And UCase(Dir(App.Path & "\controlclas.cfg")) = UCase("controlclas.cfg")
         
        ' liquidacion directa
        Toolbar1.Buttons(11).Enabled = b And vParamAplic.LiqDirecta
        Me.mnLiquidacion.Enabled = b And vParamAplic.LiqDirecta
        

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 2) And NroAlbaran = "" 'And Not Check1(0).Value = 1
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        
        If b Then
            Select Case i
              Case 0
                bAux = (b And Me.Data3.Recordset.RecordCount > 0)
              Case 1
                bAux = (b And Me.Adoaux(0).Recordset.RecordCount > 0)
              Case 2
                bAux = (b And Me.Adoaux(1).Recordset.RecordCount > 0)
            End Select
        End If
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i


End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim CadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NroCopias As Integer

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Albarán para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 22 'Impresion de Albaran de clasificacion
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de albaran
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Nº palet
        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numalbar = " & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    End If
    
    CadParam = CadParam & "pDuplicado=1|"
    numParam = numParam + 1
    
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
    With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Albarán de Clasificación"
            '[Monica]12/09/2012:en Mogente necesitan 2 copias de albaran
            If vParamAplic.Cooperativa = 3 Then .NroCopias = 2
            .ConSubInforme = True
            .Show vbModal
    End With

End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
End Sub

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
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
    
'[Monica]03/05/2013: las entradas complementarias desaparecen (marcaban entradas de facturas de siniestro),  estan en rhisfrutasin
'    '[Monica]27/03/2013: nuevo tipo de entrada solo para marcar que se tratan de entradas que vienen
'    '                    a partir de facturas de siniestro (SOLO PARA CATADAU)
'    If vParamAplic.Cooperativa = 0 Then
'        Combo1(0).AddItem "Complementaria"
'        Combo1(0).ItemData(Combo1(0).NewIndex) = 6
'    End If
    
    'recolectado por
    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    'transportado por
    Combo1(2).AddItem "Cooperativa"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Socio"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1

End Sub

Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    
    MenError = "Modificando Lineas clasificacion: "
    b = ModificandoClasificacion(Data1.Recordset.Fields(0), Text1(2).Text, MenError)

    If b Then b = ModificaDesdeFormulario2(Me, 2, "Frame2")


EModificarCab:
    If Err.Number <> 0 Or Not b Then
        MenError = "Modificando Albarán." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
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


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        txtAux3(11).visible = False
        Sql = CadenaInsertarDesdeForm(Me)
        txtAux3(11).visible = True
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
'                BotonAnyadirLinea
                Set frmLHco = New frmManLinHcoFruta
                
                frmLHco.ModoExt = 3
                frmLHco.Albaran = CLng(Text1(0).Text)
                frmLHco.Show vbModal
                
                Set frmLHco = Nothing
                
                CalcularTotales
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Albaranes
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numalbar", "numalbar", Text1(0).Text, "N")
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
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador del Albarán."
    vTipoMov.IncrementarContador (CodTipoMov)
    
EInsertarOferta:
    If Err.Number <> 0 Then
        MenError = "Insertando Albarán." & vbCrLf & "----------------------------" & vbCrLf & MenError
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


Private Sub InsertarLinea(Index As Integer)
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case Index
        Case 0: nomframe = "FrameAux0" 'clasificacion
        Case 1: nomframe = "FrameAux1" 'gastos
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            If NumTabMto = 0 Then
'yo                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtaux(2).Text)
'                End If
            End If
            ' *************************************************
            b = BloqueaRegistro("albaran", "numalbar = " & Data1.Recordset!numalbar)
            Select Case Index
                Case 0  ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid DataGrid3, Adoaux(0), True
                    If b Then BotonAnyadirLinea NumTabMto
'                LLamaLineas NumTabMto, 0
                Case 1
                    CargaGrid DataGrid4, Adoaux(1), True
                    If b Then BotonAnyadirLinea Index
            End Select
'            SSTab1.Tab = NumTabMto
        End If
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
'    NumTabMto = Index
'    If Index = 2 Then NumTabMto = 3
    
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case NumTabMto
        Case 0: vtabla = "rhisfruta_clasif"
        Case 1: vtabla = "rhisfruta_gastos"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case NumTabMto
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid3, Adoaux(0)
    
            anc = DataGrid3.Top
            If DataGrid3.Row < 0 Then
                anc = anc + 220
            Else
                anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
            End If
            
            LLamaLineas ModificaLineas, anc, "DataGrid3"
        
            LimpiarCamposLin "FrameAux0"
            
            txtAux(4).Text = Text1(0).Text 'numalbar
            txtAux(5).Text = Text1(2).Text 'variedad
            Text2(6).Text = ""
            
            BloquearTxt txtAux(6), False
'                    BloquearTxt txtaux(12), False
            PonerFoco txtAux(6)
                    
        Case 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid4, Adoaux(1)
    
            anc = DataGrid4.Top
            If DataGrid4.Row < 0 Then
                anc = anc + 220
            Else
                anc = anc + DataGrid4.RowTop(DataGrid4.Row) + 5
            End If
            
            LLamaLineas ModificaLineas, anc, "DataGrid4"
        
            LimpiarCamposLin "FrameAux1"
            
            txtAux(10).Text = Text1(0).Text 'numalbar
            txtAux(11).Text = NumF ' numlinea
            Text2(1).Text = ""
            
            BloquearTxt txtAux(9), False
'                    BloquearTxt txtaux(12), False
            PonerFoco txtAux(9)
                    
                    
                    
    End Select
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim Cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0:
            nomframe = "FrameAux0" 'clasificacion
        Case 1
            nomframe = "FrameAux1" 'gastos
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
'yo                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtaux(2).Text)
'                End If
            End If
            ' ******************************************************
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModificaLineas = 0

            Select Case NumTabMto
                Case 0

                    V = Adoaux(0).Recordset.Fields(1) 'el 2 es el nº de llinia
                    CargaGrid DataGrid3, Adoaux(0), True

                    ' *** si n'hi han tabs ***
'                    SSTab1.Tab = 1

                    DataGrid3.SetFocus
                    Adoaux(0).Recordset.Find (Adoaux(0).Recordset.Fields(1).Name & " =" & V)

                    LLamaLineas ModificaLineas, 0, "DataGrid3"
                Case 1
                    V = Adoaux(1).Recordset.Fields(1) 'el 2 es el nº de llinia
                    CargaGrid DataGrid4, Adoaux(1), True

                    ' *** si n'hi han tabs ***
'                    SSTab1.Tab = 1

                    DataGrid4.SetFocus
                    Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(1).Name & " =" & V)

                    LLamaLineas ModificaLineas, 0, "DataGrid4"
            End Select
        End If
    End If
        
'    'Cridem al form
'    ' **************** arreglar-ho per a vore lo que es desije ****************
'    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
'    Cad = ""
'    Cad = Cad & ParaGrid(text1(0), 15, "Cód.")
'    Cad = Cad & ParaGrid(text1(2), 60, "Nombre")
'    Cad = Cad & ParaGrid(text1(3), 25, "N.I.F.")
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vtabla = NombreTabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
'        frmB.vTitulo = "Clientes" ' ***** repasa açò: títol de BuscaGrid *****
'        frmB.vSelElem = 1
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha posat valors i tenim que es formulari de búsqueda llavors
'        'tindrem que tancar el form llançant l'event
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha retornat datos, es a decir NO ha retornat datos
'            PonerFoco text1(kCampo)
'        End If
'    End If
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
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

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " numalbar= " & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
'    On Error Resume Next
'
'    If numTab = 0 Or numTab = 1 Or numTab = 2 Or numTab = 3 Then
'        SSTab1.Tab = 2
'    ElseIf numTab = 4 Then
'        SSTab1.Tab = 2
'    End If
'
'    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************

Private Function ActualizaMovimiento(Mens As String) As Boolean
Dim Sql As String
    
    On Error GoTo eActualizaMovimiento
    
    
    Sql = "update smoval set fechamov = " & DBSet(Text1(1).Text, "F") & ", codigope = " & DBSet(Text1(3).Text, "N")
    Sql = Sql & " where tipomovi = 'ALV' and document = " & Data1.Recordset!numalbar
    Sql = Sql & " and codigope = " & Data1.Recordset!CodClien
    Sql = Sql & " and fechamov = " & DBSet(Data1.Recordset!FechaAlb, "F")
    
    conn.Execute Sql
    
eActualizaMovimiento:
    If Err.Number <> 0 Then
        Mens = Mens & vbCrLf & Err.Description
        ActualizaMovimiento = False
    Else
        ActualizaMovimiento = True
    End If
End Function



Private Sub CalcularTotales()
Dim cajas As Long
Dim KilosNet As Long
Dim KilosBru As Long
Dim KilosTra As Long
' gastos
Dim Transporte As Currency
Dim Acarreo As Currency
Dim Recolec As Currency
Dim Penaliza As Currency
Dim PrEstimado As Currency

Dim Rs As ADODB.Recordset
Dim Sql As String

    cajas = TotalRegistros("select sum(numcajon) from rhisfruta_entradas where numalbar = " & Data1.Recordset.Fields(0))
    KilosNet = TotalRegistros("select sum(kilosnet) from rhisfruta_entradas where numalbar = " & Data1.Recordset.Fields(0))
    KilosBru = TotalRegistros("select sum(kilosbru) from rhisfruta_entradas where numalbar = " & Data1.Recordset.Fields(0))
    KilosTra = TotalRegistros("select sum(kilostra) from rhisfruta_entradas where numalbar = " & Data1.Recordset.Fields(0))
    
    'gastos y precio estimado
    Sql = "select sum(imptrans), sum(impacarr), sum(imprecol), sum(imppenal), sum(prestimado),count(*) from rhisfruta_entradas where numalbar = " & Data1.Recordset.Fields(0)
        
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Transporte = 0
    Acarreo = 0
    Recolec = 0
    Penaliza = 0
    PrEstimado = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Transporte = DBLet(Rs.Fields(0).Value, "N")
        If Rs.Fields(1).Value <> 0 Then Acarreo = DBLet(Rs.Fields(1).Value, "N")
        If Rs.Fields(2).Value <> 0 Then Recolec = DBLet(Rs.Fields(2).Value, "N")
        If Rs.Fields(3).Value <> 0 Then Penaliza = DBLet(Rs.Fields(3).Value, "N")
        ' el precio estimado es el precio medio de las lineas sum(precios)/count(lineas)
        If Rs.Fields(4).Value <> 0 And Rs.Fields(5).Value <> 0 Then
            PrEstimado = Round2(DBLet(Rs.Fields(4).Value, "N") / DBLet(Rs.Fields(5).Value, "N"), 4)
        End If
    End If
    
    Set Rs = Nothing
    
    BotonModificar
    
    Text1(5).Text = Format(KilosBru, "###,###,##0")
    Text1(7).Text = Format(KilosNet, "###,###,##0")
    Text1(6).Text = Format(cajas, "###,###,##0")
    Text1(13).Text = Format(KilosTra, "###,###,##0")
    
    'gastos y precio estimado
    Text1(8).Text = ""
    Text1(9).Text = ""
    Text1(10).Text = ""
    Text1(11).Text = ""
    Text1(12).Text = ""
    
    
    If Transporte <> 0 Then Text1(8).Text = Format(Transporte, "###,###,##0.00")
    If Acarreo <> 0 Then Text1(9).Text = Format(Acarreo, "###,###,##0.00")
    If Recolec <> 0 Then Text1(10).Text = Format(Recolec, "###,###,##0.00")
    If Penaliza <> 0 Then Text1(11).Text = Format(Penaliza, "###,###,##0.00")
    If PrEstimado <> 0 Then Text1(12).Text = Format(PrEstimado, "###,##0.0000")

    cmdAceptar_Click
End Sub

Private Sub BotonGenerarFactura(Albaran As String)
'Dim SQL As String
'Dim fecfactu As String
'Dim vFacturaVta As CFacturaVta
'Dim b As Boolean
'Dim Observaciones As String
'
'    Observaciones = DevuelveDesdeBDNew(cAgro, "clientes", "observac", "codclien", Data1.Recordset!CodClien, "N")
'    If Observaciones <> "" Then
'        MsgBox Observaciones, vbInformation, "Observaciones del cliente"
'    End If
'
'    ' comprobamos si hay lineas con precio provisional = 0
'    SQL = "select count(*) from albaran_variedad where numalbar = " & DBSet(Albaran, "N")
'    SQL = SQL & " and (preciopro is null or preciopro = 0)"
'    If TotalRegistros(SQL) <> 0 Then
'        If MsgBox("Hay lineas de albaran sin precio provisional. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'            Exit Sub
'        End If
'    End If
'
'    fecfactu = InputBox("Fecha Factura:", "Fecha de Factura", Format(Now, "dd/mm/yyyy"))
'    If EsFechaOK(fecfactu) Then
'        Set vFacturaVta = New CFacturaVta
'        b = vFacturaVta.PasarAlbaranAFactura("albaran.numalbar=" & Albaran, fecfactu)
'        If b Then
'            Data3.Refresh
'            MsgBox "Proceso realizado correctamente.", vbExclamation
'        End If
'    Else
'        MsgBox "Fecha de Factura incorrecta.", vbExclamation
'    End If
End Sub

Private Function ModificarFechaMovimiento(Albaran As Long, Fechamov As String) As Boolean
Dim Sql As String
Dim b As Boolean
    
    On Error GoTo eModificarFechaMovimiento
        
    ModificarFechaMovimiento = False
    
    Sql = "update albaran_envase set fechamov = " & DBSet(Fechamov, "F")
    Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
    
    conn.Execute Sql
    
    ModificarFechaMovimiento = True
    Exit Function
    
eModificarFechaMovimiento:
    If Err.Number <> 0 Then
        ModificarFechaMovimiento = False
    End If
End Function

Private Sub ComprobarClasificacion()
Dim Sql As String

    Sql = "select sum(kilosnet) from rhisfruta_clasif where numalbar = " & Data1.Recordset.Fields(0)
        
    If TotalRegistros(Sql) <> Data1.Recordset!KilosNet Then
        MsgBox "Los kilos netos de la clasificación no se corresponden con el total. Revise.", vbExclamation
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
    
    Text1(4).Text = ""
    Text2(4).Text = ""
    Text2(0).Text = ""
    Text2(7).Text = ""
    Text3(3).Text = ""
    Text4(3).Text = ""
    Text5(3).Text = ""
    If Not Rs.EOF Then
        Text1(4).Text = campo
        PonerFormatoEntero Text1(5)
        Text2(4).Text = DBLet(Rs.Fields(0).Value, "N") ' codigo de partida
        If Text2(4).Text <> "" Then Text2(4).Text = Format(Text2(4).Text, "0000")
        '[Monica]04/10/2010: añadido el nro de orden
        Text2(7).Text = DBLet(Rs.Fields(5).Value, "N") ' nro de orden
        If Text2(7).Text <> "" Then Text2(7).Text = Format(Text2(7).Text, "0000")
        
        Text2(0).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text3(3).Text = DBLet(Rs.Fields(2).Value, "N") ' codigo de zona
        If Text3(3).Text <> "" Then Text3(3).Text = Format(Text3(3).Text, "0000")
        Text4(3).Text = DBLet(Rs.Fields(3).Value, "T") ' nombre de zona
        Text5(3).Text = DBLet(Rs.Fields(4).Value, "T") ' descripcion de poblacion
    End If
    
    Set Rs = Nothing
    
End Sub



Private Sub VisualizarDatosCampo(campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub
                                                    '[Monica]14/06/2010:puede que el campo lo hayan dado de baja
                                                    'posteriormente a las entradas de fruta
    Cad = "rcampos.codcampo = " & DBSet(campo, "N") '& " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla, rcampos.nrocampo from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text1(4).Text = ""
    Text2(4).Text = ""
    Text2(0).Text = ""
    Text2(7).Text = ""
    Text3(3).Text = ""
    Text4(3).Text = ""
    Text5(3).Text = ""
    If Not Rs.EOF Then
        Text1(4).Text = campo
        PonerFormatoEntero Text1(5)
        Text2(4).Text = DBLet(Rs.Fields(0).Value, "N") ' codigo de partida
        If Text2(4).Text <> "" Then Text2(4).Text = Format(Text2(4).Text, "0000")
        Text2(0).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        '[Monica]04/10/2010: añadido el nro de campo
        Text2(7).Text = DBSet(Rs.Fields(5).Value, "N") ' nro de campo
        If Text2(7).Text <> "" Then Text2(7).Text = Format(Text2(7).Text, "0000")
        
        Text3(3).Text = DBLet(Rs.Fields(2).Value, "N") ' codigo de zona
        If Text3(3).Text <> "" Then Text3(3).Text = Format(Text3(3).Text, "0000")
        Text4(3).Text = DBLet(Rs.Fields(3).Value, "T") ' nombre de zona
        Text5(3).Text = DBLet(Rs.Fields(4).Value, "T") ' descripcion de poblacion
    End If
    
    Set Rs = Nothing
    
End Sub


Private Sub PonerCamposSocioVariedad()
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If Text1(2).Text = "" Or Text1(3).Text = "" Then Exit Sub
    
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codsocio = " & DBSet(Text1(3).Text, "N") & " and rcampos.fecbajas is null"
    Cad = Cad & " and rcampos.codvarie = " & DBSet(Text1(2).Text, "N")
     
    Cad1 = "select count(*) from rcampos where " & Cad
     
    NumRegis = TotalRegistros(Cad1)
    
    If NumRegis = 0 Then Exit Sub
    If NumRegis = 1 Then
        Cad1 = "select codcampo from rcampos where " & Cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text1(4).Text = DBLet(Rs.Fields(0).Value)
            PonerDatosCampo Text1(4).Text
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadWHERE = " and " & Cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens.campo = Text1(4).Text
        frmMens.OpcionMensaje = 6
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
    
End Sub

Private Sub CalcularTotalGastos()
Dim Gastos As Double

    If Data1.Recordset.EOF Then Exit Sub

    Gastos = DevuelveValor("select sum(importe) from rhisfruta_gastos where numalbar = " & Data1.Recordset.Fields(0))
    Text2(5).Text = Format(Gastos, "###,###,##0.00")
    
End Sub



Private Function ModificandoClasificacion(numalbar As String, Variedad As String, Mens As String) As Boolean
Dim Sql As String

    On Error GoTo eModificandoClasificacion

    ModificandoClasificacion = False

    Sql = "update rhisfruta_clasif set codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " where numalbar = " & DBSet(numalbar, "N")
    
    conn.Execute Sql

    ModificandoClasificacion = True
    Exit Function
    
eModificandoClasificacion:
    Mens = Mens & vbCrLf & Err.Description
End Function


Private Sub ImportacionEntradas()
Dim Sql As String
Dim i As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String

On Error GoTo eError

'    If Not DatosOk Then Exit Sub
    
    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist

    Me.CommonDialog1.DefaultExt = "csv"
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "albaranes.csv"
    
    
    Me.CommonDialog1.CancelError = True
    
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

        If CargaInicialTabla(Me.CommonDialog1.FileName) Then
            SociosNoExisten = ""
            VariedadesNoExisten = ""
            CalidadesNoExisten = ""
            If CompruebaSociosVariedades(SociosNoExisten, VariedadesNoExisten, CalidadesNoExisten) Then
                If SociosNoExisten <> "" Then
                    MsgBox "Los siguientes socios no existen, creelos y vuelva a importar: " & vbCrLf & vbCrLf & Mid(SociosNoExisten, 1, Len(SociosNoExisten) - 2), vbExclamation
                    Exit Sub
                End If
                If VariedadesNoExisten <> "" Then
                    MsgBox "Las siguientes variedades no existen, creelas y vuelva a importar: : " & vbCrLf & vbCrLf & Mid(VariedadesNoExisten, 1, Len(VariedadesNoExisten) - 2), vbExclamation
                    Exit Sub
                End If
                If CalidadesNoExisten <> "" Then
                    MsgBox "Las siguientes variedades-calidades no existen, creelas y vuelva a importar: : " & vbCrLf & vbCrLf & Mid(VariedadesNoExisten, 1, Len(CalidadesNoExisten) - 2), vbExclamation
                    Exit Sub
                End If
            End If
        End If

        If ProcesarFicheroEntradas(Me.CommonDialog1.FileName) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            Sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
            
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Han habido errores en el Traspaso de Entradas. ", vbExclamation
                cadTitulo = "Errores en el Traspaso de Entradas"
                cadNombreRPT = "rErroresTrasEntBascula.rpt"
                
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                
                LlamarImprimir
            End If
        Else
            MsgBox "No se ha podido realizar el proceso.", vbExclamation
        End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
eError:
    If Err.Number = 32755 Then Exit Sub ' le han dado a cancelar

End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub


Private Function CargaInicialTabla(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim Sql1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String

    On Error GoTo eProcesarFicheroABN


    CargaInicialTabla = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    lblProgres(0).Caption = "Carga inicial fichero: " & nomFich
    longitud = FileLen(nomFich)
    
        
    ' salto la primera linea que es la cabecera
    Line Input #NF, Cad
    Me.Refresh
    i = 1
    
        
    b = True
    While Not EOF(NF) And b
        i = i + 1
        
        lblProgres(0).Caption = "Carga inicial fichero. Linea " & i
        Me.Refresh
        
        Cad = Cad & ";"
        If Mid(Cad, 1, 6) <> ";;;;;;" Then b = InsertarLineaPrevia(Cad)
        
        If b Then
            If i > 20 Then
                CargaInicialTabla = True
                Close #NF
                lblProgres(0).Caption = ""
                Exit Function
            End If
        End If
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" And b Then
        Cad = Cad & ";"
        If Mid(Cad, 1, 6) <> ";;;;;;" Then b = InsertarLineaPrevia(Cad)
    End If
    
    CargaInicialTabla = b
    
    lblProgres(0).Caption = ""
    
eProcesarFicheroABN:
    If Err.Number <> 0 Or Not b Then
    Else
    End If
 

End Function

Private Function InsertarLineaPrevia(Cad As String) As Boolean
Dim Sql As String
Dim cadena As String

    On Error GoTo eInsertarLineaPrevia

    InsertarLineaPrevia = True
    
    CargarVariables Cad
    
    ' insertamos la entrada
    cadena = vUsu.Codigo & "," & NumNota & "," & DBSet(FechaEnt, "F") & "," & DBSet(HoraEnt, "T") & "," & DBSet(Socio, "N") & "," & DBSet(Variedad, "N") & ","
    cadena = cadena & DBSet(Bruto, "N") & "," & DBSet(Neto, "N") & "," & DBSet(NomSocio, "T") & "," & DBSet(NIF, "T")
    
    Sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1, importe2, importe3, importe4, importe5, nombre2, nombre3) values "
    Sql = Sql & "(" & cadena & ")"
    conn.Execute Sql
    
    Exit Function
    
eInsertarLineaPrevia:
    InsertarLineaPrevia = False
    MuestraError Err.Number, "Insertar Linea Previa", Err.Description
End Function


Private Sub CargarVariables(Cad As String)
            
    NumNota = ""
    FechaEnt = ""
    HoraEnt = ""
    Bruto = ""
    Variedad = ""
    Socio = ""
    Neto = ""
    NIF = ""
    NomSocio = ""
    
    NumNota = RecuperaValorNew(Cad, ";", 1)
    FechaEnt = RecuperaValorNew(Cad, ";", 2)
    HoraEnt = RecuperaValorNew(Cad, ";", 3)
    Bruto = RecuperaValorNew(Cad, ";", 8)
    Variedad = RecuperaValorNew(Cad, ";", 7)
    Socio = RecuperaValorNew(Cad, ";", 4)
    Neto = RecuperaValorNew(Cad, ";", 9)
    NIF = RecuperaValorNew(Cad, ";", 6)
    NomSocio = RecuperaValorNew(Cad, ";", 5)
    
End Sub

Private Function CompruebaSociosVariedades(ByRef SociosNoExisten As String, ByRef VariedadesNoExisten As String, ByRef CalidadesNoExisten As String) As Boolean
Dim NumLin As String
Dim b As Boolean
Dim Sql As String

Dim Sql1 As String

Dim Mens As String
Dim numlinea As Long

Dim vError As Boolean
Dim vNota As Long
Dim cadena As String
Dim vNif As String

Dim Rs As ADODB.Recordset

    CompruebaSociosVariedades = True
    
    SociosNoExisten = ""
    VariedadesNoExisten = ""
    CalidadesNoExisten = ""
    
    Sql = " select importe2, importe3, nombre3 from tmpinformes where codusu = " & vUsu.Codigo
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Socio = DBSet(Rs!importe2, "N")
        Variedad = DBSet(Rs!importe3, "N")
        vNif = DBLet(Rs!nombre3)
        
        'Comprobamos que el socio existe
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "rsocios", "codsocio", "codsocio", Socio, "N", , "nifsocio", vNif, "T")
        
        If Sql = "" Then SociosNoExisten = SociosNoExisten & Socio & ", "
        
        'Comprobamos que la variedad existe
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "variedades", "codvarie", "codvarie", Variedad, "N")
        If Sql = "" Then
            VariedadesNoExisten = VariedadesNoExisten & Variedad & ", "
        Else
            Sql = "select min(codcalid) from rcalidad where codvarie = " & DBSet(Variedad, "N")
            Sql = DevuelveValor(Sql)
            If Sql = "0" Then CalidadesNoExisten = CalidadesNoExisten & Variedad & ","
        End If
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
End Function

Private Function ProcesarFicheroEntradas(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim Sql1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String

    On Error GoTo eProcesarFicheroEntradas


    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    conn.BeginTrans

    ProcesarFicheroEntradas = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    ' salto la primera linea que es la cabecera
    Line Input #NF, Cad
    Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
    Me.Refresh
    i = 1
    
        
    b = True
    While Not EOF(NF) And b
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(0).Caption = "Procesando Fichero. Linea " & i
            Me.Refresh
        
        Cad = Cad & ";"
        If Mid(Cad, 1, 6) <> ";;;;;;" Then b = InsertarLineaTraspasoEntradas(Cad)
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" And b Then
        Cad = Cad & ";"
        If Mid(Cad, 1, 6) <> ";;;;;;" Then b = InsertarLineaTraspasoEntradas(Cad)
    End If
    
    ProcesarFicheroEntradas = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    
eProcesarFicheroEntradas:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
    Else
        conn.CommitTrans
    End If
End Function


Private Function InsertarLineaTraspasoEntradas(Cad As String) As Boolean
Dim NumLin As String
Dim b As Boolean
Dim Sql As String

Dim Sql1 As String

Dim Mens As String
Dim numlinea As Long

Dim vError As Boolean
Dim vNota As Long
Dim cadena As String
Dim Codcampo As Long
Dim HayError As Boolean
Dim Calidad As String

    On Error GoTo EInsertarLinea

    InsertarLineaTraspasoEntradas = False
    
    CargarVariables Cad
    
    
     ' comprobaciones para poder insertar la entrada
    Sql = "select codcampo from rcampos where codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and codsocio = " & DBSet(Socio, "N")
    
    Codcampo = DevuelveValor(Sql)
    
    If CLng(Codcampo) = 0 Then
        Set frmMens3 = New frmMensajes
        frmMens3.cadena = Socio & "|" & Variedad & "||||"
        frmMens3.OpcionMensaje = 62
        frmMens3.Show vbModal
        Set frmMens3 = Nothing
        
        If Not Continuar Then Exit Function
    
        Sql = "select codcampo from rcampos where codvarie = " & DBSet(Variedad, "N")
        Sql = Sql & " and codsocio = " & DBSet(Socio, "N")
        
        Codcampo = DevuelveValor(Sql)
    End If
    
    
    ' al nro de nota le sumo por delante la cooperativa
    vNota = 2000000 + NumNota
    
    ' Comprobamos que la entrada no exista ya
    Sql = "select count(*) from rhisfruta where numalbar = " & DBSet(vNota, "N")
    If TotalRegistros(Sql) <> 0 Then
        HayError = True
    End If
    
    If HayError Then
'        Sql = "update rhisfruta set fecalbar = " & DBSet(FechaEnt, "F")
'        Sql = Sql & ", codvarie = " & DBSet(Variedad, "N")
'        Sql = Sql & ", codsocio = " & DBSet(Socio, "N")
'        Sql = Sql & ", codcampo = " & DBSet(campo, "N")
'        Sql = Sql & ", kilosbru = " & DBSet(Bruto, "N")
'        Sql = Sql & ", kilosnet = " & DBSet(Neto, "N")
'        Sql = Sql & " where numalbar = " & DBSet(vNota, "N")
'        conn.Execute Sql
'
        Exit Function
    End If
    
    ' insertamos en la tabla de rhisfruta
    Sql = "insert into rhisfruta ("
    Sql = Sql & "`numalbar`,`fecalbar`,`codvarie`,`codsocio`,`codcampo`,`tipoentr`,"
    Sql = Sql & "`recolect`,`kilosbru`,`numcajon`,`kilosnet`,`imptrans`,`impacarr`,"
    Sql = Sql & "`imprecol`,`imppenal`,`impreso`,`impentrada`,`cobradosn`,`prestimado`,"
    Sql = Sql & "`nromuestraalmz` ) VALUES ("
    Sql = Sql & DBSet(vNota, "N") & ","
    Sql = Sql & DBSet(FechaEnt, "F") & ","
    Sql = Sql & DBSet(Variedad, "N") & ","
    Sql = Sql & DBSet(Socio, "N") & ","
    
    'campo
    Sql = Sql & DBSet(Codcampo, "N") & ","
    
    Sql = Sql & "0,0,"
    Sql = Sql & DBSet(Bruto, "N") & ","
    Sql = Sql & "0," ' numero de cajones
    Sql = Sql & DBSet(Neto, "N") & ","
    Sql = Sql & "0,0,0,0,0,0,0,0,"
    Sql = Sql & ValorNulo & ")"
    
    conn.Execute Sql
    
    
    ' insertamos en la tabla rhisfruta_entradas
    Sql = "insert into rhisfruta_entradas ("
    Sql = Sql & "numalbar,numnotac,fechaent,horaentr,kilosbru,numcajon,kilosnet,kilostra,tiporecol) "
    Sql = Sql & " VALUES ("
    Sql = Sql & DBSet(vNota, "N") & ","
    Sql = Sql & DBSet(vNota, "N") & ","
    Sql = Sql & DBSet(FechaEnt, "F") & ","
    Sql = Sql & DBSet(FechaEnt & " " & HoraEnt, "FH") & ","
    Sql = Sql & DBSet(Bruto, "N") & ","
    Sql = Sql & "0,"
    Sql = Sql & DBSet(Neto, "N") & ","
    Sql = Sql & DBSet(Neto, "N") & ",0)"
    
    conn.Execute Sql
    
    Calidad = DevuelveValor("select min(codcalid) from rcalidad where codvarie = " & DBSet(Variedad, "N"))
    
    ' insertamos en la tabla rhisfruta_clasif
    Sql = "insert into rhisfruta_clasif (numalbar,codvarie,codcalid,kilosnet) values ("
    Sql = Sql & DBSet(vNota, "N") & ","
    Sql = Sql & DBSet(Variedad, "N") & ","
    Sql = Sql & DBSet(Calidad, "N") & ","
    Sql = Sql & DBSet(Neto, "N") & ")"
    
    conn.Execute Sql
    
    InsertarLineaTraspasoEntradas = True
    Exit Function
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLineaTraspasoEntradas = False
        MsgBox "Error en Insertar Línea Traspaso Entradas" & Err.Description, vbExclamation
    End If
End Function

