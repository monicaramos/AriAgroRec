VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTercHcoFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas Terceros"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   11745
   Icon            =   "frmTercHcoFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTercHcoFact.frx":000C
   ScaleHeight     =   6840
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   45
      TabIndex        =   52
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1260
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
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
      TabPicture(0)   =   "frmTercHcoFact.frx":0A0E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(11)"
      Tab(0).Control(1)=   "Text1(15)"
      Tab(0).Control(2)=   "Frame2(1)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmTercHcoFact.frx":0A2A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DataGrid2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtAux3(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtAux3(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text3(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtAux3(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtAux3(3)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtAux3(4)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtAux3(5)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtAux3(6)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtAux3(7)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "CmdTraerAlbaranes"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "FrameAnticipos"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdAnticipos"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "FrameObserva"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      Begin VB.Frame FrameObserva 
         Caption         =   "Observaciones"
         ForeColor       =   &H00972E0B&
         Height          =   2055
         Left            =   225
         TabIndex        =   54
         Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
         Top             =   2610
         Width           =   9945
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   8
            Left            =   720
            MaxLength       =   80
            TabIndex        =   45
            Tag             =   "Observación 5|T|S|||rlifter|observa5||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1560
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   7
            Left            =   720
            MaxLength       =   80
            TabIndex        =   46
            Tag             =   "Observación 4|T|S|||rlifter|observa4||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1230
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   6
            Left            =   720
            MaxLength       =   80
            TabIndex        =   44
            Tag             =   "Observación 3|T|S|||rlifter|observa3||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   900
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   5
            Left            =   720
            MaxLength       =   80
            TabIndex        =   43
            Tag             =   "Observación 2|T|S|||rlifter|observa2||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   570
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   4
            Left            =   720
            MaxLength       =   80
            TabIndex        =   42
            Tag             =   "Observación 1|T|S|||rlifter|observa1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   240
            Width           =   8940
         End
      End
      Begin VB.PictureBox cmdAnticipos 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   10320
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   114
         Top             =   540
         Width           =   375
      End
      Begin VB.Frame FrameAnticipos 
         Caption         =   "Anticipos"
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
         Height          =   2085
         Left            =   210
         TabIndex        =   106
         Top             =   2580
         Visible         =   0   'False
         Width           =   9975
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   1965
            MaxLength       =   7
            TabIndex        =   112
            Tag             =   "Socio Anti|N|N|||rliantifter|codsocioanti||S|"
            Text            =   "socanti"
            Top             =   1140
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   2445
            MaxLength       =   7
            TabIndex        =   111
            Tag             =   "Num.Fact Anti|T|N|||rliantifter|numfactuanti||S|"
            Text            =   "numfact"
            Top             =   1140
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   5
            Left            =   3105
            MaxLength       =   4
            TabIndex        =   110
            Tag             =   "Fec.Factura Anti|F|N|||rliantifter|fecfactuanti|dd/mm/yyyy|S|"
            Text            =   "fecf"
            Top             =   1140
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   495
            MaxLength       =   7
            TabIndex        =   109
            Tag             =   "Socio|N|N|||rliantifter|codsocio||S|"
            Text            =   "codsoci"
            Top             =   1140
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   945
            MaxLength       =   7
            TabIndex        =   108
            Tag             =   "Num.Fact|T|N|||rliantifter|numfactu||S|"
            Text            =   "numfact"
            Top             =   1140
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   1515
            MaxLength       =   4
            TabIndex        =   107
            Tag             =   "Fec.Factura|F|N|||rliantifter|fecfactu|dd/mm/yyyy|S|"
            Text            =   "fecf"
            Top             =   1140
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "frmTercHcoFact.frx":0A46
            Height          =   1395
            Left            =   150
            TabIndex        =   113
            Top             =   270
            Width           =   3015
            _ExtentX        =   5318
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
      Begin VB.CommandButton CmdTraerAlbaranes 
         Height          =   375
         Left            =   10860
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   540
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   8100
         MaxLength       =   12
         TabIndex        =   41
         Tag             =   "Precio Estimado|N|S|||rlifter|prestimado|###,##0.0000|N|"
         Text            =   "precio estim"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   7080
         MaxLength       =   13
         TabIndex        =   40
         Tag             =   "Importe|N|N|0||rlifter|importel|#,###,###,##0.00|N|"
         Text            =   "importe"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   6090
         MaxLength       =   18
         TabIndex        =   39
         Tag             =   "Precio|N|N|||rlifter|precio|#,###,###,##0.0000|N|"
         Text            =   "precio"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   5100
         MaxLength       =   13
         TabIndex        =   38
         Tag             =   "Kilos Netos|N|N|||rlifter|kilosnet|#,###,###,##0|N|"
         Text            =   "kilosnet"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   4110
         MaxLength       =   30
         TabIndex        =   37
         Text            =   "nomvarie"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   3060
         MaxLength       =   6
         TabIndex        =   36
         Tag             =   "Variedad|N|N|||rlifter|codvarie|000000|N|"
         Text            =   "Variedad"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Height          =   4560
         Index           =   1
         Left            =   -74865
         TabIndex        =   73
         Top             =   315
         Width           =   11175
         Begin VB.Frame FrameCliente 
            Caption         =   "Datos de Socio"
            ForeColor       =   &H00972E0B&
            Height          =   1740
            Left            =   45
            TabIndex        =   88
            Top             =   135
            Width           =   11055
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   36
               Left            =   6960
               MaxLength       =   15
               TabIndex        =   118
               Tag             =   "Imp.Cargo|N|S|||rcafter|impcargo|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1350
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   35
               Left            =   6960
               MaxLength       =   3
               TabIndex        =   116
               Tag             =   "Concepto Cargo|N|S|0|999|rcafter|concepcargo|000|N|"
               Text            =   "Text1"
               Top             =   990
               Width           =   540
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H80000018&
               Height          =   315
               Index           =   35
               Left            =   7530
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   115
               Text            =   "Text2"
               Top             =   990
               Width           =   3285
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   6
               Left            =   1125
               MaxLength       =   35
               TabIndex        =   9
               Tag             =   "Domicilio|T|N|||rcafter|domsocio||N|"
               Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
               Top             =   645
               Width           =   4030
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H80000018&
               Height          =   315
               Index           =   10
               Left            =   7530
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   89
               Text            =   "Text2"
               Top             =   645
               Width           =   3285
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   10
               Left            =   6945
               MaxLength       =   3
               TabIndex        =   13
               Tag             =   "Forma de Pago|N|N|0|999|rcafter|codforpa|000|N|"
               Text            =   "Text1"
               Top             =   645
               Width           =   540
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   4
               Left            =   1125
               MaxLength       =   15
               TabIndex        =   7
               Tag             =   "NIF Tercero|T|N|||rcafter|nifsocio||N|"
               Text            =   "123456789"
               Top             =   285
               Width           =   1110
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   5
               Left            =   3195
               MaxLength       =   20
               TabIndex        =   8
               Tag             =   "teléfono tercero|T|S|||rcafter|telsocio||N|"
               Text            =   "12345678911234567899"
               Top             =   285
               Width           =   1965
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   8
               Left            =   1755
               MaxLength       =   30
               TabIndex        =   11
               Tag             =   "Población|T|N|||rcafter|pobsocio||N|"
               Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
               Top             =   990
               Width           =   3405
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   7
               Left            =   1140
               MaxLength       =   6
               TabIndex        =   10
               Tag             =   "CPostal|T|N|||rcafter|codpobla||N|"
               Text            =   "Text15"
               Top             =   990
               Width           =   630
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   9
               Left            =   1125
               MaxLength       =   30
               TabIndex        =   12
               Tag             =   "Provincia|T|N|||rcafter|prosocio||N|"
               Text            =   "Text1 Text1 Text1 Text1 Text22"
               Top             =   1350
               Width           =   2445
            End
            Begin VB.Label Label1 
               Caption         =   "Importe"
               Height          =   255
               Index           =   18
               Left            =   5730
               TabIndex        =   119
               Top             =   1350
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Cod.Cargo"
               Height          =   255
               Index           =   13
               Left            =   5730
               TabIndex        =   117
               Top             =   990
               Width           =   855
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   7
               Left            =   6660
               ToolTipText     =   "Buscar concepto cargo"
               Top             =   1020
               Width           =   240
            End
            Begin VB.Label Label1 
               Caption         =   "Domicilio"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   95
               Top             =   645
               Width           =   735
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   3
               Left            =   6660
               ToolTipText     =   "Buscar forma de pago"
               Top             =   675
               Width           =   240
            End
            Begin VB.Label Label1 
               Caption         =   "Forma Pago"
               Height          =   255
               Index           =   15
               Left            =   5730
               TabIndex        =   94
               Top             =   645
               Width           =   855
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   1
               Left            =   855
               ToolTipText     =   "Buscar proveedor varios"
               Top             =   300
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label Label1 
               Caption         =   "NIF"
               Height          =   255
               Index           =   20
               Left            =   120
               TabIndex        =   93
               Top             =   285
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Teléfono"
               Height          =   255
               Index           =   19
               Left            =   2445
               TabIndex        =   92
               Top             =   285
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Población"
               Height          =   255
               Index           =   16
               Left            =   120
               TabIndex        =   91
               Top             =   990
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Provincia"
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   90
               Top             =   1350
               Width           =   735
            End
         End
         Begin VB.Frame FrameFactura 
            Height          =   2670
            Left            =   45
            TabIndex        =   74
            Top             =   1845
            Width           =   11055
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   34
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   30
               Tag             =   "Importe Retencion|N|S|||rcafter|trefacpr|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   2295
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   33
               Left            =   5040
               MaxLength       =   5
               TabIndex        =   28
               Tag             =   "% Ret|N|S|0|99.90|rcafter|retfacpr|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   2295
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   32
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   29
               Tag             =   "Base Retencion|N|S|||rcafter|basereten|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   2295
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   14
               Left            =   2040
               MaxLength       =   15
               TabIndex        =   14
               Tag             =   "Imp.Bruto|N|N|||rcafter|brutofac|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   435
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   16
               Left            =   3960
               MaxLength       =   15
               TabIndex        =   15
               Tag             =   "Imp. Dto Gn|N|N|||rcafter|impgnral|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   450
               Width           =   1365
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   17
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   75
               Text            =   "Text1 7"
               Top             =   435
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   24
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   18
               Tag             =   "Base Imponible 1|N|N|||rcafter|baseiva1|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1080
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   18
               Left            =   4320
               MaxLength       =   3
               TabIndex        =   16
               Tag             =   "Cod. IVA 1|N|S|0|999|rcafter|tipoiva1|000|N|"
               Text            =   "Text1 7"
               Top             =   1080
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   21
               Left            =   5040
               MaxLength       =   5
               TabIndex        =   17
               Tag             =   "% IVA 1|N|S|0|99.90|rcafter|porciva1|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1080
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   27
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   19
               Tag             =   "Importe IVA 1|N|N|||rcafter|impoiva1|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1080
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   25
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   22
               Tag             =   "Base Imponible 2 |N|S|||rcafter|baseiva2|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1395
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   19
               Left            =   4320
               MaxLength       =   3
               TabIndex        =   20
               Tag             =   "Cod. IVA 2|N|S|0|999|rcafter|tipoiva2|000|N|"
               Text            =   "Text1 7"
               Top             =   1395
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   22
               Left            =   5040
               MaxLength       =   5
               TabIndex        =   21
               Tag             =   "& IVA 2|N|S|0|99.90|rcafter|porciva2|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1395
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   28
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   23
               Tag             =   "Importe IVA 2|N|S|||rcafter|impoiva2|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1395
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   26
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   26
               Tag             =   "Base Imponible 3|N|S|||rcafter|baseiva3|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1725
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   20
               Left            =   4320
               MaxLength       =   3
               TabIndex        =   24
               Tag             =   "Cod. IVA 3|N|S|0|999|rcafter|tipoiva3|000|N|"
               Text            =   "Text1 7"
               Top             =   1725
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   23
               Left            =   5040
               MaxLength       =   5
               TabIndex        =   25
               Tag             =   "% IVA 3|N|S|0|99.90|rcafter|porciva3|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1725
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   29
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   27
               Tag             =   "Importe IVA 3|N|S|||rcafter|impoiva3|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1725
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   30
               Left            =   9360
               MaxLength       =   15
               TabIndex        =   31
               Tag             =   "Total Factura|N|N|||rcafter|totalfac|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   2310
               Width           =   1605
            End
            Begin VB.Label Label1 
               Caption         =   "Imp. Retención"
               Height          =   255
               Index           =   8
               Left            =   7560
               TabIndex        =   98
               Top             =   2070
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   7335
               TabIndex        =   97
               Top             =   2160
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "% RET"
               Height          =   255
               Index           =   4
               Left            =   5040
               TabIndex        =   96
               Top             =   2070
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Imp. Bruto"
               Height          =   255
               Index           =   10
               Left            =   2040
               TabIndex        =   87
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Imp. Descuento"
               Height          =   255
               Index           =   12
               Left            =   4080
               TabIndex        =   86
               Top             =   240
               Width           =   1170
            End
            Begin VB.Label Label1 
               Caption         =   "Base Imponible"
               Height          =   255
               Index           =   14
               Left            =   5880
               TabIndex        =   85
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   31
               Left            =   3720
               TabIndex        =   84
               Top             =   360
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   32
               Left            =   5520
               TabIndex        =   83
               Top             =   360
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "Imp. IVA"
               Height          =   255
               Index           =   33
               Left            =   7605
               TabIndex        =   82
               Top             =   855
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   37
               Left            =   7320
               TabIndex        =   81
               Top             =   960
               Width           =   135
            End
            Begin VB.Line Line1 
               X1              =   4320
               X2              =   7320
               Y1              =   825
               Y2              =   825
            End
            Begin VB.Label Label1 
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   36
               Left            =   11880
               TabIndex        =   80
               Top             =   2160
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   38
               Left            =   9120
               TabIndex        =   79
               Top             =   2310
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "TOTAL FACTURA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Index           =   39
               Left            =   9330
               TabIndex        =   78
               Top             =   2070
               Width           =   1530
            End
            Begin VB.Label Label1 
               Caption         =   "% IVA"
               Height          =   255
               Index           =   41
               Left            =   5040
               TabIndex        =   77
               Top             =   870
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Cod. IVA"
               Height          =   255
               Index           =   42
               Left            =   4320
               TabIndex        =   76
               Top             =   840
               Width           =   735
            End
         End
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   9840
         MaxLength       =   7
         TabIndex        =   71
         Tag             =   "Importe|N|S|||tcafpa|importe|###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2025
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   35
         Tag             =   "Fecha Albaran|F|N|||rlifter|fechaalb|dd/mm/yyyy|N|"
         Text            =   "fecalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   34
         Tag             =   "Nº Albaran|N|N|||rlifter|numalbar|000000|N|"
         Text            =   "numalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmTercHcoFact.frx":0A5B
         Height          =   1935
         Left            =   240
         TabIndex        =   53
         Top             =   525
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   3413
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
         Caption         =   "Albaranes de la Factura"
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
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   -72750
         MaxLength       =   15
         TabIndex        =   103
         Tag             =   "Imp. Dto PP|N|N|||rcafter|impppago|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3390
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. Dto PP"
         Height          =   255
         Index           =   11
         Left            =   -72630
         TabIndex        =   104
         Top             =   3195
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Total"
         Height          =   255
         Index           =   6
         Left            =   9960
         TabIndex        =   72
         Top             =   1740
         Visible         =   0   'False
         Width           =   1320
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   11
      Left            =   7320
      MaxLength       =   5
      TabIndex        =   100
      Tag             =   "Descuento P.Pago|N|N|0|99.90|rcafter|dtoppago|#0.00|N|"
      Text            =   "Text1 7"
      Top             =   2115
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   8685
      MaxLength       =   5
      TabIndex        =   99
      Tag             =   "Descuento General|N|N|0|99.90|rcafter|dtognral|#0.00|N|"
      Text            =   "Text1 7"
      Top             =   2115
      Width           =   525
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   4485
      MaxLength       =   30
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   2205
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   5145
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   67
      Text            =   "Text2"
      Top             =   2205
      Width           =   3525
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   4485
      MaxLength       =   30
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   1845
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   5145
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   65
      Text            =   "Text2"
      Top             =   1845
      Width           =   3525
   End
   Begin VB.Frame Frame2 
      Height          =   710
      Index           =   0
      Left            =   30
      TabIndex        =   56
      Top             =   540
      Width           =   11415
      Begin VB.CheckBox Check1 
         Caption         =   "Intracomunitario"
         Height          =   255
         Index           =   1
         Left            =   9870
         TabIndex        =   6
         Tag             =   "Intracomunitario|N|N|0|1|rcafter|intracom||N|"
         Top             =   390
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   31
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepción|F|N|||rcafter|fecrecep|dd/mm/yyyy|N|"
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   5850
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "Nombre Tercero|T|N|||rcafter|nomsocio||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   240
         Width           =   3750
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   4965
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod.Tercero|N|N|0|999999|rcafter|codsocio|000000|S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||rcafter|fecfactu|dd/mm/yyyy|S|"
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Nº Factura|T|N|||rcafter|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   315
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   255
         Index           =   0
         Left            =   9870
         TabIndex        =   5
         Tag             =   "Contabilizado|N|N|0|1|rcafter|intconta||N|"
         Top             =   150
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "F. Recepción"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   60
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         Height          =   255
         Index           =   0
         Left            =   4170
         TabIndex        =   59
         Top             =   270
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   4695
         ToolTipText     =   "Buscar tercero"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Factura"
         Height          =   255
         Index           =   29
         Left            =   1470
         TabIndex        =   58
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   270
         TabIndex        =   57
         Top             =   135
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   13
      Left            =   3555
      MaxLength       =   4
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   720
      Width           =   540
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   13
      Left            =   4140
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   62
      Text            =   "Text2"
      Top             =   720
      Width           =   3285
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   17
      Left            =   7500
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   55
      Text            =   "ABCDKFJADKSFJAK"
      Top             =   5430
      Visible         =   0   'False
      Width           =   1725
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   480
      Top             =   5160
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
      Left            =   1920
      Top             =   5160
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
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   75
      TabIndex        =   48
      Top             =   6210
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
         TabIndex        =   49
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10485
      TabIndex        =   33
      Top             =   6300
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9315
      TabIndex        =   32
      Top             =   6300
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Lineas Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir Factura"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   51
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10485
      TabIndex        =   47
      Top             =   6300
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
   Begin MSAdodcLib.Adodc Data4 
      Height          =   465
      Left            =   6840
      Top             =   765
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
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
   Begin VB.Label Label1 
      Caption         =   "Dto. P.P"
      Height          =   255
      Index           =   25
      Left            =   6660
      TabIndex        =   102
      Top             =   2115
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Dto. Gral"
      Height          =   255
      Index           =   26
      Left            =   7995
      TabIndex        =   101
      Top             =   2115
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Enabled         =   0   'False
      Height          =   240
      Index           =   2
      Left            =   4455
      Picture         =   "frmTercHcoFact.frx":0A70
      ToolTipText     =   "Buscar población"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   6
      Left            =   4125
      Picture         =   "frmTercHcoFact.frx":0B72
      ToolTipText     =   "Buscar trabajador"
      Top             =   2220
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador Pedido"
      Height          =   255
      Index           =   9
      Left            =   2565
      TabIndex        =   70
      Top             =   2220
      Width           =   1425
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   5
      Left            =   4140
      Picture         =   "frmTercHcoFact.frx":0C74
      ToolTipText     =   "Buscar trabajador"
      Top             =   1845
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador Albaran"
      Height          =   255
      Index           =   21
      Left            =   2565
      TabIndex        =   69
      Top             =   1890
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   255
      Index           =   1
      Left            =   2340
      TabIndex        =   64
      Top             =   720
      Width           =   855
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   3270
      Picture         =   "frmTercHcoFact.frx":0D76
      ToolTipText     =   "Buscar trabajador"
      Top             =   750
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Lote"
      Height          =   255
      Index           =   3
      Left            =   7500
      TabIndex        =   61
      Top             =   5250
      Visible         =   0   'False
      Width           =   615
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmTercHcoFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoFechaMovim As Date 'Fecha del Movim
Public hcoCodProve As Integer 'Codigo de Proveedor

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios  'Form Mto Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmFP As frmComFpa 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmHco As frmManHcoFruta  'Form Mto del historico de fruta
Attribute frmHco.VB_VarHelpID = -1

Private WithEvents frmMens As frmMensajes 'para asignacion de albaranes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmCargo As frmFVARConceptos ' conceptos de cargo
Attribute frmCargo.VB_VarHelpID = -1

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

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

'Si el cliente mostrado es de Varios o No
Dim EsDeVarios As Boolean
Private BuscaChekc As String

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal



Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Function AlbaranCero() As Boolean
Dim SQL As String

    SQL = "select * from rlifter where numalbar = 0 and " & ObtenerWhereCP(False)
    AlbaranCero = (TotalRegistrosConsulta(SQL) <> 0)

End Function


Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 4  'MODIFICAR
            If DatosOK Then
                '[Monica]16/07/2014: solo para el caso de montifrut y de IMG
                '                    si no hay albaranes asignados no hago recalculo de facturas
                '                    solo modifico cabecera
               If AlbaranCero And (vParamAplic.Cooperativa = 12 Or vParamAplic.Cooperativa = 15) Then
                    ModificaDesdeFormulario1 Me, 1
                    TerminaBloquear
               Else
                   If ModificarFactura Then
                        TerminaBloquear
    '                    PosicionarData
                   Else
                        '---- Laura 24/10/2006
                        'como no hemos modificado dejamos la fecha como estaba ya que ahora se puede modificar
                        Text1(1).Text = Me.Data1.Recordset!fecfactu
                   End If
               End If
               PosicionarData
            End If
            
         Case 5 'InsertarModificar LINEAS
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
'                PrimeraLin = False
'                If Data2.Recordset.EOF = True Then PrimeraLin = True
'                If InsertarLinea(NumLinea) Then
'                    'Comprobar si el Articulo tiene control de Nº de Serie
'                    ComprobarNSeriesLineas NumLinea
'                    If PrimeraLin Then
'                        CargaGrid DataGrid1, Data2, True
'                    Else
'                        CargaGrid2 DataGrid1, Data2
'                    End If
'                    BotonAnyadirLinea
'                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    NumRegElim = Data2.Recordset.AbsolutePosition
                    
'                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt text2(16), True
                    BloquearTxt text2(17), True
           
                    LLamaLineas Modo, 0, "DataGrid1"
                    PosicionarData
                    If (Not Data2.Recordset.EOF) And (Not Data2.Recordset.BOF) Then
                        SituarDataPosicion Data2, NumRegElim, ""
                    End If
                End If
'                Me.DataGrid1.Enabled = True
                Me.DataGrid2.Enabled = True
'                Me.DataGrid3.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1 'Busqueda, Insertar
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
            BloquearTxt text2(16), True
            BloquearTxt text2(17), True
'            If ModificaLineas = 1 Then 'INSERTAR
'                ModificaLineas = 0
'                DataGrid1.AllowAddNew = False
'                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
'            End If
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid1"
            PonerBotonCabecera True
'            Me.DataGrid1.Enabled = True
    End Select
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
'        LLamaLineas Modo, anc, "DataGrid3"
        
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
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos

        LimpiarDataGrids
        CadenaConsulta = "Select rcafter.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & Ordenacion
        

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

    'solo se puede modificar la factura si no esta contabilizada
    '++monica:añadida la condicion de solo si hay contabilidad
    If FactContabilizada Then
        TerminaBloquear
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(31)
'    PonerFocoChk Me.Check1(1)
        
    'Si es proveedor de Varios no se pueden modificar sus datos
'--monica
'    DeVarios = EsProveedorVarios(Text1(2).Text)
    DeVarios = False
    BloquearDatosTrans (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
On Error GoTo eModificarLinea




    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then
        TerminaBloquear
        Exit Sub '1= Insertar
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND numalbar=" & Data3.Recordset.Fields!numalbar & ""
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If

'    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
'        J = DataGrid1.Bookmark - DataGrid1.FirstRow
'        DataGrid1.Scroll 0, J
'        DataGrid1.Refresh
'    End If
'
'    anc = DataGrid1.Top
'    If DataGrid1.Row < 0 Then
'        anc = anc + 210
'    Else
'        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 20
'    End If
'
'    For J = 0 To 2
'        txtAux(J).Text = DataGrid1.Columns(J + 5).Text
'    Next J
'    Text2(16).Text = DataGrid1.Columns(J + 5).Text
'    For J = J + 1 To 8
'        txtAux(J - 1).Text = DataGrid1.Columns(J + 5).Text
'    Next J
'--monica
'    Text2(17).Text = DataGrid1.Columns(14).Text
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR LINEAS"
    PonerBotonCabecera False
    BloquearTxt text2(16), False 'Campo Ampliacion Linea
    BloquearTxt text2(17), False 'Campo Ampliacion Linea
'    PonerFoco txtAux(4)
    PonerFoco text2(16)
'    Me.DataGrid1.Enabled = False

eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim B As Boolean
    If grid = "DataGrid2" Then
        DeseleccionaGrid Me.DataGrid2
        B = (xModo = 1)
         For jj = 0 To txtaux3.Count - 1
            txtaux3(jj).Height = DataGrid2.RowHeight
            txtaux3(jj).Top = alto
            txtaux3(jj).visible = B
        Next jj
    End If
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (rcafter)
' y los registros correspondientes de las tablas cab. albaranes (rlifter)
Dim cad As String
Dim NumPedElim As Long
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    'solo se puede eliminar si no esta en la contabilidad
    If Me.Check1(0).Value = 1 Then Exit Sub
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    cad = "Cabecera de Facturas." & vbCrLf
    cad = cad & "-----------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la Factura:            "
    cad = cad & vbCrLf & "Tercero:  " & Text1(2).Text & " - " & Text1(3).Text
    cad = cad & vbCrLf & "Nº Fact.:  " & Text1(0).Text
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
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
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Factura", Err.Description
End Sub


'Private Sub cmdObserva_Click()
'    If Modo <> 2 And Modo <> 4 Then Exit Sub
'    If Me.FrameObserva.visible = False Then
''        Me.DataGrid1.visible = False
'        Me.FrameObserva.visible = True
''        Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(18).Picture
''        CargarICO Me.cmdObserva, "volver.ico"
'        Me.cmdObserva.ToolTipText = "volver lineas albaran"
'        BloqueaText3
'    Else
''        Me.DataGrid1.visible = True
'        Me.FrameObserva.visible = False
''        CargarICO Me.cmdObserva, "message.ico"
'        Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(12).Picture
'        Me.cmdObserva.ToolTipText = "ver observaciones albaran"
'    End If
'End Sub


Private Sub BloqueaText3()
Dim I As Byte
    'bloquear los Text3 que son las lineas de scafpa
    For I = 0 To 1
        BloquearTxt Text3(I), (Modo <> 4)
    Next I
    If Me.FrameObserva.visible Then
        For I = 4 To 8
            BloquearTxt Text3(I), (Modo <> 4)
        Next I
    End If
    'numpedpr, fecpedpr siempre bloqueados
    For I = 2 To 3
        BloquearTxt Text3(I), True
    Next I
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

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
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        If ModificaLineas <> 1 Then
'            Text2(16).Text = DBLet(Data2.Recordset.Fields!ampliaci)
'            Text2(17).Text = DBLet(Data2.Recordset.Fields!numlotes)
        End If
    Else
        text2(16).Text = ""
        text2(17).Text = ""
    End If

Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub CmdTraerAlbaranes_Click()
Dim SQL As String
Dim cadWHERE As String

    If AsignarAlbaranes Then
'        MsgBox "Proceso realizado correctamente.", vbExclamation
    End If

End Sub

Private Function AsignarAlbaranes() As Boolean
Dim SQL As String
Dim cadWHERE As String

Dim Coope As Long
Dim cSocios As String

    On Error GoTo eAsignaAlbaranes

    AsignarAlbaranes = False

    
    '[Monica]08/10/2013: pregunto pq se recalcula segun importe kilos
    If Me.Data3.Recordset!numalbar <> 0 Then
        If MsgBox("¿ Seguro que desea traer los albaranes ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            AsignarAlbaranes = True
            Exit Function
        End If
    End If




    '[Monica]19/09/2013: faltaria en este punto traer los albaranes de todos los socios de la cooperativa si es <> 1
    Coope = DevuelveValor("select codcoope from rsocios where codsocio = " & DBSet(Text1(2).Text, "N"))
    cSocios = "rhisfruta.codsocio = " & DBSet(Text1(2).Text, "N")
    
    If Coope <> 1 Then
        cSocios = cSocios & " or rhisfruta.codsocio in (select codsocio from rsocios where codcoope = " & Coope & ")"
    End If
    cSocios = "(" & cSocios & ")"

'[Monica]19/09/2013: cambio esta linea por la de abajo
'    Sql = "select * from rhisfruta where codsocio = " & DBSet(Text1(2).Text, "N")
    SQL = "select * from rhisfruta where " & cSocios

'08/10/2013: quitamos las comprobacion de las variedades
'    SQL = SQL & " and rhisfruta.codvarie in (select distinct codvarie from rlifter where codsocio = " & Data1.Recordset!CodSocio
'    SQL = SQL & " and numfactu = " & DBSet(Data1.Recordset!numfactu, "T") & " and fecfactu = " & DBSet(Data1.Recordset!fecfactu, "F") & ")"
    
'[Monica]19/09/2013: cambio esta linea por la de abajo
'    cadwhere = "codsocio = " & DBSet(Text1(2).Text, "N")
    cadWHERE = cSocios
    
'[Monica]19/09/2013:--- quiere que quite la condicion de que sean albaranes de la variedad que me dieron inicial
'    cadwhere = cadwhere & " and rhisfruta.codvarie in (select distinct codvarie from rlifter where codsocio = " & Data1.Recordset!CodSocio
'    cadwhere = cadwhere & " and numfactu = " & DBSet(Data1.Recordset!numfactu, "T") & " and fecfactu = " & DBSet(Data1.Recordset!fecfactu, "F") & ")"
    
    If TotalRegistrosConsulta(SQL) <> 0 Then
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 47
        frmMens.cadWHERE = cadWHERE
        
        frmMens.Show vbModal
        
        Set frmMens = Nothing
    Else
        MsgBox "No hay albaranes de este socio/variedad pdtes de asignar a la factura.", vbExclamation
    End If

    AsignarAlbaranes = True
    Exit Function
    
eAsignaAlbaranes:
    MuestraError Err.Number, "Asignar Albaranes", Err.Description
End Function


Private Sub DataGrid2_DblClick()
    If Data3.Recordset.EOF Then Exit Sub

    Set frmHco = New frmManHcoFruta
    
    frmHco.NroAlbaran = Data3.Recordset.Fields(3).Value
    frmHco.Show vbModal
    
    Set frmHco = Nothing

End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If Not Data3.Recordset.EOF Then
        Text3(2).Text = DBLet(Data3.Recordset.Fields!ImporteL, "N")
        If Text3(2).Text <> "0" Then
            FormateaCampo Text3(2)
        Else
            Text3(2).Text = ""
        End If
        
        'Observaciones
        Text3(4).Text = DBLet(Data3.Recordset.Fields!observa1, "T")
        Text3(5).Text = DBLet(Data3.Recordset.Fields!observa2, "T")
        Text3(6).Text = DBLet(Data3.Recordset.Fields!observa3, "T")
        Text3(7).Text = DBLet(Data3.Recordset.Fields!observa4, "T")
        Text3(8).Text = DBLet(Data3.Recordset.Fields!observa5, "T")
        
    Else
        For I = 0 To Text3.Count - 1
            If I <> 3 Then Text3(I).Text = ""
        Next I
        text2(0).Text = ""
        text2(1).Text = ""
    End If
    
End Sub



Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    btnPrimero = 15
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(9).Image = 15 'Mto Lineas Ofertas
        .Buttons(10).Image = 10 'Imprimir
        .Buttons(12).Image = 11  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    Me.SSTab1.Tab = 0
      
    'cargar IMAGES de busqueda
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
      
    '[Monica]12/03/2014: para montifrut cambio el nivel de usuario de los botones modificar y eliminar (estaba a 0)
    If vParamAplic.Cooperativa = 12 Then
        Me.Toolbar1.Buttons(5).Tag = 1
        Me.Toolbar1.Buttons(6).Tag = 1
    End If
      
    LimpiarCampos   'Limpia los campos TextBox
     
    Me.FrameObserva.visible = True
    
    Me.CmdTraerAlbaranes.Picture = frmPpal.imgListComun.ListImages(16).Picture '--monica antes 41
    Me.CmdTraerAlbaranes.ToolTipText = "Traer albaranes"
    Me.CmdTraerAlbaranes.visible = (vParamAplic.Cooperativa = 12)
    Me.CmdTraerAlbaranes.Enabled = (vParamAplic.Cooperativa = 12)
    
    Me.cmdAnticipos.Picture = frmPpal.imgListPpal.ListImages(9).Picture
    Me.cmdAnticipos.visible = (vParamAplic.Cooperativa = 12)
    Me.cmdAnticipos.Enabled = (vParamAplic.Cooperativa = 12)
        
    VieneDeBuscar = False
            
    '## A mano
    NombreTabla = "rcafter"
    NomTablaLineas = "rlifter" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY rcafter.fecrecep desc ,rcafter.codsocio, rcafter.numfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
    Else
        CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
        End If
        'Poner los grid sin apuntar a nada
    Else
         PonerModo 0
    End If
    LimpiarDataGrids
    PrimeraVez = False

End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1(0).Value = 0
    Me.Check1(1).Value = 0
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
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB
        CadenaConsulta = CadenaConsulta & " GROUP BY rcafter.codsocio, rcafter.numfactu,rcafter.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda
'        Text1(0).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmCargo_DatoSeleccionado(CadenaSeleccion As String)
    Text1(35).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod concepto
    text2(35).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

'--monica
'Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
''Formulario Mantenimiento C. Postales
'Dim Indice As Byte
'Dim devuelve As String
'
'        Indice = 7
'        Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
'        Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)  'Poblacion
'        'provincia
'        Text1(Indice + 2).Text = devuelve
'End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 10
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    text2(10).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        If InsertarAlbaranes(CadenaSeleccion) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            CargaGrid DataGrid2, Data1, True
            CargaGrid DataGrid3, Data4, True
        End If
    End If
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de socio
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Trans
End Sub

Private Sub frmPV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores Varios
Dim Indice As Byte

    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosProveVario (Text1(Indice).Text)
End Sub

'--monica
'Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
''Form Mantenimiento de Trabajadores
'Dim Indice As Byte
'
'    Indice = Val(Me.imgBuscar(4).Tag)
'    If Indice = 4 Then
'        Indice = Indice + 9
'        Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
'        Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
'    Else
'        Text3(Indice - 5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
'        Text2(Indice - 5).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
'    End If
'End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. socio
            PonerFoco Text1(2)
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            Indice = 2
            PonerFoco Text1(Indice)
            
'--monica
'        Case 1 'NIF para proveedor de Varios
'            Set frmPV = New frmComProveV
'            frmPV.DatosADevolverBusqueda = "0|"
'            frmPV.Show vbModal
'            Set frmPV = Nothing
'            Indice = 7
'            PonerFoco Text1(Indice)
            
'--monica
'        Case 2 'Cod. Postal
'            Set frmCP = New frmCPostal
'            frmCP.DatosADevolverBusqueda = "0"
'            frmCP.Show vbModal
'            Set frmCP = Nothing
'            Indice = 7
'            VieneDeBuscar = True
'            PonerFoco Text1(Indice)
      
         Case 3 'Forma de Pago
            Indice = 10
            PonerFoco Text1(Indice)
            Set frmFP = New frmComFpa
            frmFP.DatosADevolverBusqueda = "0|"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
'--monica
'        Case 4, 5, 6 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
'            Me.imgBuscar(4).Tag = Index
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
'            Set frmT = Nothing
'            If Index = 4 Then
'                PonerFoco Text1(13)
'            Else
'                PonerFoco Text3(Index - 5)
'            End If

        Case 7 ' concepto de cargo
            Set frmCargo = New frmFVARConceptos
            frmCargo.DatosADevolverBusqueda = "0|1|"
            frmCargo.CodigoActual = Text1(35).Text
            frmCargo.Show vbModal
            Set frmCargo = Nothing
            PonerFoco Text1(35)
        

    End Select
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Factura
'    BotonImprimir (53) '53: Informe de Facturas
End Sub


Private Sub mnLineas_Click()
    BotonMtoLineas 1, "Facturas"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then BotonModificarLinea
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: scafpc
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafpa
            If BloqueaAlbxFac Then BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim SQL As String
On Error GoTo EBloqueaAlb

    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM rcafter "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea TODAS las lineas de la factura
Dim SQL As String
    
    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM rlifter "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
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
    If Index = 9 Then HaCambiadoCP = False 'CPostal
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
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 31 'Fecha factura,fecha recepcion
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
'--monica
'        Case 13 'Cod trabajador
'            Text2(Index).Text = PonerNombreDeCod(Text1(Index), cAgro, "straba", "nomtraba", "codtraba")

        Case 2 'Cod. socio
            If Modo = 1 Then 'Modo=1 Busqueda
                '[Monica]24/10/2013: si estoy en busqueda no traigo nada (puede q me hayan cambiado el nombre del socio tercero)
                'Text1(Index + 1).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
            Else
                PonerDatosTransportista (Text1(Index).Text)
            End If
        
        Case 4 'NIF
            If Not EsDeVarios Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(4).Text = DBLet(Data1.Recordset!nifProve, "T") Then
                    Exit Sub
                End If
            End If
            PonerDatosProveVario (Text1(Index).Text)
            
'--monica
'        Case 7 'Cod. Postal
'            If Text1(Index).Locked Then Exit Sub
'            If Text1(Index).Text = "" Then
'                Text1(Index + 1).Text = ""
'                Text1(Index + 2).Text = ""
'                Exit Sub
'            End If
'            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
'                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
'                 Text1(Index + 2).Text = devuelve
'            End If
'            VieneDeBuscar = False
        
        
        Case 10 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(Text1(Index), "forpago", "nomforpa")
            Else
                text2(Index).Text = ""
            End If
            
        Case 11, 12 'Descuentos
            If Modo = 4 Then 'comprobar que el dato a cambiado
                If Index = 11 Then
                    If CCur(Text1(Index).Text) = CCur(Data1.Recordset!DtoPPago) Then Exit Sub
                ElseIf Index = 12 Then
                    If CCur(Text1(Index).Text) = CCur(Data1.Recordset!DtoGnral) Then Exit Sub
                End If
            End If
            
            If Modo = 3 Or Modo = 4 Then
                If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4 'Tipo 4: Decimal(4,2)
                If Not ActualizarDatosFactura Then
                   If Index = 11 Then Text1(Index).Text = Data1.Recordset!DtoPPago
                   If Index = 12 Then Text1(Index).Text = Data1.Recordset!DtoGnral
                   FormateaCampo Text1(Index)
                End If
            End If
            
         Case 35 ' concepto de cargo
            If PonerFormatoEntero(Text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(Text1(Index), "fvarconce", "nomconce", "codconce", "N")
            Else
                text2(Index).Text = ""
            End If
         
            
         Case 36 ' Importe de cargo
            PonerFormatoDecimal Text1(Index), 1
            
            
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select " & NombreTabla & ".* from (" & NombreTabla & " LEFT OUTER JOIN rlifter ON " & NombreTabla & ".codsocio=rlifter.codsocio AND " & NombreTabla & ".numfactu=rlifter.numfactu AND " & NombreTabla & ".fecfactu=rlifter.fecfactu) "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB
        CadenaConsulta = CadenaConsulta & " GROUP BY rcafter.codsocio, rcafter.numfactu, rcafter.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim devuelve As String
    
    'Llamamos a al form
    '##A mano
    cad = ""
'        cad = cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
        cad = cad & ParaGrid(Text1(0), 18, "Nº Factura")
        cad = cad & ParaGrid(Text1(1), 15, "Fecha Fac.")
        cad = cad & ParaGrid(Text1(2), 12, "Socio")
        cad = cad & ParaGrid(Text1(3), 55, "Nombre Socio")
        tabla = NombreTabla
        Titulo = "Facturas"
        devuelve = "0|1|2|"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'--monica
'        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges

'        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
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
        If Modo = 1 Then PonerFoco Text1(0)
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        LLamaLineas Modo, 0, "DataGrid2"
'        LLamaLineas Modo, 0, "DataGrid3"
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
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafpc de la factura seleccionada
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafpa
    CargaGrid DataGrid2, Data3, True
    CargaGrid DataGrid3, Data4, True
    
   
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single
Dim ImpTotal As Currency

    On Error Resume Next

    Me.CmdTraerAlbaranes.visible = False

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Poner la base imponible (impbruto - dtoppago - dtognral
    BrutoFac = CSng(Text1(14).Text) - CSng(Text1(15).Text) - CSng(Text1(16).Text)
    Text1(17).Text = Format(BrutoFac, FormatoImporte)
'    Text1(32).Text = Format(BrutoFac, FormatoImporte)
    
    'poner descripcion campos
    text2(10).Text = PonerNombreDeCod(Text1(10), "forpago", "nomforpa")
'--monica
'    Text2(13).Text = PonerNombreDeCod(Text1(13), conAri, "straba", "nomtraba", "codtraba")
    
'    ImpTotal = DevuelveValor("select sum(importel) from tcafpv where " & ObtenerWhereCP(False))
'    Text5.Text = Format(ImpTotal, FormatoImporte)
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    Me.CmdTraerAlbaranes.visible = (vParamAplic.Cooperativa = 12)
    
'++monica
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
'++
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
    
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo
    
    BuscaChekc = ""
    
    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
    ActualizarToolbar Modo, Kmodo
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
        
    cmdAnticipos.visible = B
    cmdAnticipos.Enabled = B
    CmdTraerAlbaranes.visible = B
    CmdTraerAlbaranes.Enabled = B
    
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    '---- laura 24/10/2006: si ponemos las claves de la tabla con ON UPDATE CASCADE
    'podemos permitir modificar la fecha de la factura que es clave primaria
'    If Modo = 4 Then BloquearTxt Text1(1), False
    
    For I = 0 To 9
        Text1(I).Enabled = (Modo = 1)
    Next I
    For I = 14 To 34
        Text1(I).Enabled = (Modo = 1)
    Next I
    
    Text1(31).Enabled = (Modo = 1 Or Modo = 4)
    
    Me.Check1(0).Enabled = (Modo = 1 Or Modo = 3)
    Me.Check1(1).Enabled = (Modo = 1 Or Modo = 3)
    
    B = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), B, True
    BloquearTxt Text1(3), B 'referencia
    
    'Importes siempre bloqueados
    For I = 14 To 30
        BloquearTxt Text1(I), (Modo <> 1)
    Next I
    For I = 32 To 34
        BloquearTxt Text1(I), (Modo <> 1)
    Next I

    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(17).BackColor = &HFFFFC0
    Text1(27).BackColor = &HFFFFC0
    Text1(28).BackColor = &HFFFFC0
    Text1(29).BackColor = &HFFFFC0
    Text1(30).BackColor = &HC0C0FF
    
    'bloquear los Text3 que son las lineas de scafac1
    BloqueaText3
    
    'Si no es modo lineas Boquear los TxtAux
'    For i = 0 To txtAux.Count - 1
'        BloquearTxt txtAux(i), (Modo <> 5)
'    Next i
'    BloquearTxt txtAux(8), True
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For I = 0 To txtaux3.Count - 1
        BloquearTxt txtaux3(I), (Modo <> 1)
    Next I
''    For i = 0 To Text4.Count - 1
''        BloquearTxt Text4(i), (Modo <> 1)
''    Next i
    
    'ampliacion linea
'    b = (Modo = 5) And Me.DataGrid1.visible
    'Modo Linea de Albaranes
'    Me.Label1(35).visible = b
    Me.Label1(3).visible = B
'    Me.Text2(16).visible = b
    Me.text2(17).visible = B
    BloquearTxt text2(16), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)
    BloquearTxt text2(17), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)

    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = B
    Next I
    Me.imgBuscar(0).Enabled = (Modo = 1)
    Me.imgBuscar(1).visible = False
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
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


Private Function DatosOK() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean
On Error GoTo EDatosOK

    DatosOK = False
    
    'Para que no den errores los 0's de los importes de dtos
    ComprobarDatosTotales
        
    'comprobamos datos OK de la tabla scafac
    B = CompForm(Me) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not B Then Exit Function
       
    DatosOK = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim B As Boolean
Dim I As Byte
On Error GoTo EDatosOkLinea

'    DatosOkLinea = False
'    b = True
'
'    For i = 0 To txtAux.Count - 1
'        If i = 4 Or i = 5 Or i = 6 Then
'            If txtAux(i).Text = "" Then
'                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
'                b = False
'                PonerFoco txtAux(i)
'                Exit Function
'            End If
'        End If
'    Next i
'
'    DatosOkLinea = b
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
    If Index = 17 And KeyAscii = 13 Then 'campo nº de lote y ENTER
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text3_LostFocus(Index As Integer)
    Select Case Index
'--monica
'        Case 0, 1 'trabajador
'            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
        Case 8 'observa 5
            PonerFocoBtn Me.cmdAceptar
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos

        Case 5  'Modificar
            mnModificar_Click
        Case 6  'Borrar
            mnEliminar_Click
        Case 9  'Lineas
            mnLineas_Click
        Case 10 'Imprimir Albaran
            mnImprimir_Click
        Case 12    'Salir
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
    
    
Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim SQL As String
Dim vWhere As String
Dim B As Boolean

'    On Error GoTo EModificarLinea
'
'    ModificarLinea = False
'    If Data2.Recordset.EOF Then Exit Function
'
'    vWhere = ObtenerWhereCP(True)
'    vWhere = vWhere & " AND numalbar='" & Data3.Recordset.Fields!numalbar & "'"
'    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
'
'    If DatosOkLinea() Then
'        SQL = "UPDATE " & NomTablaLineas & " SET "
'        SQL = SQL & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
'        SQL = SQL & "precioar= " & DBSet(txtAux(4).Text, "N") & ", "
'        SQL = SQL & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
'        SQL = SQL & "importel= " & DBSet(txtAux(7).Text, "N")
'        SQL = SQL & ", numlotes=" & DBSet(Text2(17).Text, "T")
'        SQL = SQL & vWhere
'    End If
'
'    If SQL <> "" Then
'        'actualizar la factura y vencimientos
'        b = ModificarFactura(SQL)
'        ModificarLinea = b
'    End If
'
'EModificarLinea:
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description
'        b = False
'    End If
'    ModificarLinea = b
End Function


Private Sub PonerBotonCabecera(B As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not B
'    DataGrid3.Enabled = Not b
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim B As Boolean
Dim Opcion As Byte
Dim SQL As String

    On Error GoTo ECargaGRid

'    b = DataGrid1.Enabled
 
    Select Case vDataGrid.Name
        Case "DataGrid2"
            Opcion = 2
        Case "DataGrid3"
            Opcion = 3
    End Select
    
    SQL = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
    CargaGrid2 vDataGrid, vData
    
    vDataGrid.ScrollBars = dbgAutomatic
    
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not B
'    PrimeraVez = False
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String

    On Error GoTo ECargaGRid
    
    vData.Refresh
    Select Case vDataGrid.Name
         Case "DataGrid2" 'albaranes x articulo
            'SQL = "SELECT codsocio,numfactu,fecfactu,numalbar, fechaalb, importel,observa1,observa2,observa3,observa4,observa5  "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(0)|T|Albarán|800|;S|txtAux3(1)|T|Fecha|1000|;S|txtAux3(2)|T|Código|900|;S|txtAux3(3)|T|Variedad|2000|;"
            tots = tots & "S|txtAux3(4)|T|Kilos Netos|1300|;S|txtAux3(5)|T|Precio|1100|;"
            tots = tots & "S|txtAux3(6)|T|Importe|1100|;S|txtAux3(7)|T|Pr.Estimado|1100|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla tots, DataGrid2, Me
        
            DataGrid2_RowColChange 1, 1
    
        Case "DataGrid3" 'anticipos
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux4(4)|T|Factura|1100|;S|txtAux4(5)|T|Fecha|1100|;"
            arregla tots, DataGrid3, Me
        
            'DataGrid3_RowColChange 1, 1
        
    
    
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Private Sub txtAux_GotFocus(Index As Integer)
'    ConseguirFoco txtAux(Index), Modo
'End Sub
'
'Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
''Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
'    KEYdown KeyCode
'End Sub
'
'
'Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
'
'
'Private Sub txtAux_LostFocus(Index As Integer)
'
'    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
'
'    Select Case Index
'        Case 4 'Precio
'            If txtAux(Index).Text <> "" Then
'                PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
'            End If
'
'        Case 5, 6 'Descuentos
'            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
'            If Index = 6 Then PonerFoco Me.Text2(16)
'
'        Case 7 'Importe Linea
'            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
'    End Select
'
'    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
'        If txtAux(1).Text = "" Then Exit Sub
''        txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, vParamAplic.TipoDtos)
'        txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, 0, 0)
'        PonerFormatoDecimal txtAux(7), 1
'    End If
'End Sub
'

Private Sub BotonMtoLineas(numTab As Integer, cad As String)
    Me.SSTab1.Tab = numTab
    
'    If Me.DataGrid1.visible Then 'Lineas de Albaranes
'        If Me.Data2.Recordset.RecordCount < 1 Then
'            MsgBox "La factura no tiene lineas.", vbInformation
'            Exit Sub
'        End If
'        TituloLinea = cad
'
'        ModificaLineas = 0
'        PonerModo 5
'        PonerBotonCabecera True
'    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim cta As String
Dim B As Boolean
Dim vSeccion As CSeccion

    On Error GoTo FinEliminar

        B = False
        Eliminar = False
        If Data1.Recordset.EOF Then Exit Function
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If vSeccion.AbrirConta Then
                ConnConta.BeginTrans
            Else
                Exit Function
            End If
        End If
        
        conn.BeginTrans
        
        
        'Eliminar en la tabla pagos de la Contabilidad: spagop
        '------------------------------------------------
        If Not vParamAplic.ContabilidadNueva Then
            cta = DevuelveDesdeBDNew(cAgro, "rsocios_seccion", "codmacpro", "codsocio", Text1(2).Text, "N", , "codsecci", vParamAplic.Seccionhorto, "N")
            SQL = " ctaprove='" & cta & "' AND numfactu='" & Data1.Recordset.Fields!numfactu & "'"
            SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!fecfactu, FormatoFecha) & "'"
            ConnConta.Execute "Delete from spagop WHERE " & SQL
        End If
        
        B = True
        
        'Eliminar en tablas de factura de Ariagro: tcafter, tlifter
        '---------------------------------------------------------------
        If B Then
            SQL = " " & ObtenerWhereCP(True)
        
            ' actualizamos el importe asignado a cada albaran a cero
            Sql2 = "update rhisfruta set cobradosn = 0 "
            Sql2 = Sql2 & " where numalbar in (select numalbar from rlifter " & SQL & ")"
            conn.Execute Sql2
        
            'Lineas de anticipos (rliantifter)
            conn.Execute "Delete from rliantifter " & SQL
        
            'Lineas de facturas (rlifter)
            conn.Execute "Delete from " & NomTablaLineas & SQL
        
            'Cabecera de facturas (tcafpc)
            conn.Execute "Delete from " & NombreTabla & SQL
        End If
        
        'Eliminar los movimientos generados por el albaran que genero la factura
        '-----------------------------------------------------------------------
        If B Then
        
        End If
        
'        b = True
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description
        B = False
    End If
    If Not B Then
        conn.RollbackTrans
        ConnConta.RollbackTrans
    Else
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
    Eliminar = B
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid2, Data3, False
    CargaGrid DataGrid3, Data4, False
'    CargaGrid DataGrid1, Data2, False
'    CargaGrid DataGrid3, Data4, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             If Modo <> 5 Then
                PonerModo 2
                PonerCampos
             End If
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
Dim SQL As String
On Error Resume Next
    SQL = "codsocio= " & Text1(2).Text & " and numfactu= '" & Text1(0).Text & "' and fecfactu='" & Format(Text1(1).Text, FormatoFecha) & "' "
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
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
Dim SQL As String
    
    Select Case Opcion
        Case 2 ' lineas
            SQL = "SELECT codsocio,numfactu,fecfactu,numalbar, fechaalb,rlifter.codvarie, variedades.nomvarie,rlifter.kilosnet,round(importel / rlifter.kilosnet,4), importel,prestimado,observa1,observa2,observa3,observa4,observa5  "
            SQL = SQL & " FROM rlifter, variedades " 'cabeceras albaranes de la factura
            
            If enlaza Then
                SQL = SQL & " " & ObtenerWhereCP(True)
                'lineas factura proveedor
            Else
                SQL = SQL & " WHERE numfactu = -1"
            End If
            SQL = SQL & " and rlifter.codvarie = variedades.codvarie "
            SQL = SQL & " ORDER BY codsocio, numfactu, fecfactu,rlifter.numalbar "
    
        Case 3 ' anticipos
            SQL = "SELECT codsocio,numfactu,fecfactu,codsocioanti, numfactuanti, fecfactuanti  "
            SQL = SQL & " FROM rliantifter " 'anticipos de la factura
            
            If enlaza Then
                SQL = SQL & " " & ObtenerWhereCP(True)
                'lineas factura proveedor
            Else
                SQL = SQL & " WHERE numfactu = '-1'"
            End If
            SQL = SQL & " ORDER BY 1,2,3,4,5,6 "
        
    End Select
    
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean

        B = ((Modo = 2) Or (Modo = 5 And ModificaLineas = 0)) And Check1(0).Value = 0
        'Modificar
        Toolbar1.Buttons(5).Enabled = B
        Me.mnModificar.Enabled = B
        'eliminar
        Toolbar1.Buttons(6).Enabled = B
        Me.mnEliminar.Enabled = B
            
'        b = (Modo = 2)
'        'Mantenimiento lineas
'        Toolbar1.Buttons(9).Enabled = b
'        Me.mnLineas.Enabled = b
        'Imprimir
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnImprimir.Enabled = b
        
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
End Sub


Private Sub PonerDatosTransportista(codTrans As String, Optional NIFTrans As String)
Dim vSoc As cSocio
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If codTrans = "" Then
        LimpiarDatosTrans
        Exit Sub
    End If

    Set vSoc = New cSocio
    'si se ha modificado el proveedor volver a cargar los datos
    If vSoc.Existe(codTrans) Then
        If vSoc.LeerDatos(codTrans) Then
        
'--monica
'            EsDeVarios = vProve.DeVarios
'            BloquearDatosProve (EsDeVarios)
'++monica
            EsDeVarios = False
            BloquearDatosTrans (EsDeVarios)
            
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(Text1(2).Text) = CLng(Data1.Recordset!codTrans) Then
                    Set vSoc = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(2).Text = vSoc.Codigo
            FormateaCampo Text1(2)
            
            If (Modo = 3) Or (Modo = 4) Then
                Text1(3).Text = vSoc.Nombre  'Nom prove
                Text1(6).Text = vSoc.Direccion
                Text1(7).Text = vSoc.CPostal
                Text1(8).Text = vSoc.Poblacion
                Text1(9).Text = vSoc.Provincia
                Text1(4).Text = vSoc.nif
                Text1(5).Text = DBLet(vSoc.Tfno1, "T")
            End If
            
            Observaciones = DBLet(vSoc.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del socio"
            End If
        End If
    Else
        LimpiarDatosTrans
    End If
    Set vSoc = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Socio", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del proveedor
Dim vSoc As cSocio
Dim B As Boolean
   
'    If nifProve = "" Then Exit Sub
'
'    Set vSoc = New CSocio
'    b = vSoc.LeerDatosProveVario(nifProve)
'    If b Then
'        Text1(3).Text = vSoc.Nombre   'Nom proveedor
'        Text1(6).Text = vSoc.Domicilio
'        Text1(7).Text = vSoc.CPostal
'        Text1(8).Text = vSoc.Poblacion
'        Text1(9).Text = vSoc.Provincia
'        Text1(5).Text = DBLet(vSoc.Tfno1, "T")
'    End If
'    Set vSoc = Nothing
End Sub


Private Sub LimpiarDatosTrans()
Dim I As Byte

    For I = 3 To 9
        Text1(I).Text = ""
    Next I
End Sub
   

Private Function ModificaAlbxFac() As Boolean
Dim SQL As String
Dim B As Boolean
On Error GoTo EModificaAlb
    
    ModificaAlbxFac = False
    If Data3.Recordset.EOF Then
        ModificaAlbxFac = True
        Exit Function
    End If
    'comprobar datos OK de la scafac1
     B = CompForm(Me) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not B Then Exit Function

'--monica
'    SQL = "UPDATE scafpa SET codtrab2=" & DBSet(Text3(0).Text, "N", "S") & ", "
'    SQL = SQL & "codtrab1=" & DBSet(Text3(1).Text, "N", "S")
    If Me.FrameObserva.visible Then
        SQL = "UPDATE rlifter SET "
        SQL = SQL & " observa1=" & DBSet(Text3(4).Text, "T")
        SQL = SQL & ", observa2=" & DBSet(Text3(5).Text, "T")
        SQL = SQL & ", observa3=" & DBSet(Text3(6).Text, "T")
        SQL = SQL & ", observa4=" & DBSet(Text3(7).Text, "T")
        SQL = SQL & ", observa5=" & DBSet(Text3(8).Text, "T")
        SQL = SQL & ObtenerWhereCP(True)
        SQL = SQL & " AND numalbar=" & Data3.Recordset.Fields!numalbar
        conn.Execute SQL
    End If
'--monica
'    SQL = SQL & ObtenerWhereCP(True)
'    SQL = SQL & " AND numalbar=" & Data3.Recordset.Fields!numalbar
'    Conn.Execute SQL
    ModificaAlbxFac = True
    
EModificaAlb:
If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Albaranes de factura", Err.Description
End Function



Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifac, añadir a la transaccion la modificación de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim SQL As String
Dim vFactu As CFacturaTer
Dim vSeccion As CSeccion

On Error GoTo EModFact

    bol = False
    conn.BeginTrans
    
    ModificarFactura = False
    
    If sqlLineas <> "" Then
        'actualizar el importe de la linea modificada
        MenError = "Modificando lineas de Factura."
        conn.Execute sqlLineas
    End If
    
    
    SQL = "update rhisfruta, rlifter set rhisfruta.impentrada = rlifter.importel where rlifter.numalbar = rhisfruta.numalbar and rlifter.fechaalb = rhisfruta.fecalbar "
    SQL = SQL & " and rlifter.codsocio = " & DBSet(Text1(2).Text, "N") & " and rlifter.numfactu = " & DBSet(Text1(0).Text, "T") & " and rlifter.fecfactu = " & DBSet(Text1(1).Text, "F")
    conn.Execute SQL
    
    
    
    'recalcular las bases imponibles x IVA
    MenError = "Recalcular importes IVA"
    bol = ActualizarDatosFactura
    
    
    If bol Then
        'modificamos la scafpc
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario1(Me, 1)
        
        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafpa
            bol = ModificaAlbxFac
            '++monica:añadida la condicion de solo si hay contabilidad
            If bol Then 'si se ha modificado la factura
                MenError = "Actualizando en Tesoreria"
                'y eliminar de tesoreria conta.spagop los registros de la factura
                
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                    If vSeccion.AbrirConta Then
                        ConnConta.BeginTrans
                    Else
                        bol = False
                    End If
                End If
                
                If bol Then
                    'antes de Eliminar en las tablas de la Contabilidad
                    Set vFactu = New CFacturaTer
                    bol = vFactu.LeerDatos(Text1(2).Text, Text1(0).Text, Text1(1).Text)
                    
                    If bol Then
                        'Eliminar de la spagop
                        If vParamAplic.ContabilidadNueva Then
                            SQL = " codmacta='" & vFactu.CtaTerce & "' AND numfactu='" & Data1.Recordset.Fields!numfactu & "'"
                            SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!fecfactu, FormatoFecha) & "'"
                            ConnConta.Execute "Delete from pagos WHERE " & SQL
                        
                        Else
                            SQL = " ctaprove='" & vFactu.CtaTerce & "' AND numfactu='" & Data1.Recordset.Fields!numfactu & "'"
                            SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!fecfactu, FormatoFecha) & "'"
                            ConnConta.Execute "Delete from spagop WHERE " & SQL
                        End If
                        'Volvemos a grabar en TESORERIA. Tabla de Contabilidad: sconta.spagop
                        If bol Then
                            bol = vFactu.InsertarEnTesoreria(MenError)
                        End If
                    End If
                    Set vFactu = Nothing
                End If
            End If
        
        SQL = "update rhisfruta, rlifter set rhisfruta.impentrada = 0 where rlifter.numalbar = rhisfruta.numalbar and rlifter.fechaalb = rhisfruta.fecalbar "
        SQL = SQL & " and rlifter.codsocio = " & DBSet(Text1(2).Text, "N") & " and rlifter.numfactu = " & DBSet(Text1(0).Text, "T") & " and rlifter.fecfactu = " & DBSet(Text1(1).Text, "F")
        conn.Execute SQL
        
        End If
    End If

EModFact:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ConnConta.CommitTrans
        
        vSeccion.CerrarConta
        Set vSeccion = Nothing

        ModificarFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        ModificarFactura = False
        MenError = "Actualizando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function



Private Function FactContabilizada() As Boolean
Dim cta As String, numasien As String
Dim vSeccion As CSeccion

On Error GoTo EContab

    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1(0).Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        cta = DevuelveDesdeBDNew(cAgro, "rsocios_seccion", "codmacpro", "codsocio", Text1(2).Text, "N", , "codsecci", vParamAplic.Seccionhorto, "N")
        If cta <> "" Then
            Set vSeccion = New CSeccion
            
            If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                If vSeccion.AbrirConta Then
                    numasien = DevuelveDesdeBDNew(cConta, "cabfactprov", "numasien", "codmacta", cta, "T", , "numfacpr", Text1(0).Text, "T", "fecfacpr", Text1(1).Text, "F")
                    If numasien <> "" Then
                        FactContabilizada = True
                        MsgBox "La factura esta contabilizada y no se puede modificar ni eliminar.", vbInformation
                    Else
                        FactContabilizada = False
                    End If
                End If
            Else
                FactContabilizada = False
            End If
            
            Set vSeccion = Nothing
        Else
            FactContabilizada = True
            Exit Function
        End If
    Else
        FactContabilizada = False
    End If
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function


Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtaux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtaux3(Index), Modo) Then Exit Sub
End Sub


'Private Sub Text4_GotFocus(Index As Integer)
'    ConseguirFoco Text4(Index), Modo
'End Sub
'
'Private Sub Text4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
'End Sub
'
'Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
'
'Private Sub Text4_LostFocus(Index As Integer)
'    If Not PerderFocoGnral(Text4(Index), Modo) Then Exit Sub
'
'    Select Case Index
'        Case 0 'Cta Contable
'            If Text4(Index).Text = "" Then Exit Sub
'            Text4(2).Text = PonerNombreCuenta(Text4(Index), Modo)
'            If Text4(2).Text = "" Then
'                PonerFoco Text4(Index)
'            End If
'
'        Case 1 ' Importe
'            If Text4(Index).Text <> "" Then PonerFormatoDecimal Text4(Index), 1
'
'    End Select
'End Sub
'


Private Sub BloquearDatosTrans(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol 'NIF
        Me.imgBuscar(1).Enabled = bol 'NIF
        Me.imgBuscar(2).Enabled = bol 'poblacion
        
        For I = 3 To 9 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
    
    End If
End Sub

'--monica
'Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
''Modifica los datos de la tabla de Proveedores Varios
'Dim vProve As CProveedor
'On Error GoTo EActualizarCV
'
'    ActualizarProveVarios = False
'
'    Set vProve = New CProveedor
'    If EsProveedorVarios(Prove) Then
'        vProve.NIF = NIF
'        vProve.Nombre = Text1(3).Text
'        vProve.Domicilio = Text1(6).Text
'        vProve.CPostal = Text1(7).Text
'        vProve.Poblacion = Text1(8).Text
'        vProve.Provincia = Text1(9).Text
'        vProve.TfnoAdmon = Text1(5).Text
'        vProve.ActualizarProveV (NIF)
'    End If
'    Set vProve = Nothing
'
'    ActualizarProveVarios = True
'
'EActualizarCV:
'    If Err.Number <> 0 Then
'        ActualizarProveVarios = False
'    Else
'        ActualizarProveVarios = True
'    End If
'End Function


Private Function ObtenerSelFactura() As String
'Cuando venimos desde dobleClick en Movimientos de Articulos para Albaranes ya
'Facturados, abrimos este form pero cargando los datos de la factura
'correspendiente al albaran que se selecciono
Dim cad As String
Dim Rs As ADODB.Recordset
On Error Resume Next

    cad = "SELECT codprove,numfactu,fecfactu FROM scafpa "
    cad = cad & " WHERE codprove=" & DBSet(hcoCodProve, "N") & " AND numalbar=" & DBSet(hcoCodMovim, "T")
    cad = cad & " AND fechaalb=" & DBSet(hcoFechaMovim, "F")

    Set Rs = New ADODB.Recordset
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then 'where para la factura
        cad = " WHERE codprove=" & Rs!codProve & " AND numfactu= '" & Rs!numfactu & "' AND fecfactu=" & DBSet(Rs!fecfactu, "F")
    Else
        cad = " where numfactu=-1"
    End If
    Rs.Close
    Set Rs = Nothing

    ObtenerSelFactura = cad
End Function



Private Function ActualizarDatosFactura() As Boolean
Dim vFactu As CFacturaTer
Dim cadSel As String

    Set vFactu = New CFacturaTer
    cadSel = ObtenerWhereCP(False)
    cadSel = "numalbar in (select numalbar from rlifter where " & cadSel & ")"
    vFactu.DtoPPago = CCur(Text1(11).Text)
    vFactu.DtoGnral = CCur(Text1(16).Text)
    vFactu.Intracomunitario = Check1(1).Value

    If vFactu.CalcularDatosFactura(cadSel, Text1(2).Text, True) Then '"tcafpa") Then
        Text1(14).Text = vFactu.BrutoFac
        Text1(15).Text = vFactu.ImpPPago
        Text1(16).Text = vFactu.ImpGnral
        Text1(17).Text = vFactu.BaseImp
        Text1(18).Text = vFactu.TipoIVA1
        Text1(19).Text = vFactu.TipoIVA2
        Text1(20).Text = vFactu.TipoIVA3
        Text1(21).Text = vFactu.PorceIVA1
        Text1(22).Text = vFactu.PorceIVA2
        Text1(23).Text = vFactu.PorceIVA3
        Text1(24).Text = vFactu.BaseIVA1
        Text1(25).Text = vFactu.BaseIVA2
        Text1(26).Text = vFactu.BaseIVA3
        Text1(27).Text = vFactu.ImpIVA1
        Text1(28).Text = vFactu.ImpIVA2
        Text1(29).Text = vFactu.ImpIVA3
        Text1(30).Text = vFactu.TotalFac
        Text1(32).Text = vFactu.BaseReten
        Text1(33).Text = vFactu.PorcReten
        Text1(34).Text = vFactu.ImpReten
        FormatoDatosTotales
        
        ActualizarDatosFactura = True
    Else
        ActualizarDatosFactura = False
        MuestraError Err.Number, "Recalculando Factura", Err.Description
    End If
    Set vFactu = Nothing
End Function


Private Sub FormatoDatosTotales()
Dim I As Byte

    For I = 14 To 17
'        Text1(I).Text = QuitarCero(Text1(I).Text)
        FormateaCampo Text1(I)
    Next I
    
    For I = 24 To 26
        If Text1(I).Text <> "" Then
            'Si la Base Imp. es 0
            If CSng(Text1(I).Text) = 0 Then
                Text1(I).Text = QuitarCero(Text1(I).Text)
                Text1(I - 3).Text = QuitarCero(Text1(I - 3).Text)
                Text1(I - 6).Text = QuitarCero(Text1(I - 6).Text)
                Text1(I + 3).Text = QuitarCero(Text1(I + 3).Text)
            Else
                FormateaCampo Text1(I)
                FormateaCampo Text1(I - 3)
                FormateaCampo Text1(I - 6)
                FormateaCampo Text1(I + 3)
            End If
        Else 'No hay Base Imponible
            Text1(I - 3).Text = QuitarCero(Text1(I - 3).Text)
            Text1(I - 6).Text = QuitarCero(Text1(I - 6).Text)
            Text1(I + 3).Text = ""
        End If
    Next I
End Sub



Private Sub ComprobarDatosTotales()
Dim I As Byte

    For I = 14 To 17
        Text1(I).Text = ComprobarCero(Text1(I).Text)
    Next I
End Sub




Private Function InsertarAlbaranes(Albaranes As String)
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim TotalKilos As Currency
Dim ImporteVar As Currency
Dim vImporte As Currency
Dim ImporteAlb As Currency
Dim PrecioAlb As Currency
Dim vSQL As String
Dim CadValues As String
Dim Rs2 As ADODB.Recordset
Dim ImporteTot As Currency
Dim Diferencia As Currency
Dim Importe As Currency
Dim Precio As Currency


    On Error GoTo eInsertarAlbaranes

    InsertarAlbaranes = False
    
    conn.BeginTrans
    
    
    '[Monica]08/10/2013: metemos en slog quien realiza la insercion de albaranes
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 15, vUsu, "Albaranes de Terceros: " & vbCrLf & Albaranes & vbCrLf & " de " & ObtenerWhereCP(False)
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    
    vSQL = "select codvarie, sum(kilosnet) kilosnet from rhisfruta where numalbar in ( " & Albaranes & ") group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    ImporteTot = 0
    
    CadValues = ""
    While Not Rs.EOF
        TotalKilos = DBLet(Rs!KilosNet, "N")
        ImporteVar = DevuelveValor("select sum(importel) from rlifter where " & ObtenerWhereCP(False) & " and codvarie = " & DBSet(Rs!codvarie, "N"))
    
        ImporteTot = ImporteTot + ImporteVar
    
        Sql2 = "select * from rhisfruta where numalbar in (" & Albaranes & ") and codvarie = " & DBSet(Rs!codvarie, "N")
        Set Rs2 = New ADODB.Recordset
        
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs2.EOF
            If TotalKilos <> 0 Then
                ImporteAlb = Round2(ImporteVar * Rs2!KilosNet / TotalKilos, 2)
            Else
                ImporteAlb = 0
            End If
            PrecioAlb = Round2(ImporteAlb / Rs2!KilosNet, 4)
            
            CadValues = CadValues & "(" & DBSet(Text1(2).Text, "N") & "," & DBSet(Text1(0).Text, "T") & "," & DBSet(Text1(1).Text, "F") & ","
            CadValues = CadValues & DBSet(Rs2!numalbar, "N") & "," & DBSet(Rs2!Fecalbar, "F") & "," & DBSet(Rs2!codvarie, "N") & "," & DBSet(Rs2!KilosNet, "N") & ","
            CadValues = CadValues & DBSet(ImporteAlb, "N") & "," & DBSet(PrecioAlb, "N") & ",0),"
            
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
        
        
        Rs.MoveNext
    
    Wend
    
    ' primero hay que borrar los albaranes que hayan
    vSQL = "delete from rlifter where " & ObtenerWhereCP(False)
    conn.Execute vSQL
    
    
    Set Rs = Nothing
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
    
        ' igual que el insert pero reemplaza las columnas existentes
        SQL = "insert into rlifter (codsocio, numfactu, fecfactu, numalbar, fechaalb, codvarie, kilosnet, importel, prestimado, descontado) "
        SQL = SQL & " values "
    
        conn.Execute SQL & CadValues
        
    End If
    
    '[Monica]08/10/2013: si no coinciden con el importe que habia en rlifter
    If Text1(14).Text <> ImporteTot Then
        Diferencia = ImporteSinFormato(Text1(14).Text) - ImporteTot
        Importe = 0
        SQL = "select * from rlifter where importel = 0 and " & ObtenerWhereCP(False)
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        TotalKilos = DevuelveValor("select sum(kilosnet) from rlifter where importel = 0 and " & ObtenerWhereCP(False))
        
        While Not Rs.EOF
            ImporteVar = Round2(Rs!KilosNet * Diferencia / TotalKilos, 2)
            Importe = Importe + ImporteVar
            
            Precio = Round2(ImporteVar / DBLet(Rs!KilosNet), 4)
            
            Sql2 = "update rlifter set importel = " & DBSet(ImporteVar, "N") & ", prestimado = " & DBSet(Precio, "N")
            Sql2 = Sql2 & " where " & ObtenerWhereCP(False) & " and numalbar = " & DBSet(Rs!numalbar, "N")
            
            conn.Execute Sql2
            
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
    End If
    
    
    
    
    conn.CommitTrans
    InsertarAlbaranes = True
    Exit Function
    
eInsertarAlbaranes:
    conn.RollbackTrans
    MuestraError Err.Number, "Insertando Albaranes", Err.Description
End Function

Private Sub cmdAnticipos_Click()
    If Modo <> 2 Then Exit Sub
    If Me.FrameAnticipos.visible = False Then
'        Me.DataGrid1.visible = False
        Me.FrameAnticipos.visible = True
        Me.FrameObserva.visible = False
        Me.cmdAnticipos.Picture = frmPpal.imgListPpal.ListImages(36).Picture
        Me.cmdAnticipos.ToolTipText = "Volver de Anticipos"
    Else
'        Me.DataGrid1.visible = True
        Me.FrameAnticipos.visible = False
        Me.FrameObserva.visible = True
        Me.cmdAnticipos.Picture = frmPpal.imgListPpal.ListImages(9).Picture
        Me.cmdAnticipos.ToolTipText = "Ver Anticipos de Liquidación"
    End If

End Sub

