VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManFactTranspor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Transportistas"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   9660
   Icon            =   "frmManFactTranspor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4425
      Left            =   60
      TabIndex        =   39
      Top             =   3600
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   7805
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   12582912
      TabCaption(0)   =   "Albaranes"
      TabPicture(0)   =   "frmManFactTranspor.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAlbaranes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Rectificativa"
      TabPicture(1)   =   "frmManFactTranspor.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo1(2)"
      Tab(1).Control(1)=   "Text1(12)"
      Tab(1).Control(2)=   "Text1(13)"
      Tab(1).Control(3)=   "Text1(14)"
      Tab(1).Control(4)=   "Text1(15)"
      Tab(1).Control(5)=   "Label1(11)"
      Tab(1).Control(6)=   "Label1(12)"
      Tab(1).Control(7)=   "Label1(13)"
      Tab(1).Control(8)=   "imgFec(1)"
      Tab(1).Control(9)=   "Label1(14)"
      Tab(1).ControlCount=   10
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1080
         Width           =   3330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   12
         Left            =   -71130
         MaxLength       =   7
         TabIndex        =   56
         Tag             =   "Nº Factura Rectificada|N|S|||rfacttra|rectif_numfactu|0000000|N|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   13
         Left            =   -70080
         MaxLength       =   10
         TabIndex        =   55
         Tag             =   "Fecha Factura Rectificada|F|S|||rfacttra|rectif_fecfactu|dd/mm/yyyy|N|"
         Top             =   1080
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Height          =   555
         Index           =   14
         Left            =   -74520
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   54
         Tag             =   "Motivo Rectif. Factura|T|S|||rfacttra|rectif_motivo|||"
         Text            =   "frmManFactTranspor.frx":0044
         Top             =   1830
         Width           =   8775
      End
      Begin VB.Frame FrameAlbaranes 
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
         Height          =   3720
         Left            =   60
         TabIndex        =   40
         Top             =   540
         Width           =   9390
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   10
            Left            =   1980
            MaxLength       =   7
            TabIndex        =   65
            Tag             =   "Transportista|T|N|||rfacttra_albaran|codtrans||S|"
            Text            =   "transpo"
            Top             =   1140
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   1
            Left            =   2850
            MaskColor       =   &H00000000&
            TabIndex        =   51
            ToolTipText     =   "Buscar concepto gasto"
            Top             =   1170
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   9
            Left            =   7110
            MaxLength       =   30
            TabIndex        =   50
            Tag             =   "Importe|N|N|||rfacttra_albaran|importe|###,##0.00||"
            Text            =   "Importe"
            Top             =   1170
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   8
            Left            =   6345
            MaxLength       =   30
            TabIndex        =   49
            Tag             =   "Kilos Netos|N|N|||rfacttra_albaran|kilosnet|###,##0||"
            Text            =   "K.Neto"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   7
            Left            =   5160
            MaxLength       =   30
            TabIndex        =   48
            Tag             =   "NumNotac|N|N|||rfacttra_albaran|numnotac|0000000|S|"
            Text            =   "numnotac"
            Top             =   1170
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   6
            Left            =   4365
            MaxLength       =   25
            TabIndex        =   47
            Tag             =   "Campo|N|N|||rfacttra_albaran|codcampo|00000000|N|"
            Text            =   "campo"
            Top             =   1170
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   2295
            MaxLength       =   7
            TabIndex        =   46
            Tag             =   "Albaran|N|N|||rfacttra_albaran|numalbar|0000000|S|"
            Text            =   "albaran"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   2970
            MaxLength       =   10
            TabIndex        =   45
            Tag             =   "Fecha Alb|F|N|||rfacttra_albaran|fecalbar|dd/mm/yyyy|S|"
            Text            =   "fecalbar"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   315
            MaxLength       =   7
            TabIndex        =   44
            Tag             =   "Tipo Movim.|T|N|||rfacttra_albaran|codtipom||S|"
            Text            =   "tipof"
            Top             =   1170
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   1035
            MaxLength       =   7
            TabIndex        =   43
            Tag             =   "Nº.Factura|N|N|||rfacttra_albaran|numfactu|0000000|S|"
            Text            =   "numfact"
            Top             =   1170
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   1665
            MaxLength       =   10
            TabIndex        =   42
            Tag             =   "Fecha Fact.|F|N|||rfacttra_albaran|fecfactu|dd/mm/yyyy|S|"
            Text            =   "FecFact"
            Top             =   1155
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   5
            Left            =   3645
            MaxLength       =   25
            TabIndex        =   41
            Tag             =   "Variedad|N|N|||rfacttra_albaran|codvarie|000000||"
            Text            =   "variedad"
            Top             =   1170
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   180
            TabIndex        =   52
            Top             =   30
            Visible         =   0   'False
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
         Begin MSDataGridLib.DataGrid DataGrid5 
            Bindings        =   "frmManFactTranspor.frx":00A9
            Height          =   3540
            Left            =   180
            TabIndex        =   53
            Top             =   30
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   6244
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
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   -74250
         MaxLength       =   10
         TabIndex        =   64
         Tag             =   "Tipo Movimiento|T|S|||rfacttra|rectif_codtipom||N|"
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Factura"
         Height          =   255
         Index           =   11
         Left            =   -74550
         TabIndex        =   61
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fac"
         Height          =   255
         Index           =   12
         Left            =   -70080
         TabIndex        =   60
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   13
         Left            =   -71130
         TabIndex        =   59
         Top             =   840
         Width           =   855
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   -69240
         Picture         =   "frmManFactTranspor.frx":00BE
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo"
         Height          =   255
         Index           =   14
         Left            =   -74520
         TabIndex        =   58
         Top             =   1590
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   30
      TabIndex        =   22
      Top             =   630
      Width           =   9555
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   240
         Width           =   1890
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   10
         Left            =   300
         MaxLength       =   10
         TabIndex        =   63
         Tag             =   "Tipo Movimiento|T|N|||rfacttra|codtipom||S|"
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pasa Aridoc"
         Height          =   195
         Index           =   2
         Left            =   7350
         TabIndex        =   6
         Tag             =   "Aridoc|N|N|0|1|rfacttra|pasaridoc|0||"
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   240
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Base Imponible|N|N|||rfacttra|baseimpo|###,##0.00||"
         Top             =   1770
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CAE3FD&
         Height          =   315
         Index           =   7
         Left            =   7290
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "Total Factura|N|N|||rfacttra|totalfac|###,##0.00||"
         Top             =   1770
         Width           =   1830
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   5595
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Importe Iva|N|N|||rfacttra|imporiva|###,##0.00||"
         Top             =   1770
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "Tipo Iva|N|N|0|99|rfacttra|tipoiva|00||"
         Text            =   "Text1"
         Top             =   1770
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2250
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   1770
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   4590
         MaxLength       =   6
         TabIndex        =   9
         Tag             =   "Porc.Iva|N|N|||rfacttra|porc_iva|##0.00||"
         Text            =   "123"
         Top             =   1770
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   4590
         MaxLength       =   6
         TabIndex        =   12
         Tag             =   "Porc.Retención|N|S|||rfacttra|porc_ret|##0.00||"
         Text            =   "123"
         Top             =   2340
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   5595
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Importe Retención|N|S|||rfacttra|impreten|#,##0.00||"
         Text            =   "123"
         Top             =   2340
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   240
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Base Retención|N|S|||rfacttra|basereten|###,##0.00||"
         Top             =   2340
         Width           =   1350
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   4590
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Tipo IRPF|N|N|0|3|rfacttra|tipoirpf|0|N|"
         Top             =   855
         Width           =   2130
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   195
         Index           =   1
         Left            =   7350
         TabIndex        =   5
         Tag             =   "Contabilizado|N|N|0|1|rfacttra|contabilizado|0||"
         Top             =   360
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   150
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Cod.Transportista|T|N|||rfacttra|codtrans||S|"
         Text            =   "Text1"
         Top             =   870
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Factura|N|S|||rfacttra|numfactu|0000000|S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   885
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Impreso"
         Height          =   195
         Index           =   0
         Left            =   7350
         TabIndex        =   4
         Tag             =   "Impreso|N|N|0|1|rfacttra|impreso|0||"
         Top             =   90
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   3210
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||rfacttra|fecfactu|dd/mm/yyyy|S|"
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   855
         Width           =   3150
      End
      Begin VB.Frame Frame5 
         Caption         =   "Total Factura"
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
         Height          =   1575
         Left            =   60
         TabIndex        =   24
         Top             =   1290
         Width           =   9435
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   10
            Left            =   180
            TabIndex        =   38
            Top             =   270
            Width           =   1455
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
            Height          =   255
            Index           =   9
            Left            =   7230
            TabIndex        =   37
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
            Height          =   255
            Index           =   7
            Left            =   5535
            TabIndex        =   36
            Top             =   270
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1950
            ToolTipText     =   "Buscar Iva"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Iva"
            Height          =   255
            Index           =   8
            Left            =   1620
            TabIndex        =   35
            Top             =   270
            Width           =   285
         End
         Begin VB.Label Label2 
            Caption         =   "% Iva"
            Height          =   255
            Left            =   4530
            TabIndex        =   34
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label18 
            Caption         =   "% Ret"
            Height          =   255
            Left            =   4530
            TabIndex        =   33
            Top             =   810
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Importe Retención"
            Height          =   255
            Left            =   5535
            TabIndex        =   32
            Top             =   810
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "Base Retención"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   31
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo IRPF"
         Height          =   255
         Index           =   3
         Left            =   4590
         TabIndex        =   29
         Top             =   585
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Factura"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fac"
         Height          =   255
         Index           =   29
         Left            =   3210
         TabIndex        =   27
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Transportista"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   585
         Width           =   930
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1290
         ToolTipText     =   "Buscar Socio"
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   2160
         TabIndex        =   25
         Top             =   0
         Width           =   855
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   4050
         Picture         =   "frmManFactTranspor.frx":0149
         ToolTipText     =   "Buscar fecha"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   45
      TabIndex        =   18
      Top             =   8040
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
         TabIndex        =   19
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   16
      Top             =   8130
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7335
      TabIndex        =   15
      Top             =   8130
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7410
         TabIndex        =   21
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8505
      TabIndex        =   17
      Top             =   8130
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3090
      Top             =   1860
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
      Top             =   7830
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
      Left            =   390
      Top             =   7890
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
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
      Height          =   405
      Left            =   780
      Top             =   7680
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   714
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
   Begin MSAdodcLib.Adodc Data5 
      Height          =   360
      Left            =   870
      Top             =   7560
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   635
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
   Begin MSAdodcLib.Adodc Data6 
      Height          =   405
      Left            =   930
      Top             =   7560
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   714
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
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManFactTranspor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public numalbar As String  ' venimos de pedidos para insertar envases paletizacion

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
Private WithEvents frmLFac As frmManLinFactTranspor 'Lineas de variedades de facturas transporte
Attribute frmLFac.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de iva de conta
Attribute frmTIva.VB_VarHelpID = -1

Private WithEvents frmVar As frmComVar 'Form Mto de variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTranspor 'Form Mto de transportistas
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'Form Mto de calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Form mensajes
Attribute frmMens.VB_VarHelpID = -1

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
Dim indice As Byte
Dim Facturas As String

Dim Cliente As String
Private BuscaChekc As String
Dim vTrans As CTransportista

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 1 'albaranes del transportista pendientes de facturar
            Set frmMens = New frmMensajes
            
            frmMens.OpcionMensaje = 24
            frmMens.cadWHERE = "codtrans = " & DBSet(Data1.Recordset!codTrans, "T")
            frmMens.Show vbModal
            Set frmMens = Nothing
            
            PonerFoco txtAux1(3)
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

Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub


Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
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
                CalcularTotales
                If ModificaCabecera Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
                    PonerCampos
                    PonerCamposLineas
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            Select Case ModificaLineas
                Case 1 'afegir llínia
                    InsertarLinea 0
                Case 2 'modificar llínies
                    ModificarLinea
                    PosicionarData
            End Select
'            PonerCampos
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
            LLamaLineas Modo, 0, "DataGrid5"
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            
'            ComprobarClasificacion
            
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid5.AllowAddNew = False
                If Not Data6.Recordset.EOF Then Data6.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid5"
            PonerModo 2
            DataGrid5.Enabled = True
            If Not Data1.Recordset.EOF Then _
                Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

            'Habilitar las opciones correctas del menu segun Modo
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid5.Enabled = True
            PonerFocoGrid DataGrid5
                
            PonerCampos
    End Select
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
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    
    Text1(5).Text = 0
    Text1(6).Text = 0
    Text1(7).Text = 0
    Text1(8).Text = vParamAplic.PorcreteFacSoc
    
    LimpiarDataGrids
    Combo1(0).ListIndex = 0
    Combo1(0).SetFocus
'    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
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
        anc = DataGrid5.Top
        If DataGrid5.Row < 0 Then
            anc = anc + 230 '440
        Else
            anc = anc + DataGrid5.RowTop(DataGrid5.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid5"
        
        
        'Si pasamos el control aqui lo ponemos en amarillo
        Combo1(0).SetFocus
        Combo1(0).BackColor = vbYellow
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
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select rfacttra.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
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
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(2) '*** 1r camp visible que siga PK ***
        
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

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
    If Data6.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    NumTabMto = Index
    If Index = 2 Then NumTabMto = 3
    ModificaLineas = 2
    PonerModo 5, Index
 

    
        vWhere = ObtenerWhereCP(False)
        If Not BloqueaRegistro("rfacttra", vWhere) Then
            TerminaBloquear
            Exit Sub
        End If
        If DataGrid5.Bookmark < DataGrid5.FirstRow Or DataGrid5.Bookmark > (DataGrid5.FirstRow + DataGrid5.VisibleRows - 1) Then
            J = DataGrid5.Bookmark - DataGrid5.FirstRow
            DataGrid5.Scroll 0, J
            DataGrid5.Refresh
        End If

    '    anc = ObtenerAlto(Me.DataGrid1)
        anc = DataGrid5.Top
        If DataGrid5.Row < 0 Then
            anc = anc + 210
        Else
            anc = anc + DataGrid5.RowTop(DataGrid5.Row) + 10
        End If

        For J = 3 To 4
            txtAux1(J).Text = DataGrid5.Columns(J).Text
        Next J

        For J = 5 To 9
            txtAux1(J).Text = DataGrid5.Columns(J + 1).Text
        Next J

        ModificaLineas = 2 'Modificar
        LLamaLineas ModificaLineas, anc, "DataGrid5"

        'Añadiremos el boton de aceptar y demas objetos para insertar
        Me.lblIndicador.Caption = "MODIFICAR"
        PonerModoOpcionesMenu (Modo)
        PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
        DataGrid5.Enabled = True

'            PonerBotonCabecera False
        PonerFoco txtAux1(9)
        Me.DataGrid5.Enabled = False


eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean
    
    Select Case grid
        Case "DataGrid5"
            DeseleccionaGrid Me.DataGrid5
            b = (xModo = 1)
             For jj = 3 To 10
                txtAux1(jj).Height = DataGrid5.RowHeight - 10
                txtAux1(jj).Top = alto + 5
                txtAux1(jj).visible = b
            Next jj
'            btnBuscar(1).Height = DataGrid5.RowHeight - 10
'            btnBuscar(1).Top = alto + 5
'            btnBuscar(1).visible = b
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
    
    Cad = "Cabecera de Facturas." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Factura del Transportista:            "
    Cad = Cad & vbCrLf & "Transport.:  " & Text1(2).Text
    Cad = Cad & vbCrLf & "Nº Factura:  " & Format(Text1(0).Text, "0000000")
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
    MuestraError Err.Number, "Eliminar Factura", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid5.Enabled = True
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

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim i As Integer
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    Select Case Index
        Case 0
'            Select Case Combo1(Index).ListIndex
'                Case -1
'                    CodTipoMov = ""
'                Case 0 ' Anticipo
'                    CodTipoMov = "FAA"
'                Case 1 ' Liquidacion
'                    CodTipoMov = "FAL"
'                Case 2 ' Anticipo Venta Campo
'                    CodTipoMov = "FAC"
'                Case 3 ' Liquidacion Venta Campo
'                    CodTipoMov = "FLC"
'            End Select
'            Text1(12).Text = CodTipoMov
            i = Combo1(Index).ListIndex
'            PosicionarCombo Combo1(Index), CInt(i)
            Text1(10).Text = Mid(Trim(Combo1(Index).List(i)), 1, 3)
            CodTipoMov = Text1(10).Text
        Case 1
            If (Modo = 3 Or Modo = 4) Then
                PonerCamposRet
                CalculoTotales
            End If
        Case 2
            i = Combo1(Index).ListIndex
            Text1(20).Text = Mid(Trim(Combo1(Index).List(i)), 1, 3)
    End Select

End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If PrimeraVez Then
        PonerFocoChk Me.chkVistaPrevia
        PrimeraVez = False
    End If
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then
        PonerCadenaBusqueda
    End If
    
    
'    PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 13
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(4).Image = 3   'Insertar
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(8).Image = 10  'Impresión de factura
        .Buttons(10).Image = 11  'Salir
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
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    
    CodTipoMov = "FTR"
    VieneDeBuscar = False
    Me.SSTab1.Tab = 0
        
    '## A mano
    NombreTabla = "rfacttra"
    NomTablaLineas = "rfacttra_variedad" 'Tabla de variedades de factura
    Ordenacion = " ORDER BY rfacttra.codtipom, rfacttra.numfactu, rfacttra.fecfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from rfacttra "
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmManSocios
        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
    Else
        CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    SSTab1.TabVisible(1) = False
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    End If
'    If DatosADevolverBusqueda = "" Then
'        If numalbar = "" Then
'            PonerModo 0
'        Else
'            Text1(0).Text = numalbar
'            HacerBusqueda
'        End If
'    Else
'        BotonBuscar
'    End If

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
    For i = 0 To Check1.Count - 1
        Me.Check1(i).Value = 0
    Next i
'    Label2(2).Caption = ""
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(10), CadenaDevuelta, 1)
        CadB = CadB & Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If imgFec(0).Tag < 1 Then
        Text1(CByte(imgFec(0).Tag) + 1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Else
        Text1(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")  '<===
    End If
    ' ********************************************
End Sub

' devolvemos la linea del datagrid en donde estabamos
Private Sub frmLFac_DatoSeleccionado(CadenaSeleccion As String)
Dim vWhere As String
             
   PonerCamposLineas
   
   If CadenaSeleccion = "" Then Exit Sub
             
   vWhere = "(codtipom = '" & RecuperaValor(CadenaSeleccion, 1) & "' and numfactu = " & RecuperaValor(CadenaSeleccion, 2) & " and fecfactu = " & RecuperaValor(CadenaSeleccion, 3)
   
   SituarDataMULTI Data3, vWhere, "" ', Indicador
   
   PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
   PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
   

End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'numero de albaran
    FormateaCampo txtAux1(3)
    txtAux1(4).Text = RecuperaValor(CadenaSeleccion, 2) 'fecha de albaran
    FormateaCampo txtAux1(4)
    txtAux1(5).Text = RecuperaValor(CadenaSeleccion, 3) 'variedad
    txtAux1(6).Text = RecuperaValor(CadenaSeleccion, 4) 'campo
    FormateaCampo txtAux1(6)
    txtAux1(7).Text = RecuperaValor(CadenaSeleccion, 5) 'numero de nota
    FormateaCampo txtAux1(7)
    txtAux1(8).Text = RecuperaValor(CadenaSeleccion, 6) 'kilos
    FormateaCampo txtAux1(8)
    txtAux1(9).Text = RecuperaValor(CadenaSeleccion, 7) 'importe
    FormateaCampo txtAux1(9)
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Tipo de iva
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomiva
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 3) 'porcentaje iva
    FormateaCampo Text1(4)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Variedades
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Variedad
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Variedad
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Transportista
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Transportista
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim vSeccion As CSeccion

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 1 'Tipo de IVA
            indice = 3
            PonerFoco Text1(indice)
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                If vSeccion.AbrirConta Then
                    Set frmTIva = New frmTipIVAConta
                    frmTIva.DeConsulta = True
                    frmTIva.DatosADevolverBusqueda = "0|1|2|"
                    frmTIva.CodigoActual = Text1(3).Text
                    frmTIva.Show vbModal
                    Set frmTIva = Nothing
                    PonerFoco Text1(3)
                End If
            End If
            Set vSeccion = Nothing
        
        Case 0 'Transportistas
            indice = 2
            PonerFoco Text1(indice)
            Set frmTra = New frmManTranspor
            frmTra.DatosADevolverBusqueda = "0|1|"
            frmTra.Show vbModal
            Set frmTra = Nothing
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

    Select Case Index
        Case 0
            indice = 1
        Case 1
            indice = 18
    End Select

    imgFec(0).Tag = indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text1(indice).Text <> "" Then frmC.NovaData = Text1(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(CByte(imgFec(0).Tag)) '<===
    ' ********************************************
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


Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir
End Sub



Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()

'    If Data1.Recordset!impreso = 1 Then
'        If MsgBox("Este albarán está facturado y/o cobrado. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'            Exit Sub
'        End If
'    End If

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
'    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
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
Dim vSeccion As CSeccion
Dim vSocio As cSocio

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha factura
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
    
        Case 2 'Transportistas
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rtransporte", "nomtrans")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Transportista: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra = New frmManTranspor
                        frmTra.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTra.Show vbModal
                        Set frmTra = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    ' sacamos el iva del socio
                    Set vTrans = New CTransportista
                    If vTrans.LeerDatos(Text1(2).Text) Then
                        Text1(3).Text = vTrans.CodIva
                        Combo1(1).ListIndex = vTrans.TipoIRPF
                        If Text1(3).Text <> "" Then
                            Set vSeccion = New CSeccion
                            If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                                If vSeccion.AbrirConta Then
                                    Text2(3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(3).Text, "N")
                                    Text1(4).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Text1(3).Text, "N")
                                End If
                            End If
                            Set vSeccion = Nothing
                        End If
                    End If
                    Set vTrans = Nothing
                End If
            End If
            
        Case 3 'Tipo de IVA
            If Text1(Index).Text <> "" Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                    If vSeccion.AbrirConta Then
                        Text2(3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(3).Text, "N")
                        Text1(4).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Text1(3).Text, "N")
                    End If
                End If
                Set vSeccion = Nothing
            End If
            
        
        Case 10 ' importe fondo operativo
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 1
    End Select
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

    '[Monica]09/01/2015: nuevo tipo de dato para busqueda sin asteriscos
    Text1(2).Tag = "Cod.Transportista|TT|N|||rfacttra|codtrans||S|"
    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    Text1(2).Tag = "Cod.Transportista|T|N|||rfacttra|codtrans||S|"

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select rfacttra.* from " & NombreTabla & " LEFT JOIN rfacttra_albaran ON rfacttra.codtipom=rfacttra_albaran.codtipom "
        CadenaConsulta = CadenaConsulta & " and rfacttra_albaran.numfactu = rfacttra.numfactu and rfacttra_albaran.fecfactu = rfacttra.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY rfacttra.codtipom, rfacttra.numfactu, rfacttra.fecfactu, rfacttra.codtrans " & Ordenacion
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
'    Cad = Cad & "Tipo|if(rfactsoc.codtipom='FAA','Anticipo','Liquidación') as a|T||10·"
    Cad = Cad & "Tipo Factura|stipom.nomtipom|T||22·"
    Cad = Cad & "Tipo|rfacttra.codtipom|T||7·"
    Cad = Cad & "Nº.Factura|rfacttra.numfactu|N||10·"
    Cad = Cad & "Fecha|rfacttra.fecfactu|F||15·"
    Cad = Cad & "Transportista|rfacttra.codtrans|N||15·"
    Cad = Cad & "Nombre Transportista|rtransporte.nomtrans|T||30·"
    Tabla = NombreTabla & " INNER JOIN rtransporte ON rfacttra.codtrans=rtransporte.codtrans "
    Tabla = "(" & Tabla & ") INNER JOIN usuarios.stipom stipom ON rfacttra.codtipom = stipom.codtipom "
    Titulo = "Facturas Transportistas"
    devuelve = "1|2|3|"
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vtabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
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
        LLamaLineas Modo, 0, "DataGrid5"
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
    
    CargaGrid DataGrid5, Data6, True
    
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
Dim vSeccion As CSeccion

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    b = PonerCamposForma2(Me, Data1, 2, "Frame2")

    PosicionarCombo2 Combo1(0), Text1(10).Text
    
    ' datos de la factura que rectifica si es una factura rectificativa
    If b Then b = PonerCamposForma2(Me, Data1, 2, "Frame6")
    Combo1(2).ListIndex = -1
    PosicionarCombo2 Combo1(2), Text1(15).Text
    
    
'    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    
    Text2(2).Text = PonerNombreDeCod(Text1(2), "rtransporte", "nomtrans", "codtrans", "T") 'transportistas
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        b = vSeccion.AbrirConta
        If b Then
            Text2(3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(3).Text, "N")
        End If
    End If
    Set vSeccion = Nothing
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
Dim b1 As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
    BuscaChekc = ""

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or hcoCodMovim <> "" Then
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
    
    '+++ bloqueamos el combo1(0) como si tuviera tag
    b1 = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la capçalera mentre treballe en les llínies
    
    If (Modo = 4 Or Modo = 5) Then
        Combo1(0).Enabled = False
        Combo1(0).BackColor = &H80000018 'groc
    Else
        Combo1(0).Enabled = b1
        If b1 Then
            Combo1(0).BackColor = vbWhite
        Else
            Combo1(0).BackColor = &H80000018 'Amarillo Claro
        End If
        If Modo = 3 Then Combo1(0).ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
    End If
    '+++
    
    
    For i = 4 To 14 '7
        BloquearTxt Text1(i), Not (Modo = 1)
        Text1(i).Enabled = (Modo = 1)
    Next i
    
    Text1(0).Enabled = (Modo = 1)
    
    BloquearTxt Text1(9), Not (Modo = 1)
    Text1(9).Enabled = (Modo = 1)
    BloquearTxt Text1(11), Not (Modo = 1)
    Text1(11).Enabled = (Modo = 1)
    
    For i = 0 To Check1.Count - 1
        Me.Check1(i).Enabled = (Modo = 1)
    Next i
    
    b = (Modo <> 1)
    'Campos Nº Pedido bloqueado y en azul
    BloquearTxt Text1(0), b, True
         
    BloquearTxt Text1(2), (Modo <> 1) '<= 2)
    
    For i = 0 To txtAux1.Count - 1
         BloquearTxt txtAux1(i), True
         txtAux1(i).visible = False
    Next
    
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 3 To txtAux1.Count - 1
        If i <> 5 Then
            txtAux1(i).visible = (Modo = 1)
            BloquearTxt txtAux1(i), Not (Modo = 1)
        End If
    Next i


    

'    txtAux(6).visible = False
'    txtAux(6).Enabled = True
'    For i = 0 To 7
'        BloquearTxt txtAux3(i), True
'        txtAux3(i).visible = False
'    Next i
'    For i = 3 To 8
'        If i <> 4 Then
'            BloquearTxt txtAux3(i), (Modo <> 1)
'            txtAux3(i).visible = (Modo = 1)
'        End If
'    Next i
'    BloquearChk chkAux(0), (Modo <> 1)
'    chkAux(0).visible = (Modo = 1)
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    
    imgBuscar(0).Enabled = (Modo = 1)
    imgBuscar(0).visible = (Modo = 1)
    
    BloquearImgFec Me, 0, Modo
    
    imgFec(0).Enabled = (Modo = 1 Or Modo = 3)
    imgFec(0).visible = (Modo = 1 Or Modo = 3)
    
    Text1(2).Enabled = (Modo = 1 Or Modo = 3)
    
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
'    b = (Modo = 5) And ModificaLineas = 1
'    BloquearTxt txtAux1(3), Not b
'    BloquearBtn Me.btnBuscar(0),not b
    
'    txtAux1(15).visible = False
'    txtAux1(15).Enabled = False
'
'    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
'    BloquearTxt txtAux1(16), Not b
       
       
    ' DATOS SI ES RECTIFICATIVA
    b = (Modo = 1)
    Combo1(2).Enabled = b
    If b Then
        Combo1(2).BackColor = vbWhite
    Else
        Combo1(2).BackColor = &H80000018 'Amarillo Claro
    End If
    Text1(12).Enabled = b
    Text1(13).Enabled = b
    Text1(14).Enabled = (b Or Modo = 4)
    imgFec(1).Enabled = b
    imgFec(1).visible = b
       
       
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
'    Select Case NumTabMto
'        Case 1
'            BloquearFrameAux Me, "FrameGastos", Modo, NumTabMto
'    End Select
    
'    If indFrame = 1 Then
'        txtAux(6).Enabled = (ModificaLineas = 1)
'        txtAux(6).visible = (ModificaLineas = 1)
''        btnBuscar(0).Enabled = (ModificaLineas = 1)
''        btnBuscar(0).visible = (ModificaLineas = 1)
'    End If
        
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

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scafac
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
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

    For i = 0 To txtAux1.Count - 1
        If i = 4 Or i = 6 Or i = 7 Then
            If txtAux1(i).Text = "" Then
                MsgBox "El campo " & txtAux1(i).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux1(i)
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
    
    If CInt(Data1.Recordset!impreso) = 1 Then
        If MsgBox("Esta factura ya ha sido impresa. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If
    
    If CInt(Data1.Recordset!contabilizado) = 1 Then
        If MsgBox("Esta factura ya ha sido contabilizada. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If
    
    If BloqueaRegistro(NombreTabla, "codtipom = '" & Data1.Recordset!CodTipom & "' and numfactu = " & Data1.Recordset!numfactu & " and fecfactu = " & DBSet(Data1.Recordset!fecfactu, "F")) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Index
            Case 0 'rfacttra_gastos
                Select Case Button.Index
                    Case 1
                        BotonAnyadirLinea Index
                    Case 2
                        BotonModificarLinea Index
                    Case 3
                        BotonEliminarLinea Index
                    Case Else
                End Select
'                CalcularTotales
''                CalcularGastos
'                PonerCampos
'                TerminaBloquear
                
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
        Case 0 'variedad
            
            If Data6.Recordset.RecordCount = 1 Then
                MsgBox "No podemos dejar la factura sin lineas. Elimine la factura.", vbExclamation
                Exit Sub
            End If
            
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar la nota del Albaran?"
            Cad = Cad & vbCrLf & "Tipo: " & Data6.Recordset.Fields(0)
            Cad = Cad & vbCrLf & "Factura: " & Data6.Recordset.Fields(1)
            Cad = Cad & vbCrLf & "Fecha: " & Data6.Recordset.Fields(2)
            Cad = Cad & vbCrLf & "Albarán: " & Data6.Recordset.Fields(3)
            Cad = Cad & vbCrLf & "Nota: " & Data6.Recordset!numnotac
            
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Data6.Recordset.AbsolutePosition
                
                If Not EliminarLinea(Index) Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    If SituarDataTrasEliminar(Data6, NumRegElim) Then
                        PonerCampos
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
            Screen.MousePointer = vbDefault
       
       
        Case 1 'gasto a pie de factura
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar el Gasto?"
            Cad = Cad & vbCrLf & "Tipo: " & Data5.Recordset.Fields(0)
            Cad = Cad & vbCrLf & "Factura: " & Data5.Recordset.Fields(1)
            Cad = Cad & vbCrLf & "Fecha: " & Data5.Recordset.Fields(2)
            Cad = Cad & vbCrLf & "Código: " & Data5.Recordset.Fields(4)
            
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Data5.Recordset.AbsolutePosition
                
                If Not EliminarLinea(Index) Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    If SituarDataTrasEliminar(Data5, NumRegElim) Then
'                        CalcularGastos
                        PonerCampos
                    Else
                        PonerCampos
'                    Else
'                        LimpiarCampos
'                        PonerModo 0
                    End If
                End If
            End If
            Screen.MousePointer = vbDefault
       
       
       
    End Select
       
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Linea de Factura", Err.Description

End Sub




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
        Case 10   'Salir
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
    DataGrid5.Enabled = Not b
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGRid

    b = DataGrid5.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid5" 'Albaranes de almazara y de bodega
            Opcion = 5
    End Select
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
     b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
     vDataGrid.Enabled = Not b
    
   
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGRid

    Select Case vDataGrid.Name
         Case "DataGrid5" 'rfactsoc_albaran
'           SQL = "SELECT codtipom,numfactu,fecfactu,codtipomanti,numfactuanti,fecfactuanti,codvarieanti,codvarieanti,nomvarie,codcampoanti,baseimpo "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux1(3)|T|Albaran|1000|;S|btnBuscar(1)|B|||;S|txtAux1(4)|T|Fecha|1000|;N||||0|;"
            tots = tots & "S|txtAux1(5)|T|Variedad|2100|;"
            tots = tots & "S|txtAux1(6)|T|Campo|1000|;S|txtAux1(7)|T|Nota|1000|;"
            tots = tots & "S|txtAux1(8)|T|K.Neto|1200|;"
            tots = tots & "S|txtAux1(9)|T|Importe|1200|;"
            arregla tots, DataGrid5, Me
    
            DataGrid5.Columns(3).Alignment = dbgLeft
            DataGrid5.Columns(6).Alignment = dbgLeft
    
    
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub




Private Sub txtAux1_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux1(Index)
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux1(Index), ModificaLineas) Then Exit Sub

    Select Case Index
        Case 4 ' calidad
            If txtAux1(Index) <> "" Then
                Text2(6) = DevuelveDesdeBDNew(cAgro, "rcalidad", "nomcalid", "codvarie", txtAux1(5), "N", , "codcalid", txtAux1(6).Text, "N")
                If Text2(6).Text = "" Then
                    cadMen = "No existe la Calidad: " & txtAux1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCal = New frmManCalidades
                        frmCal.DatosADevolverBusqueda = "2|3|"
                        frmCal.ParamVariedad = txtAux1(5).Text
                        frmCal.NuevoCodigo = txtAux1(6).Text
                        TerminaBloquear
                        frmCal.Show vbModal
                        Set frmCal = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        PonerFoco txtAux1(6)
                    Else
                        txtAux1(Index).Text = ""
                    End If
                    PonerFoco txtAux1(Index)
                End If
            Else
                Text2(6).Text = ""
            End If

        Case 7 'peso neto
            If txtAux1(Index) <> "" Then
                PonerFormatoEntero txtAux1(Index)
                cmdAceptar.SetFocus
            End If

    End Select
    
End Sub




Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim vTrans As CTransportista
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en tablas de cabecera de factura
    '------------------------------------------
    Sql = " " & ObtenerWhereCP(True)
    
    'Lineas de albaranes (rfacttra_albaran)
    conn.Execute "Delete from rfacttra_albaran " & Sql
    
    'Cabecera de factura (rfacttra)
    conn.Execute "Delete from " & NombreTabla & Sql
    
    'Decrementar contador si borramos el ultima factura
    
    If vParamAplic.TipoContadorTRA = 0 Then ' contador global de facturas de transporte 'FTR
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador Text1(10).Text, Val(Text1(0).Text)
        Set vTipoMov = Nothing
    Else
        Set vTrans = New CTransportista
        vTrans.DevolverContador Text1(2).Text, Val(Text1(0).Text)
        Set vTrans = Nothing
    End If
    b = True
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description & " " & Mens
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

Private Function EliminarLinea(Aux As Integer) As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    
    Select Case Aux
        Case 0
            'Eliminar en tablas de albaranes
            '------------------------------------------
            If Data6.Recordset.EOF Then Exit Function
                
            conn.BeginTrans
                
            Mens = ""
            
            
            Sql = " where codtipom = '" & Data6.Recordset.Fields(0) & "'"
            Sql = Sql & " and numfactu = " & Data6.Recordset.Fields(1)
            Sql = Sql & " and fecfactu = " & DBSet(Data6.Recordset.Fields(2), "F")
            Sql = Sql & " and numalbar = " & DBSet(Data6.Recordset.Fields(3), "N")
            Sql = Sql & " and numnotac = " & Data6.Recordset.Fields(8)
            
            'Lineas de albaranes (rfacttra_albaran)
            conn.Execute "Delete from rfacttra_albaran " & Sql
        
    End Select
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Linea de Factura ", Err.Description & " " & Mens
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

    CargaGrid DataGrid5, Data6, False  ' albaranes de bodega y de almazara
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
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
    
    Sql = " codtipom= '" & Text1(10).Text & "'"
    Sql = Sql & " and numfactu = " & Text1(0).Text
    Sql = Sql & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    Sql = Sql & " and codtrans = '" & Trim(Text1(2).Text) & "'"

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
    Case 5  ' albaranes de la factura
        Sql = "SELECT rfacttra_albaran.codtipom, rfacttra_albaran.numfactu, rfacttra_albaran.fecfactu, "
        Sql = Sql & " rfacttra_albaran.numalbar, rfacttra_albaran.fecalbar, rfacttra_albaran.codvarie, "
        Sql = Sql & " variedades.nomvarie, rfacttra_albaran.codcampo, rfacttra_albaran.numnotac, "
        Sql = Sql & " rfacttra_albaran.kilosnet, rfacttra_albaran.importe "
        Sql = Sql & " FROM rfacttra_albaran, variedades " 'lineas de albaranes de la factura socio
        Sql = Sql & " WHERE rfacttra_albaran.codvarie = variedades.codvarie "
    End Select
    
    If enlaza Then
        Sql = Sql & " and " & ObtenerWhereCP(False)
    Else
        Sql = Sql & " and numfactu = -1"
    End If
    Sql = Sql & " ORDER BY numfactu"
    If Opcion = 5 Then Sql = Sql & ", fecalbar "
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

        b = ((Modo = 2) Or (Modo = 0)) And (hcoCodMovim = "") 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(1).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(2).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(4).Enabled = b
        Me.mnModificar.Enabled = b
        
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (hcoCodMovim = "") And Not (Check1(0).Value = 1) And Not (Check1(1).Value = 1)
        'Modificar
        Toolbar1.Buttons(5).Enabled = b And vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16
        Me.mnModificar.Enabled = b And vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16
        'eliminar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnEliminar.Enabled = b
        'Impresión de albaran
        Toolbar1.Buttons(8).Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
        Me.mnImprimir.Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
        

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 2) And Not Check1(1).Value = 1
    For i = 0 To 0
        ToolAux(i).Buttons(1).Enabled = b
        
        If b Then
            Select Case i
              Case 0
                bAux = (b And Me.Data6.Recordset.RecordCount > 0)
            End Select
        End If
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
    
'    ' toolbar de gastos
'    b = (Modo = 2) Or (Modo = 5)
'    ToolAux(1).Buttons(1).Enabled = b
'
'    If b Then
'        bAux = (b And Me.Data5.Recordset.RecordCount > 0)
'    End If
'    ToolAux(1).Buttons(2).Enabled = bAux
'    ToolAux(1).Buttons(3).Enabled = bAux


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
Dim Tipo As Byte

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Factura para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de albaran
    '---------------------------------------------------
    'Tipo de factura
    devuelve = "{" & NombreTabla & ".codtipom}='" & Mid(Combo1(0).Text, 1, 3) & "'"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "codtipom = '" & Mid(Combo1(0).Text, 1, 3) & "'"
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
    Select Case Mid(Combo1(0).Text, 1, 3)
        Case "FTR"
            indRPT = 49
    End Select
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    
    
    'Nº factura
    devuelve = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "numfactu = " & Val(Text1(0).Text)
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
    'Fecha Factura
    devuelve = "{" & NombreTabla & ".fecfactu}=Date(" & Year(Text1(1).Text) & "," & Month(Text1(1).Text) & "," & Day(Text1(1).Text) & ")"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "fecfactu = " & DBSet(Text1(1).Text, "F")
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
    'Transportista
    devuelve = "{" & NombreTabla & ".codtrans}=""" & Trim(Text1(2).Text) & """"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "codtrans = '" & Trim(Text1(2).Text) & "'"
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
    
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
            .Titulo = "Impresión de Factura de Transportistas"
            .ConSubInforme = True
            .Show vbModal
    End With

    If frmVisReport.EstaImpreso Then
        ActualizarRegistros "rfacttra", cadSelect
    End If
End Sub


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 1 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    
    'tipo de IRPF
    Combo1(0).AddItem "FTR-Transporte"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
'    Combo1(0).AddItem "FRS-Rectificativa"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    
    
    'tipo de IRPF
    Combo1(1).AddItem "Módulo Agrario"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "E.D."
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Entidad"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    Combo1(1).AddItem "Módulo Transportista"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    
    'tipo de factura
    Sql = "select codtipom, nomtipom from usuarios.stipom where tipodocu > 0 and tipodocu <> 11"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 1
    While Not Rs.EOF
'        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
        Sql = Rs.Fields(1).Value
        Sql = Rs.Fields(0).Value & " - " & Sql
        Combo1(2).AddItem Sql 'campo del codigo
        Combo1(2).ItemData(Combo1(2).NewIndex) = i
        i = i + 1
        Rs.MoveNext
    Wend
    

End Sub

Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    b = ModificaDesdeFormulario2(Me, 2, "Frame2")
'    If b Then b = ModificaDesdeFormulario2(Me, 2, "Frame6")

EModificarCab:
    If Err.Number <> 0 Or Not b Then
        MenError = "Modificando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
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
Dim vTrans As CTransportista ' clase de transportista
Dim Sql As String

    On Error GoTo EInsertarCab
    
    
    If vParamAplic.TipoContadorTRA = 0 Then ' contador automatico de FTR facturas de transporte
        
        Set vTipoMov = New CTiposMov
        If vTipoMov.Leer(CodTipoMov) Then
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            Sql = CadenaInsertarDesdeForm(Me)
            If Sql <> "" Then
                If InsertarOferta(Sql, vTipoMov) Then
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    PonerModo 2
                    'Ponerse en Modo Insertar Lineas
    '                BotonMtoLineas 0, "Variedades"
    '                BotonAnyadirLinea
                    Set frmLFac = New frmManLinFactTranspor
                    
                    frmLFac.ModoExt = 3
                    frmLFac.tipoMov = Text1(10).Text
                    frmLFac.Factura = Text1(0).Text
                    frmLFac.fecfactu = Text1(1).Text
                    
                    frmLFac.Show vbModal
                    
                    Set frmLFac = Nothing
                    
                    CalcularTotales
                End If
            End If
            Text1(0).Text = Format(Text1(0).Text, "0000000")
        End If
        Set vTipoMov = Nothing
        
    Else
        Set vTrans = New CTransportista
        If vTrans.LeerDatos(Text1(2).Text) Then
            Text1(0).Text = vTrans.ConseguirContador()
            Sql = CadenaInsertarDesdeForm(Me)
            If Sql <> "" Then
                If InsertarOferta2(Sql, vTrans) Then
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    PonerModo 2
                    'Ponerse en Modo Insertar Lineas
    '                BotonMtoLineas 0, "Variedades"
    '                BotonAnyadirLinea
                    Set frmLFac = New frmManLinFactTranspor
                    
                    frmLFac.ModoExt = 3
                    frmLFac.tipoMov = Text1(10).Text
                    frmLFac.Factura = Text1(0).Text
                    frmLFac.fecfactu = Text1(1).Text
                    
                    frmLFac.Show vbModal
                    
                    Set frmLFac = Nothing
                    
                    CalcularTotales
                End If
            End If
            Text1(0).Text = Format(Text1(0).Text, "0000000")
        End If
        
        
        Set vTrans = Nothing
    
    End If
    
    
    
    
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
    'Comprobar si mientras tanto se incremento el contador de Factura
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numfactu", "numfactu", Text1(0).Text, "N", , "codtipom", Text1(12).Text, "T")
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
    MenError = "Error al insertar en la tabla Cabecera de Factura (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador de la Factura."
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


Private Function InsertarOferta2(vSQL As String, vTrans As CTransportista) As Boolean
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
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numfactu", "numfactu", Text1(0).Text, "N", , "codtipom", Text1(12).Text, "T")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTrans.IncrementarContador
            Text1(0).Text = vTrans.ConseguirContador()
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
    
    MenError = "Error al actualizar el contador de la Factura."
    vTrans.IncrementarContador
    
EInsertarOferta:
    If Err.Number <> 0 Then
        MenError = "Insertando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        InsertarOferta2 = True
    Else
        conn.RollbackTrans
        InsertarOferta2 = False
    End If
End Function




Private Sub InsertarLinea(Index As Integer)
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case Index
        Case 0: nomframe = "FrameAlbaranes" 'albaranes
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
'yo                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtaux(2).Text)
'                End If
            ' *************************************************
            b = BloqueaRegistro("rfacttra", "numfactu = " & Data1.Recordset!numfactu)
            CargaGrid DataGrid5, Data6, True
            
'            CalcularGastos
            PonerCampos
            
            If b Then
                BotonAnyadirLinea NumTabMto
                LLamaLineas 1, 0
            End If
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
       
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    vtabla = "rfacttra_albaran"
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case NumTabMto
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
'             *** canviar la clau primaria de les llínies,
'            pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
'             ***************************************************************

            AnyadirLinea DataGrid5, Data6

            anc = DataGrid5.Top
            If DataGrid5.Row < 0 Then
                anc = anc + 220
            Else
                anc = anc + DataGrid5.RowTop(DataGrid5.Row) + 5
            End If

            LLamaLineas ModificaLineas, anc, "DataGrid5"

            LimpiarCamposLin "FrameAlbaranes"

            txtAux1(0).Text = Mid(Combo1(0).Text, 1, 3) 'tipo de movimiento
            txtAux1(1).Text = Text1(0).Text 'nro de factura
            txtAux1(2).Text = Text1(1).Text 'fecha de factura
            
'            BloquearTxt txtAux(14), False
'                    BloquearTxt txtaux(12), False
            PonerFoco txtAux1(3)


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
    nomframe = "FrameGastos" 'gastos
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            ' ******************************************************
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModificaLineas = 0
            
'            CalcularGastos
            PonerCampos
  
            V = Data5.Recordset.Fields(3) 'el 2 es el nº de llinia
            CargaGrid DataGrid5, Data5, True

            ' *** si n'hi han tabs ***
'                    SSTab1.Tab = 1

            DataGrid5.SetFocus
            Data5.Recordset.Find (Data5.Recordset.Fields(3).Name & " =" & V)

            LLamaLineas ModificaLineas, 0, "DataGrid5"
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

    DatosOkLlin = True
'
'    On Error GoTo EDatosOKLlin
'
'    Mens = ""
'    DatosOkLlin = False
'
'    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
'    If Not b Then Exit Function
'
'    Select Case nomframe
'        Case "FrameGastos"
'            If txtAux3(16).Text = "" Then
'                MsgBox "Debe introducir un importe. Reintroduzca.", vbExclamation
'                PonerFoco txtAux3(16)
'            End If
'
'    End Select
'
''    ' ******************************************************************************
'    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codtipom = '" & Mid(Combo1(0).Text, 1, 3) & "' and numfactu = " & Val(Text1(0).Text) & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function


Private Sub CalcularTotales()
Dim Importe  As Currency
Dim Rs As ADODB.Recordset
Dim Sql As String


    Sql = "select sum(importe) from rfacttra_albaran where codtipom = " & DBSet(Data1.Recordset.Fields(0).Value, "T")
    Sql = Sql & " and numfactu = " & Data1.Recordset.Fields(1).Value
    Sql = Sql & " and fecfactu = " & DBSet(Data1.Recordset.Fields(2).Value, "F")
    Sql = Sql & " and codtrans = " & DBSet(Data1.Recordset!codTrans, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Importe = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Importe = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    Text1(5).Text = Format(Importe, "###,##0.00")
    
    CalculoTotales
    If Modo <> 4 Then
        BotonModificar
        cmdAceptar_Click
    End If
End Sub


Private Sub CalculoTotales()
Dim Base As Currency
Dim Tiva As Currency
Dim PorIva As Currency
Dim impiva As Currency
Dim BaseReten As Currency
Dim BaseAFO As Currency
Dim PorRet As Currency
Dim ImpRet As Currency
Dim PorAFO As Currency
Dim ImpAFO As Currency
Dim TotFac As Currency

    
    Base = CCur(ComprobarCero(Text1(5).Text))
    PorIva = CCur(ComprobarCero(Text1(4).Text))
    impiva = Round2(Base * PorIva / 100, 2)
    
    Select Case Combo1(1).ListIndex
        Case 0
            BaseReten = Base + impiva
        Case 1
            BaseReten = Base
        Case 2
            BaseReten = 0
        Case 3
            BaseReten = Base
    End Select
    
    'solo en el caso de que estemos insertando y modificando y no haya % de retencion
    'le daremos el que hay en parametros
    If Text1(8).Text = "" And Combo1(1).ListIndex <> 2 And (Modo = 3 Or Modo = 4) Then
        Select Case Combo1(1).ListIndex = 1
            Case 0, 1
                Text1(8).Text = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
            Case 3
                Text1(8).Text = CCur(ComprobarCero(vParamAplic.PorcreteFacTra))
        End Select
    End If
    
    ' calculo de la retencion
    PorRet = CCur(ComprobarCero(Text1(8).Text))
    ImpRet = Round2(BaseReten * PorRet / 100, 2)
    
    TotFac = Base + impiva - ImpRet
    
    If impiva = 0 Then
        Text1(6).Text = "0"
    Else
        Text1(6).Text = Format(impiva, "###,##0.00")
    End If
    
    If BaseReten = 0 Then
        Text1(11).Text = ""
    Else
        Text1(11).Text = Format(BaseReten, "###,##0.00")
    End If
    
    If ImpRet = 0 Then
        Text1(9).Text = ""
    Else
        Text1(9).Text = ImpRet
    End If
    
    If TotFac = 0 Then
        Text1(7).Text = "0"
    Else
        Text1(7).Text = Format(TotFac, "###,##0.00")
    End If
End Sub


Private Sub PonerCamposRet()
Dim i As Integer
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub
    
    For i = 9 To 11
        If i <> 10 Then
            Text1(i).Enabled = (Combo1(0).ListIndex <> 2)
            If (Combo1(1).ListIndex = 2) Then
                 Text1(i).Text = ""
            End If
        End If
    Next i

End Sub

