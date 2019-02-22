VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManFactSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Socios"
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   13950
   Icon            =   "frmManFactSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   12690
      TabIndex        =   25
      Top             =   10350
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   45
      TabIndex        =   157
      Top             =   0
      Width           =   3495
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   158
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
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
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
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
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3570
      TabIndex        =   155
      Top             =   0
      Width           =   1740
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   165
         TabIndex        =   156
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Recepción Facturas"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generación Entradas"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anticipos sin Entradas"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5340
      TabIndex        =   153
      Top             =   0
      Width           =   2370
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   154
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10800
      TabIndex        =   152
      Top             =   180
      Width           =   1605
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4425
      Left            =   90
      TabIndex        =   52
      Top             =   5835
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   7805
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Variedad/Calidad"
      TabPicture(0)   =   "frmManFactSocios.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameAnticipos"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Gastos a Pie"
      TabPicture(1)   =   "frmManFactSocios.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameGastos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Albaranes"
      TabPicture(2)   =   "frmManFactSocios.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "FrameAlbaranes"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Rectificativa"
      TabPicture(3)   =   "frmManFactSocios.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Facturas Varias"
      TabPicture(4)   =   "frmManFactSocios.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameFVarias"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   2085
         Left            =   -74730
         TabIndex        =   121
         Top             =   780
         Width           =   9835
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
            Height          =   915
            Index           =   19
            Left            =   30
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   128
            Tag             =   "Motivo Rectif. Factura|T|S|||rfactsoc|rectif_motivo|||"
            Text            =   "frmManFactSocios.frx":0098
            Top             =   990
            Width           =   8775
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
            Index           =   18
            Left            =   6855
            MaxLength       =   10
            TabIndex        =   124
            Tag             =   "Fecha Factura Rectificada|F|S|||rfactsoc|rectif_fecfactu|dd/mm/yyyy|N|"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
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
            Left            =   5580
            MaxLength       =   7
            TabIndex        =   123
            Tag             =   "Nº Factura Rectificada|N|S|||rfactsoc|rectif_numfactu|0000000|S|"
            Text            =   "Text1"
            Top             =   240
            Width           =   1110
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
            Left            =   30
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   240
            Width           =   5310
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   130
            Tag             =   "Tipo Movimiento Fact.Rectificada|T|S|||rfactsoc|rectif_codtipom||N|"
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Motivo"
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
            Left            =   30
            TabIndex        =   129
            Top             =   705
            Width           =   780
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   7965
            Picture         =   "frmManFactSocios.frx":00FD
            ToolTipText     =   "Buscar fecha"
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Factura"
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
            Index           =   13
            Left            =   5580
            TabIndex        =   127
            Top             =   -45
            Width           =   1170
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Fac"
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
            Left            =   6855
            TabIndex        =   126
            Top             =   -45
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Factura"
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
            Index           =   11
            Left            =   0
            TabIndex        =   125
            Top             =   -45
            Width           =   945
         End
      End
      Begin VB.Frame FrameFVarias 
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
         Left            =   -74955
         TabIndex        =   134
         Top             =   360
         Width           =   13390
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   7
            Left            =   5340
            MaxLength       =   30
            TabIndex        =   146
            Tag             =   "Nomforpa|N|N|||rfactsoc_fvarias|nomforpa|||"
            Text            =   "Nomforpa "
            Top             =   1170
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   6
            Left            =   4350
            MaxLength       =   30
            TabIndex        =   145
            Tag             =   "Forma Pago|N|N|||rfactsoc_fvarias|codforpa|000||"
            Text            =   "Forma de pago"
            Top             =   1170
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   2
            Left            =   1485
            MaskColor       =   &H00000000&
            TabIndex        =   142
            ToolTipText     =   "Buscar concepto gasto"
            Top             =   45
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   8
            Left            =   6780
            MaxLength       =   30
            TabIndex        =   141
            Tag             =   "Importe|N|N|||rfactsoc_fvarias|importe|###,##0.00||"
            Text            =   "Importe"
            Top             =   1140
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   3
            Left            =   2295
            MaxLength       =   7
            TabIndex        =   140
            Tag             =   "Tipo Mov FVar|T|N|||rfactsoc_fvarias|codtipomfvar||S|"
            Text            =   "Tipom"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   4
            Left            =   2970
            MaxLength       =   10
            TabIndex        =   139
            Tag             =   "Numfactu FVar|N|N|||rfactsoc_fvarias|numfactufvar|0000000|S|"
            Text            =   "Numfact"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   0
            Left            =   315
            MaxLength       =   7
            TabIndex        =   138
            Tag             =   "Tipo Movim.|T|N|||rfactsoc_fvarias|codtipom||S|"
            Text            =   "tipof"
            Top             =   1170
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   1
            Left            =   1035
            MaxLength       =   7
            TabIndex        =   137
            Tag             =   "Nº.Factura|N|N|||rfactsoc_fvarias|numfactu|0000000|S|"
            Text            =   "numfact"
            Top             =   1170
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   2
            Left            =   1665
            MaxLength       =   10
            TabIndex        =   136
            Tag             =   "Fecha Fact.|F|N|||rfactsoc_fvarias|fecfactu|dd/mm/yyyy|S|"
            Text            =   "FecFact"
            Top             =   1155
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Left            =   3645
            MaxLength       =   25
            TabIndex        =   135
            Tag             =   "Fecfactu Fvar|F|N|||rfactsoc_fvarias|fecfactufvar|dd/mm/yyyy||"
            Text            =   "fecfactu"
            Top             =   1170
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   3
            Left            =   180
            TabIndex        =   143
            Top             =   0
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
         Begin MSDataGridLib.DataGrid DataGrid6 
            Bindings        =   "frmManFactSocios.frx":0188
            Height          =   3120
            Left            =   180
            TabIndex        =   144
            Top             =   450
            Width           =   13120
            _ExtentX        =   23151
            _ExtentY        =   5503
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
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
         Height          =   2190
         Left            =   -74955
         TabIndex        =   80
         Top             =   360
         Width           =   13390
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   6
            Left            =   5895
            MaxLength       =   7
            TabIndex        =   89
            Tag             =   "Precio Medio|N|N|||rfactsoc_variedad|preciomed|#0.0000||"
            Text            =   "precmed"
            Top             =   1125
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   7
            Left            =   6480
            MaxLength       =   30
            TabIndex        =   91
            Tag             =   "Imp.Variedad|N|S|||rfactsoc_variedad|imporvar|###,##0.00||"
            Text            =   "Imp.Varie"
            Top             =   1125
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Left            =   5310
            MaxLength       =   7
            TabIndex        =   88
            Tag             =   "Kilos Netos|N|N|||rfactsoc_variedad|kilosnet|###,##0|N|"
            Text            =   "kilosne"
            Top             =   1140
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   3
            Left            =   2430
            MaxLength       =   7
            TabIndex        =   86
            Tag             =   "Variedad|N|N|||rfactsoc_variedad|codvarie|000000|S|"
            Text            =   "varieda"
            Top             =   1125
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   0
            Left            =   495
            MaxLength       =   7
            TabIndex        =   85
            Tag             =   "Tipo Movim.|T|N|||rfactsoc_variedad|codtipom||S|"
            Text            =   "tipof"
            Top             =   1155
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   1
            Left            =   1170
            MaxLength       =   7
            TabIndex        =   84
            Tag             =   "Nº.Factura|N|N|||rfactsoc_variedad|numfactu|0000000|S|"
            Text            =   "numfact"
            Top             =   1155
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   2
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   83
            Tag             =   "Fecha Fact.|F|N|||rfactsoc_variedad|fecfactu|dd/mm/yyyy|S|"
            Text            =   "FecFact"
            Top             =   1155
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   4
            Left            =   3090
            MaxLength       =   25
            TabIndex        =   82
            Text            =   "nomvari"
            Top             =   1140
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   8
            Left            =   4710
            MaxLength       =   30
            TabIndex        =   87
            Tag             =   "Campo|N|N|||rfactsoc_variedad|codcampo|00000000|S|"
            Text            =   "Campo"
            Top             =   1140
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   9
            Left            =   7110
            MaxLength       =   25
            TabIndex        =   93
            Tag             =   "Imp.Gasto|N|S|||rfactsoc_variedad|imporgasto|###,##0.00||"
            Text            =   "Imp.Gasto"
            Top             =   1125
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CheckBox chkAux 
            Height          =   195
            Index           =   0
            Left            =   8400
            TabIndex        =   81
            Tag             =   "Descontado Liq.|N|N|0|1|rfactsoc_variedad|descontado|0||"
            Top             =   1200
            Visible         =   0   'False
            Width           =   225
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   180
            TabIndex        =   90
            Top             =   0
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
            Bindings        =   "frmManFactSocios.frx":019D
            Height          =   1680
            Left            =   180
            TabIndex        =   92
            Top             =   420
            Width           =   13125
            _ExtentX        =   23151
            _ExtentY        =   2963
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
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
      Begin VB.Frame FrameGastos 
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
         Height          =   3690
         Left            =   -74865
         TabIndex        =   69
         Top             =   360
         Width           =   13255
         Begin VB.TextBox txtAux3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   15
            Left            =   4095
            MaxLength       =   25
            TabIndex        =   76
            Text            =   "nomconcepto"
            Top             =   1125
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   12
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   75
            Tag             =   "Fecha Fact.|F|N|||rfactsoc_gastos|fecfactu|dd/mm/yyyy|S|"
            Text            =   "FecFact"
            Top             =   1155
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   11
            Left            =   1170
            MaxLength       =   7
            TabIndex        =   74
            Tag             =   "Nº.Factura|N|N|||rfactsoc_gastos|numfactu|0000000|S|"
            Text            =   "numfact"
            Top             =   1155
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   10
            Left            =   495
            MaxLength       =   7
            TabIndex        =   73
            Tag             =   "Tipo Movim.|T|N|||rfactsoc_gastos|codtipom||S|"
            Text            =   "tipof"
            Top             =   1155
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   14
            Left            =   3195
            MaxLength       =   7
            TabIndex        =   72
            Tag             =   "Gasto|N|N|||rfactsoc_gastos|codgasto|000|N|"
            Text            =   "gasto"
            Top             =   1140
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   16
            Left            =   5805
            MaxLength       =   30
            TabIndex        =   79
            Tag             =   "Imp.Gastos|N|N|||rfactsoc_gastos|importe|###,##0.00||"
            Text            =   "Imp.Gasto"
            Top             =   1125
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   3870
            MaskColor       =   &H00000000&
            TabIndex        =   71
            ToolTipText     =   "Buscar concepto gasto"
            Top             =   1170
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Left            =   2430
            MaxLength       =   3
            TabIndex        =   70
            Tag             =   "Linea|N|N|||rfactsoc_gastos|numlinea|00000|S|"
            Text            =   "linea"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   90
            TabIndex        =   77
            Top             =   0
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
            Bindings        =   "frmManFactSocios.frx":01B2
            Height          =   3120
            Left            =   90
            TabIndex        =   78
            Top             =   450
            Width           =   13125
            _ExtentX        =   23151
            _ExtentY        =   5503
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
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
         Left            =   45
         TabIndex        =   53
         Top             =   360
         Width           =   13390
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Left            =   3645
            MaxLength       =   25
            TabIndex        =   66
            Tag             =   "Variedad|N|N|||rfactsoc_albaran|codvarie|000000||"
            Text            =   "variedad"
            Top             =   1170
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   2
            Left            =   1665
            MaxLength       =   10
            TabIndex        =   65
            Tag             =   "Fecha Fact.|F|N|||rfactsoc_albaran|fecfactu|dd/mm/yyyy|S|"
            Text            =   "FecFact"
            Top             =   1155
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   1
            Left            =   1035
            MaxLength       =   7
            TabIndex        =   64
            Tag             =   "Nº.Factura|N|N|||rfactsoc_albaran|numfactu|0000000|S|"
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
            Index           =   0
            Left            =   315
            MaxLength       =   7
            TabIndex        =   63
            Tag             =   "Tipo Movim.|T|N|||rfactsoc_albaran|codtipom||S|"
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
            Index           =   4
            Left            =   2970
            MaxLength       =   10
            TabIndex        =   62
            Tag             =   "Fecha Alb|F|N|||rfactsoc_albaran|fecalbar|dd/mm/yyyy|S|"
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
            Index           =   11
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   61
            Tag             =   "Imp.Gastos|N|N|||rfactsoc_albaran|imporgasto|###,##0.00||"
            Text            =   "Imp.Gasto"
            Top             =   1170
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   3
            Left            =   2295
            MaxLength       =   7
            TabIndex        =   60
            Tag             =   "Albaran|N|N|||rfactsoc_albaran|numalbar|0000000|S|"
            Text            =   "albaran"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   6
            Left            =   4365
            MaxLength       =   25
            TabIndex        =   59
            Tag             =   "Campo|N|N|||rfactsoc_albaran|codcampo|00000000|N|"
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
            Index           =   7
            Left            =   5040
            MaxLength       =   30
            TabIndex        =   58
            Tag             =   "Kilos Netos|N|N|||rfactsoc_albaran|kilosnet|###,##0||"
            Text            =   "K.Neto"
            Top             =   1170
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   8
            Left            =   5805
            MaxLength       =   30
            TabIndex        =   57
            Tag             =   "Grado|N|N|||rfactsoc_albaran|grado|##0.00||"
            Text            =   "Grado"
            Top             =   1170
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   9
            Left            =   6345
            MaxLength       =   30
            TabIndex        =   56
            Tag             =   "Precio|N|N|||rfactsoc_albaran|precio|#,##0.0000||"
            Text            =   "Precio"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   10
            Left            =   7065
            MaxLength       =   30
            TabIndex        =   55
            Tag             =   "Importe|N|N|||rfactsoc_albaran|importe|###,##0.00||"
            Text            =   "Importe"
            Top             =   1170
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   1
            Left            =   1485
            MaskColor       =   &H00000000&
            TabIndex        =   54
            ToolTipText     =   "Buscar concepto gasto"
            Top             =   45
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   2
            Left            =   180
            TabIndex        =   67
            Top             =   0
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
                  Object.ToolTipText     =   "Insertar Albaranes"
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
            Bindings        =   "frmManFactSocios.frx":01C7
            Height          =   3120
            Left            =   180
            TabIndex        =   68
            Top             =   450
            Width           =   13120
            _ExtentX        =   23151
            _ExtentY        =   5503
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
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
         Caption         =   "Calidades"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1755
         Left            =   -74910
         TabIndex        =   94
         Top             =   2520
         Width           =   13345
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   10
            Left            =   5580
            MaxLength       =   12
            TabIndex        =   105
            Tag             =   "Precio aplicado|N|N|||rfactsoc_calidad|preciocalidad|#0.0000|N|"
            Text            =   "precio"
            Top             =   1260
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   9
            Left            =   2370
            MaxLength       =   12
            TabIndex        =   104
            Tag             =   "Campo|N|N|||rfactsoc_calidad|codcampo|00000000|S|"
            Text            =   "campo"
            Top             =   1230
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   8
            Left            =   4920
            MaxLength       =   12
            TabIndex        =   103
            Tag             =   "Importe|N|N|||rfactsoc_calidad|imporcal|###,##0.00|N|"
            Text            =   "importe"
            Top             =   1260
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   7
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   102
            Tag             =   "Precio|N|N|||rfactsoc_calidad|precio|#0.0000|N|"
            Text            =   "precio"
            Top             =   1260
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   6
            Left            =   3990
            MaxLength       =   12
            TabIndex        =   101
            Tag             =   "Kilos Netos|N|N|||rfactsoc_calidad|kilosnet|###,##0|N|"
            Text            =   "k.net"
            Top             =   1260
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Left            =   3420
            MaxLength       =   12
            TabIndex        =   100
            Text            =   "nomcalid"
            Top             =   1260
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   4
            Left            =   2850
            MaxLength       =   12
            TabIndex        =   99
            Tag             =   "Calidad|N|N|||rfactsoc_calidad|codcalid|0000|S|"
            Text            =   "calidad"
            Top             =   1260
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   3
            Left            =   1920
            MaxLength       =   12
            TabIndex        =   98
            Tag             =   "Variedad|N|N|||rfactsoc_calidad|codvarie|000000|N|"
            Text            =   "variedad"
            Top             =   1230
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   2
            Left            =   1470
            MaxLength       =   4
            TabIndex        =   97
            Tag             =   "Fec.Factura|F|N|||rfactsoc_calidad|fecfactu|dd/mm/yyyy|N|"
            Text            =   "fecf"
            Top             =   1230
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   1
            Left            =   900
            MaxLength       =   7
            TabIndex        =   96
            Tag             =   "Num.Fact|N|N|||rfactsoc_calidad|numfactu|0000000|S|"
            Text            =   "numfact"
            Top             =   1230
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   0
            Left            =   480
            MaxLength       =   7
            TabIndex        =   95
            Tag             =   "Tipo Movimiento|T|N|||rfactsoc_calidad|codtipom||S|"
            Text            =   "tipofac"
            Top             =   1230
            Visible         =   0   'False
            Width           =   405
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmManFactSocios.frx":01DC
            Height          =   1395
            Left            =   150
            TabIndex        =   106
            Top             =   270
            Width           =   13105
            _ExtentX        =   23125
            _ExtentY        =   2461
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
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
      Begin VB.Frame FrameAnticipos 
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
         ForeColor       =   &H00C00000&
         Height          =   1755
         Left            =   -74910
         TabIndex        =   107
         Top             =   2520
         Visible         =   0   'False
         Width           =   13365
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   8
            Left            =   5685
            MaxLength       =   12
            TabIndex        =   117
            Tag             =   "Campo Anti|N|N|||ranticipos|codcampoanti|00000000|S|"
            Text            =   "campo"
            Top             =   1140
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   9
            Left            =   6195
            MaxLength       =   12
            TabIndex        =   116
            Tag             =   "Importe|N|N|||ranticipos|baseimpo|###,##0.00|N|"
            Text            =   "importe"
            Top             =   1140
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   6
            Left            =   3585
            MaxLength       =   12
            TabIndex        =   115
            Tag             =   "Variedad Anti|N|N|||ranticipos|codvarieanti|000000|S|"
            Text            =   "variedad"
            Top             =   1140
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   2
            Left            =   1515
            MaxLength       =   4
            TabIndex        =   114
            Tag             =   "Fec.Factura|F|N|||ranticipos|fecfactu|dd/mm/yyyy|S|"
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
            Index           =   1
            Left            =   945
            MaxLength       =   7
            TabIndex        =   113
            Tag             =   "Num.Fact|N|N|||ranticipos|numfactu|0000000|S|"
            Text            =   "numfact"
            Top             =   1140
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   0
            Left            =   495
            MaxLength       =   7
            TabIndex        =   112
            Tag             =   "Tipo Movimiento|T|N|||ranticipos|codtipom||S|"
            Text            =   "tipofac"
            Top             =   1140
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Left            =   3105
            MaxLength       =   4
            TabIndex        =   111
            Tag             =   "Fec.Factura Anti|F|N|||ranticipos|fecfactuanti|dd/mm/yyyy|S|"
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
            Index           =   4
            Left            =   2445
            MaxLength       =   7
            TabIndex        =   110
            Tag             =   "Num.Fact Anti|N|N|||ranticipos|numfactuanti|0000000|S|"
            Text            =   "numfact"
            Top             =   1140
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   3
            Left            =   1965
            MaxLength       =   7
            TabIndex        =   109
            Tag             =   "Tipo Movimiento Anti|T|N|||ranticipos|codtipomanti||S|"
            Text            =   "tipofac"
            Top             =   1140
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Index           =   7
            Left            =   4035
            MaxLength       =   25
            TabIndex        =   108
            Text            =   "nomvari"
            Top             =   1140
            Visible         =   0   'False
            Width           =   1560
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "frmManFactSocios.frx":01F1
            Height          =   1395
            Left            =   150
            TabIndex        =   118
            Top             =   270
            Width           =   13105
            _ExtentX        =   23125
            _ExtentY        =   2461
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
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
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   4875
      Left            =   60
      TabIndex        =   28
      Top             =   795
      Width           =   13650
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
         Height          =   600
         Index           =   23
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   151
         Tag             =   "Observaciones|T|S|||rfactsoc|observaciones||N|"
         Text            =   "frmManFactSocios.frx":0206
         Top             =   1830
         Width           =   13350
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
         Left            =   12600
         MaxLength       =   6
         TabIndex        =   149
         Tag             =   "Porc.Corredor|N|S|||rfactsoc|porccorredor|##0.00||"
         Text            =   "123456"
         Top             =   1080
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Caption         =   "FacturaE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   11370
         TabIndex        =   147
         Tag             =   "FacturaE|N|N|0|1|rfactsoc|enfacturae|0||"
         Top             =   435
         Width           =   1410
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Complementaria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   9480
         TabIndex        =   133
         Tag             =   "Es Liq.Complementaria|N|N|0|1|rfactsoc|esliqcomplem|0||"
         Top             =   1470
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   7455
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Nº Factura Rec|T|S|||rfactsoc|numfacrec||N|"
         Text            =   "Text1"
         Top             =   360
         Width           =   1665
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
         Left            =   7455
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "Pdte.Recibir Nro.Fac|N|N|0|2|rfactsoc|pdtenrofact|0||"
         Top             =   1080
         Width           =   1620
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Anticipo Retirada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   9480
         TabIndex        =   11
         Tag             =   "Es Anticipo de Retirada|N|N|0|1|rfactsoc|esretirada|0||"
         Top             =   1125
         Width           =   2085
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
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   4410
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   960
         MaxLength       =   10
         TabIndex        =   120
         Tag             =   "Tipo Movimiento|T|N|||rfactsoc|codtipom||S|"
         Top             =   360
         Width           =   1065
      End
      Begin VB.PictureBox cmdAnticipos 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   12915
         ScaleHeight     =   495
         ScaleWidth      =   525
         TabIndex        =   119
         Top             =   135
         Width           =   525
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Anticipo Gasto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   9480
         TabIndex        =   10
         Tag             =   "Es Anticipo de Gasto|N|N|0|1|rfactsoc|esanticipogasto|0||"
         Top             =   780
         Width           =   1860
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
         Index           =   14
         Left            =   240
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Base Aportación|N|S|||rfactsoc|baseaport|###,##0.00||"
         Top             =   4215
         Width           =   1575
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
         Left            =   5895
         MaxLength       =   6
         TabIndex        =   20
         Tag             =   "Porc.AFO|N|S|||rfactsoc|porc_apo|##0.00||"
         Text            =   "123"
         Top             =   4215
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pasa Aridoc"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   11370
         TabIndex        =   9
         Tag             =   "Aridoc|N|N|0|1|rfactsoc|pasaridoc|0||"
         Top             =   90
         Width           =   1500
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
         Index           =   5
         Left            =   240
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Base Imponible|N|N|||rfactsoc|baseimpo|###,##0.00||"
         Top             =   3015
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CAE3FD&
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
         Left            =   10305
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "Total Factura|N|N|||rfactsoc|totalfac|###,##0.00||"
         Top             =   3015
         Width           =   2325
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
         Index           =   6
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Importe Iva|N|N|||rfactsoc|imporiva|###,##0.00||"
         Top             =   3015
         Width           =   1875
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
         Index           =   3
         Left            =   1890
         MaxLength       =   2
         TabIndex        =   13
         Tag             =   "Tipo Iva|N|N|0|99|rfactsoc|tipoiva|00||"
         Text            =   "Te"
         Top             =   3015
         Width           =   600
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
         Index           =   3
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   3015
         Width           =   3135
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
         Left            =   5895
         MaxLength       =   6
         TabIndex        =   14
         Tag             =   "Porc.Iva|N|N|||rfactsoc|porc_iva|##0.00||"
         Text            =   "123"
         Top             =   3015
         Width           =   870
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
         Index           =   8
         Left            =   5895
         MaxLength       =   6
         TabIndex        =   17
         Tag             =   "Porc.Retención|N|S|||rfactsoc|porc_ret|##0.00||"
         Text            =   "123"
         Top             =   3630
         Width           =   870
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
         Index           =   9
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Importe Retención|N|S|||rfactsoc|impreten|#,##0.00||"
         Text            =   "123"
         Top             =   3630
         Width           =   1875
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Base Retención|N|S|||rfactsoc|basereten|###,##0.00||"
         Top             =   3630
         Width           =   1575
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
         Left            =   5895
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo IRPF|N|N|0|3|rfactsoc|tipoirpf|0|N|"
         Top             =   1080
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
         Index           =   10
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Importe Aportización FO|N|S|||rfactsoc|impapor|###,##0.00||"
         Text            =   "123"
         Top             =   4215
         Width           =   1845
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   9480
         TabIndex        =   8
         Tag             =   "Contabilizado|N|N|0|1|rfactsoc|contabilizado|0||"
         Top             =   420
         Width           =   2250
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
         Index           =   2
         Left            =   120
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Cod.Socio|N|N|0|999999|rfactsoc|codsocio|000000|N|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   4620
         MaxLength       =   7
         TabIndex        =   1
         Tag             =   "Nº Factura|N|S|||rfactsoc|numfactu|0000000|S|"
         Text            =   "Text1"
         Top             =   360
         Width           =   1020
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Impreso"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   9480
         TabIndex        =   6
         Tag             =   "Impreso|N|N|0|1|rfactsoc|impreso|0||"
         Top             =   75
         Width           =   1860
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
         Index           =   1
         Left            =   5895
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||rfactsoc|fecfactu|dd/mm/yyyy|S|"
         Top             =   360
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
         Index           =   2
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Frame Frame5 
         Caption         =   "Total Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2295
         Left            =   60
         TabIndex        =   30
         Top             =   2490
         Width           =   13435
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00CAE3FD&
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
            Index           =   16
            Left            =   10260
            MaxLength       =   10
            TabIndex        =   51
            Top             =   1710
            Width           =   2325
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00CAE3FD&
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
            Index           =   15
            Left            =   10260
            MaxLength       =   10
            TabIndex        =   49
            Top             =   1125
            Width           =   2325
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL A COBRAR"
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
            Index           =   6
            Left            =   10260
            TabIndex        =   50
            Top             =   1485
            Width           =   1995
         End
         Begin VB.Label Label1 
            Caption         =   "Total Gastos a Pie"
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
            Index           =   5
            Left            =   10260
            TabIndex        =   48
            Top             =   900
            Width           =   2085
         End
         Begin VB.Label Label4 
            Caption         =   "Importe A.F.O"
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
            Left            =   7380
            TabIndex        =   47
            Top             =   1485
            Width           =   1725
         End
         Begin VB.Label Label1 
            Caption         =   "Base AFO"
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
            Index           =   4
            Left            =   180
            TabIndex        =   46
            Top             =   1470
            Width           =   1590
         End
         Begin VB.Label Label3 
            Caption         =   "% AFO"
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
            Left            =   5835
            TabIndex        =   45
            Top             =   1485
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
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
            Left            =   180
            TabIndex        =   44
            Top             =   270
            Width           =   1590
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL FACTURA"
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
            Index           =   9
            Left            =   10245
            TabIndex        =   43
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
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
            Left            =   7380
            TabIndex        =   42
            Top             =   270
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   2220
            ToolTipText     =   "Buscar Iva"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Iva"
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
            Left            =   1845
            TabIndex        =   41
            Top             =   270
            Width           =   330
         End
         Begin VB.Label Label2 
            Caption         =   "% Iva"
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
            Left            =   5835
            TabIndex        =   40
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label18 
            Caption         =   "% Ret"
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
            Left            =   5835
            TabIndex        =   39
            Top             =   900
            Width           =   720
         End
         Begin VB.Label Label6 
            Caption         =   "Importe Retención"
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
            Left            =   7380
            TabIndex        =   38
            Top             =   900
            Width           =   1890
         End
         Begin VB.Label Label1 
            Caption         =   "Base Retención"
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
            Index           =   2
            Left            =   180
            TabIndex        =   37
            Top             =   885
            Width           =   1590
         End
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1695
         Tag             =   "-1"
         ToolTipText     =   "Zoom descripción"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
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
         Index           =   17
         Left            =   150
         TabIndex        =   150
         Top             =   1545
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "% Corred."
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
         Left            =   11610
         TabIndex        =   148
         Top             =   1125
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Factura Recibida"
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
         Left            =   7455
         TabIndex        =   132
         Top             =   75
         Width           =   1950
      End
      Begin VB.Label Label1 
         Caption         =   "Estado"
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
         Index           =   15
         Left            =   7455
         TabIndex        =   131
         Top             =   810
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo IRPF"
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
         Index           =   3
         Left            =   5895
         TabIndex        =   35
         Top             =   795
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Factura"
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
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   75
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fac"
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
         Left            =   5895
         TabIndex        =   33
         Top             =   75
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
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
         Left            =   150
         TabIndex        =   32
         Top             =   795
         Width           =   555
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   780
         ToolTipText     =   "Buscar Socio"
         Top             =   795
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "NºFactura"
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
         Index           =   28
         Left            =   4620
         TabIndex        =   31
         Top             =   75
         Width           =   1125
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   6960
         Picture         =   "frmManFactSocios.frx":020C
         ToolTipText     =   "Buscar fecha"
         Top             =   75
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   90
      TabIndex        =   26
      Top             =   10275
      Width           =   2175
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
         Height          =   240
         Left            =   240
         TabIndex        =   27
         Top             =   135
         Width           =   1755
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
      Left            =   12690
      TabIndex        =   24
      Top             =   10350
      Width           =   1065
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
      Left            =   11520
      TabIndex        =   23
      Top             =   10350
      Width           =   1065
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
      Left            =   2520
      Top             =   8370
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
      Left            =   2385
      Top             =   8370
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
      Height          =   375
      Left            =   4350
      Top             =   8400
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
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
   Begin MSAdodcLib.Adodc Data7 
      Height          =   465
      Left            =   4410
      Top             =   8340
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   13320
      TabIndex        =   159
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
      Begin VB.Menu mnRecepcion 
         Caption         =   "&Recepción Factura"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnGeneracion 
         Caption         =   "&Generación Entradas"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnAnticipo 
         Caption         =   "&Anticipo sin Entrada"
         Shortcut        =   ^A
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
Attribute VB_Name = "frmManFactSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 6001



'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public NumAlbar As String  ' venimos de pedidos para insertar envases paletizacion

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
Private WithEvents frmLFac As frmManLinFactSocios 'Lineas de variedades de facturas socios
Attribute frmLFac.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de iva de conta
Attribute frmTIva.VB_VarHelpID = -1

Private WithEvents frmVar As frmComVar 'Form Mto de variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Form Mto de socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'Form Mto de calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmGas As frmManConcepGasto 'Form Mto de concepto de gastos
Attribute frmGas.VB_VarHelpID = -1
Private WithEvents frmList As frmListado  'para recepcion de nro de factura
Attribute frmList.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'para asignacion de albaranes
Attribute frmMens.VB_VarHelpID = -1

Private WithEvents frmFac As frmBasico2
Attribute frmFac.VB_VarHelpID = -1


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

'[Monica]26/07/2013: cambio de la fecha anterior en la factura sin contabilizar
Dim FecAnterior As String
Dim SocioAnterior As String
Dim IRPFAnterior As Integer
Dim IvaAnterior As String
Dim ObsAnterior As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Conceptos de gastos
            Set frmGas = New frmManConcepGasto
            frmGas.DatosADevolverBusqueda = "0|1|"
            frmGas.CodigoActual = txtAux3(14).Text
            frmGas.Show vbModal
            Set frmGas = Nothing
            PonerFoco txtAux3(14)
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
                '[Monica]18/01/2017: solo en el caso de Montifrut no recalculo totales pq lo hacia incorrecto
                If vParamAplic.Cooperativa <> 12 Then CalcularTotales
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
                    InsertarLinea 1
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


Private Sub CmdAlbaranes_Click()
    
'    If Modo <> 2 Then Exit Sub
'    If Me.FrameAlbaranes.visible = False Then
'        Me.FrameAlbaranes.visible = True
'        Me.Frame3.visible = False
'        Me.CmdAlbaranes.Picture = frmPpal.imgListPpal.ListImages(36).Picture
'        Me.CmdAlbaranes.ToolTipText = "Volver de Albaranes"
'    Else
'        Me.FrameAlbaranes.visible = False
'        Me.Frame3.visible = True
'        Me.CmdAlbaranes.Picture = frmPpal.imgListPpal.ListImages(8).Picture
'        Me.CmdAlbaranes.ToolTipText = "Ver Albaranes de Factura"
'    End If

End Sub

Private Sub cmdAnticipos_Click()
    If Modo <> 2 Then Exit Sub
    If Me.FrameAnticipos.visible = False Then
        Me.DataGrid1.visible = False
        Me.FrameAnticipos.visible = True
        Me.Frame4.visible = False
        Me.cmdAnticipos.Picture = frmPpal.imgListPpal.ListImages(36).Picture
        Me.cmdAnticipos.ToolTipText = "Volver de Anticipos"
    Else
        Me.DataGrid1.visible = True
        Me.FrameAnticipos.visible = False
        Me.Frame4.visible = True
        Me.cmdAnticipos.Picture = frmPpal.imgListPpal.ListImages(9).Picture
        Me.cmdAnticipos.ToolTipText = "Ver Anticipos de Liquidación"
    End If

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
            
'            ComprobarClasificacion
            
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid4.AllowAddNew = False
                If Not Data5.Recordset.EOF Then Data5.Recordset.MoveFirst
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
    Combo1(3).ListIndex = 0
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    
    Text1(5).Text = 0
    Text1(6).Text = 0
    Text1(7).Text = 0
    
    Text1(8).Text = vParamAplic.PorcreteFacSoc
    Text1(13).Text = vParamAplic.PorcenAFO
    
    
    If vParamAplic.Cooperativa = 14 And (Mid(Combo1(0).Text, 1, 3) = "SUB" Or Mid(Combo1(0).Text, 1, 3) = "SIN") Then
        Text1(8).Text = 0
        Text1(13).Text = 0
    End If
    
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
        anc = DataGrid2.Top
        If DataGrid2.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid2"
        
        
        'Si pasamos el control aqui lo ponemos en amarillo
        Combo1(0).SetFocus
        Combo1(0).BackColor = vbLightBlue
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
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
        CadenaConsulta = "Select rfactsoc.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
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

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    NumTabMto = Index
    If Index = 2 Then NumTabMto = 3
    ModificaLineas = 2
    PonerModo 5, Index
 

    
        vWhere = ObtenerWhereCP(False)
        If Not BloqueaRegistro("rfactsoc", vWhere) Then
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

        For J = 10 To 13
            txtAux3(J).Text = DataGrid4.Columns(J - 10).Text
        Next J
        txtAux3(14).Text = DataGrid4.Columns(4).Text

        txtAux3(15).Text = DataGrid4.Columns(5).Text
        txtAux3(16).Text = DataGrid4.Columns(6).Text

        ModificaLineas = 2 'Modificar
        LLamaLineas ModificaLineas, anc, "DataGrid4"

        'Añadiremos el boton de aceptar y demas objetos para insertar
        Me.lblIndicador.Caption = "MODIFICAR"
        PonerModoOpcionesMenu (Modo)
        PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
        DataGrid4.Enabled = True

'            PonerBotonCabecera False
        PonerFoco txtAux3(16)
        Me.DataGrid4.Enabled = False


    
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
                If jj = 4 Or jj = 5 Or jj = 6 Or jj = 7 Or jj = 8 Then
                    txtAux(jj).Height = DataGrid1.RowHeight
                    txtAux(jj).Top = alto
                    txtAux(jj).visible = b
                End If
            Next jj
            
        Case "DataGrid2"
            DeseleccionaGrid Me.DataGrid2
            b = (xModo = 1)
            For jj = 0 To 8
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto - 210
'                txtAux3(jj).visible = b
            Next jj
            chkAux(0).Top = alto - 210
'            chkAux(0).visible = b
            
        Case "DataGrid4"
            DeseleccionaGrid Me.DataGrid4
            b = (xModo = 1 Or xModo = 2)
             For jj = 14 To 16
                txtAux3(jj).Height = DataGrid3.RowHeight - 10
                txtAux3(jj).Top = alto + 5
                txtAux3(jj).visible = b
            Next jj
            btnBuscar(0).Height = DataGrid3.RowHeight - 10
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = b
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
    Cad = Cad & "Va a eliminar la Factura de Socio:            "
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


Private Sub CmdGastos_Click()
Dim cadena As String

'    If Modo <> 2 Then Exit Sub
'    If Me.FrameGastos.visible = False Then
'        Me.FrameGastos.visible = True
'        Me.Frame3.visible = False
'        Me.CmdGastos.Picture = frmPpal.imgListPpal.ListImages(36).Picture
'        Me.CmdGastos.ToolTipText = "Volver de Gastos"
'    Else
'        Me.FrameGastos.visible = False
'        Me.Frame3.visible = True
'        Me.CmdGastos.Picture = frmPpal.imgListPpal.ListImages(8).Picture
'        Me.CmdGastos.ToolTipText = "Ver Gastos a pie de Factura"
'    End If

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

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim i As Integer
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
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
            Text1(12).Text = Mid(Trim(Combo1(Index).List(i)), 1, 3)
            CodTipoMov = Text1(12).Text
            
            If Modo = 3 And vParamAplic.Cooperativa = 14 And (Mid(Combo1(0).Text, 1, 3) = "SUB" Or Mid(Combo1(0).Text, 1, 3) = "SIN") Then
                Text1(8).Text = 0
                Text1(13).Text = 0
            End If
            
            
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

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

'    If LastCol = -1 Then Exit Sub

    'Datos de la tabla albaran_calibres
    If Not Data3.Recordset.EOF Then
        'Datos de la tabla rfactsoc_calidad
        CargaGrid DataGrid1, Data2, True
        CargaGrid DataGrid3, Data4, True
    Else
        'Datos de la tabla rfactsoc_calidad
        CargaGrid DataGrid1, Data2, False
        CargaGrid DataGrid3, Data4, False
    End If
    
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
    btnPrimero = 16
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(4).Image = 3   'Insertar
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(6).Image = 5   'Borrar
'        .Buttons(8).Image = 10  'Impresión de factura
'        .Buttons(9).Image = 26  'recepcion de nro de factura
'        .Buttons(10).Image = 34  'generacion de entradas de facturas de siniestros
'        .Buttons(11).Image = 25 ' generacion de una factura de anticipo sin entradas
'        .Buttons(13).Image = 11  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With
    
    ' ICONETS DE LA BARRA
    btnPrimero = 18 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 10  son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26  'recepcion de nro de factura
        .Buttons(2).Image = 34  'generacion de entradas de facturas de siniestros
        .Buttons(3).Image = 25 ' generacion de una factura de anticipo sin entradas
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
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
    Me.cmdAnticipos.Picture = frmPpal.imgListPpal.ListImages(9).Picture
    
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    
    CodTipoMov = "FAA"
    VieneDeBuscar = False
    Me.SSTab1.Tab = 0
        
    '## A mano
    NombreTabla = "rfactsoc"
    NomTablaLineas = "rfactsoc_variedad" 'Tabla de variedades de factura
    Ordenacion = " ORDER BY rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from rfactsoc "
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
    
    Label5.visible = (vParamAplic.Cooperativa = 12)
    Text1(22).visible = (vParamAplic.Cooperativa = 12)
    Text1(22).Enabled = (vParamAplic.Cooperativa = 12)
    
    
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

    '[Monica]13/05/2013: En Montifrut todo son liquidaciones, amrcarmos si son o no anticipos y
    '                    son o no facturas de venta campo
    If vParamAplic.Cooperativa = 12 Then
        Check1(3).Caption = "Es Anticipo"
        Check1(4).Caption = "Es Venta Campo"
    End If


End Sub


Private Sub frmFac_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String

    If CadenaSeleccion <> "" Then
        Text1(12).Text = RecuperaValor(CadenaSeleccion, 1) 'codtipom
        Text1(0).Text = RecuperaValor(CadenaSeleccion, 2) 'numfactu
        Text1(1).Text = RecuperaValor(CadenaSeleccion, 3) 'fecfactu
        
        cadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
        If cadB <> "" Then
            'Se muestran en el mismo form
            CadenaConsulta = "select * from rfactsoc WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If

End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' Recepcion de nro de factura
            mnRecepcion_Click
        Case 2 ' generacion de entradas de siniestro
            mnGeneracion_Click
        Case 3 ' generacion de una factura de anticipo sin entradas
            mnAnticipo_Click
    End Select
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
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
    Me.Combo1(3).ListIndex = -1
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
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(12), CadenaDevuelta, 1)
        cadB = cadB & Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        cadB = cadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
        cadB = cadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
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

Private Sub frmGas_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Concepto de gastos
    txtAux3(14).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Concepto
    txtAux3(15).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Concepto de gasto

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

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Calidades
    txtAux(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'Cod Calidad
    txtAux(6).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Calidad
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        If vParamAplic.Cooperativa = 18 Then
            If InsertarAlbaranesFrutasInma(CadenaSeleccion) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                CargaGrid DataGrid5, Data6, True
                CargaGrid DataGrid2, Data3, True
                CargaGrid DataGrid4, Data5, True
            End If
        Else
            If InsertarAlbaranes(CadenaSeleccion) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                CargaGrid DataGrid5, Data6, True
                CargaGrid DataGrid2, Data3, True
                CargaGrid DataGrid4, Data5, True
            End If
        End If
    End If
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

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Socios
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Socios
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
        
        Case 0 'Socios
            indice = 2
            PonerFoco Text1(indice)
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
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


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 23
        frmZ.pTitulo = "Observaciones de la Factura"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub



Private Sub mnAnticipo_Click()
    AbrirListadoAnticipos (161)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()

    '[Monica]02/12/2014: solo en el caso de picassent damos aviso de que puede haber algo en ringresos
    If (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16) And Mid(Combo1(0).Text, 1, 3) = "FAT" Then
        MsgBox "Puede que esta factura tenga un registro asociado en ingresos a incluir en liquidación. Actualícelos", vbExclamation
    End If
    


    '[Monica]05/02/2014: Solo para el caso de Montifrut dejamos eliminar una factura contabilizada
    If Check1(1).Value = 1 Then
        If MsgBox("Esta factura está en Contabilidad y Arimoney. " & vbCrLf & vbCrLf & "Si la elimina, elimínela también en estas aplicaciones." & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        Else
        '[Monica]02/09/2013: añadida la fecha de ultima liquidacion de iva
            If CDate(Text1(1).Text) <= vEmpresa.FechaUltIVA Then
                If MsgBox("La factura es de un período liquidado. " & vbCrLf & vbCrLf & "¿ Seguro que desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If


    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub



Private Sub mnGeneracion_Click()

    AbrirListado 39

End Sub

Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir
End Sub

Private Sub mnRecepcion_Click()
'Recepcion de nro de factura

    'If Check1(5).Value = 0 Then
    If Combo1(3).ListIndex <> 1 Then
        MsgBox "Esta factura no está pendiente de recepción de número.", vbExclamation
        Exit Sub
    End If

    Set frmList = New frmListado
    
    frmList.Opcionlistado = 38
    frmList.NumCod = "codtipom = '" & Mid(Combo1(0).Text, 1, 3) & "' and numfactu = " & Text1(0).Text & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    frmList.Show vbModal
    
    Set frmList = Nothing
    
'    CadB = "codtipom = '" & Mid(Combo1(0).Text, 1, 3) & "' and numfactu = " & Text1(0).Text & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    PonerCadenaBusqueda
End Sub



Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()

    '[Monica]02/12/2014: solo en el caso de picassent damos aviso de que puede haber algo en ringresos
    If (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16) And Mid(Combo1(0).Text, 1, 3) = "FAT" Then
        MsgBox "Puede que esta factura tenga un registro asociado en ingresos a incluir en liquidación. Actualícelos", vbExclamation
    End If


    If Check1(1).Value = 1 Then
        If MsgBox("Esta factura está en Contabilidad y Arimoney. " & vbCrLf & vbCrLf & "Si la modifica realice los cambios en estas aplicaciones." & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        Else
        '[Monica]02/09/2013: añadida la fecha de ultima liquidacion de iva
            If CDate(Text1(1).Text) <= vEmpresa.FechaUltIVA Then
                If MsgBox("La factura es de un período liquidado. " & vbCrLf & vbCrLf & "¿ Seguro que desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If

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
    
        Case 2 'Socio
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
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
                    ' sacamos el iva del socio
                    Set vSocio = New cSocio
                    If vSocio.LeerDatosSeccion(Text1(2).Text, vParamAplic.Seccionhorto) Then
                        Text1(3).Text = vSocio.CodIva
                        Combo1(1).ListIndex = vSocio.TipoIRPF
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
                    Set vSocio = Nothing
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
            
        Case 4 'Campo
        
        
        Case 10 ' importe fondo operativo
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 1
            
        
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
Dim cadAux As String
    
'    '--- Laura 12/01/2007
'    cadAux = Text1(5).Text
'    If Text1(4).Text <> "" Then Text1(5).Text = ""
'    '---
    
'    '--- Laura 12/01/2007
'    Text1(5).Text = cadAux
'    '---
'    CadB = ObtenerBusqueda(Me)
    cadB = ObtenerBusqueda2(Me, BuscaChekc, 1)

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select rfactsoc.* from " & NombreTabla & " LEFT JOIN rfactsoc_variedad ON rfactsoc.codtipom=rfactsoc_variedad.codtipom "
        CadenaConsulta = CadenaConsulta & " and rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB & " GROUP BY rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu " & Ordenacion
'        CadenaConsulta = "select palets.* from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY palets.numpalet " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
    
    Set frmFac = New frmBasico2
    AyudaFrasSocios frmFac, , cadB
    Set frmFac = Nothing

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
'            Text1(0).BackColor = vbLightBlue
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
    
    CargaGrid DataGrid2, Data3, True
    CargaGrid DataGrid4, Data5, True
    CargaGrid DataGrid5, Data6, True
    CargaGrid DataGrid6, Data7, True
    '++monica
    If Data3.Recordset.RecordCount > 0 Then
        CargaGrid DataGrid1, Data2, True
        CargaGrid DataGrid3, Data4, True
    Else
        CargaGrid DataGrid1, Data2, False
        CargaGrid DataGrid3, Data4, False
    End If
    
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

    PosicionarCombo2 Combo1(0), Text1(12).Text
    
    ' datos de la factura que rectifica si es una factura rectificativa
    If b Then b = PonerCamposForma2(Me, Data1, 2, "Frame6")
    Combo1(2).ListIndex = -1
    PosicionarCombo2 Combo1(2), Text1(20).Text
    
    '[Monica]02/05/2013: si es montifrut y no tenemos albaranes pq es una factura a ojo, podemos añadirlos
    If vParamAplic.Cooperativa = 12 Or vParamAplic.Cooperativa = 18 Then
        '???? aqui
        Me.ToolAux(2).visible = True
        Me.ToolAux(2).Enabled = True
    End If
    
    VisualizarAnticipos

'    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    
    Text2(2).Text = PonerNombreDeCod(Text1(2), "rsocios", "nomsocio", "codsocio", "N") 'socios
    
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
    
    CalcularGastos
    
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
'    ActualizarToolbar Modo, Kmodo
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
    
    
    cmdAnticipos.visible = b
    cmdAnticipos.Enabled = b
'    CmdGastos.visible = b
'    CmdGastos.Enabled = b
'    CmdAlbaranes.visible = b
'    CmdAlbaranes.Enabled = b
    
    
    
    
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    '[Monica]25/07/2013: cambiamos la fecha de factura
    If Modo = 4 Then
        Text1(1).Locked = False
        Text1(1).Enabled = True
        Text1(1).BackColor = vbWhite
        
        FecAnterior = Text1(1).Text
        SocioAnterior = Text1(2).Text
        IRPFAnterior = Combo1(1).ListIndex
        IvaAnterior = Text1(3).Text
        ObsAnterior = Text1(23).Text
    End If
    
    
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
    Combo1(3).Enabled = (Modo = 1)
    
    For i = 4 To 18 '7
        BloquearTxt Text1(i), Not (Modo = 1)
        Text1(i).Enabled = (Modo = 1)
    Next i
    Text1(0).Enabled = (Modo = 1)
    
    BloquearTxt Text1(9), Not (Modo = 1)
    Text1(9).Enabled = (Modo = 1)
    BloquearTxt Text1(11), Not (Modo = 1)
    Text1(11).Enabled = (Modo = 1)
    ' recepcion de nro de factura alfa para contabilidad
    BloquearTxt Text1(21), Not (Modo = 1)
    Text1(21).Enabled = (Modo = 1)
    
    For i = 0 To Check1.Count - 1
        Me.Check1(i).Enabled = (Modo = 1)
    Next i
    
    b = (Modo <> 1)
    'Campos Nº Pedido bloqueado y en azul
    BloquearTxt Text1(0), b, True
'    BloquearTxt Text1(3), b 'referencia
 
    
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i

    txtAux(6).visible = False
    txtAux(6).Enabled = True
    For i = 0 To 7
        BloquearTxt txtAux3(i), True
        txtAux3(i).visible = False
    Next i
    For i = 3 To 8
        If i <> 4 Then
            BloquearTxt txtAux3(i), (Modo <> 1)
            txtAux3(i).visible = (Modo = 1)
        End If
    Next i
    BloquearChk chkAux(0), (Modo <> 1)
    chkAux(0).visible = (Modo = 1)
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    CmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
    
    imgFec(0).Enabled = (Modo = 1 Or Modo = 3)
    imgFec(0).visible = (Modo = 1 Or Modo = 3)
    
    Text1(2).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    b = (Modo = 5) And ModificaLineas = 1
    BloquearTxt txtAux3(14), Not b
    BloquearBtn Me.btnBuscar(0), Not b
    
    txtAux3(15).visible = False
    txtAux3(15).Enabled = False
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    BloquearTxt txtAux3(16), Not b
       
       
    '[Monica]21/05/2013: introducimos el porcentaje de corredor en la factura
    Text1(22).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4) And vParamAplic.Cooperativa = 12
       
       
       
       
    ' DATOS SI ES RECTIFICATIVA
    b = (Modo = 1)
    Combo1(2).Enabled = b
    If b Then
        Combo1(2).BackColor = vbWhite
    Else
        Combo1(2).BackColor = &H80000018 'Amarillo Claro
    End If
    Text1(17).Enabled = b
    Text1(18).Enabled = b
    Text1(19).Enabled = (b Or Modo = 4)
    imgFec(1).Enabled = b
    imgFec(1).visible = b
       
       
    BloquearImgZoom Me, Modo
    
    
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
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
    If Modo = 3 Then
        If vParamAplic.Cooperativa = 12 Then
            '[Monica]20/06/2017: control de fechas que antes no estaba, solo para Montifrut que cuando integra no coge la fecha de recepcion
            ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(1).Text))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                Exit Function
            End If
        End If
    End If



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
        PonerFocoBtn Me.CmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.CmdAceptar
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
            Case 0 'rfactsoc_variedades
                Select Case Button.Index
                    Case 1 'añadir variedad
                        Set frmLFac = New frmManLinFactSocios
                        
                        frmLFac.ModoExt = 3
                        frmLFac.tipoMov = Data1.Recordset.Fields(0).Value
                        frmLFac.Factura = Data1.Recordset.Fields(1).Value
                        frmLFac.fecfactu = Data1.Recordset.Fields(2).Value
                        frmLFac.Show vbModal
                    
                        Set frmLFac = Nothing
                        
                    Case 2 'modificar variedad
                        Set frmLFac = New frmManLinFactSocios
                        
                        frmLFac.ModoExt = 4
                        frmLFac.tipoMov = Data3.Recordset.Fields(0).Value
                        frmLFac.Factura = Data3.Recordset.Fields(1).Value
                        frmLFac.fecfactu = Data3.Recordset.Fields(2).Value
                        frmLFac.Variedad = Data3.Recordset.Fields(3).Value
                        frmLFac.campo = Data3.Recordset.Fields(5).Value
                        frmLFac.Show vbModal
                        
                        Set frmLFac = Nothing
                        
                    Case 3 ' boton eliminar linea de variedades
                        BotonEliminarLinea 0
                    Case Else
                End Select
                CalcularTotales
                CalcularGastos
                PonerCampos
                TerminaBloquear
                
            Case 1 'rfactsoc_gastos
                Select Case Button.Index
                    Case 1
                        BotonAnyadirLinea Index
                    Case 2
                        BotonModificarLinea Index
                    Case 3
                        BotonEliminarLinea Index
                    Case Else
                End Select
                

'                CalcularGastos
'                PonerCampos
'                TerminaBloquear
                
            Case 2 ' rfactsoc_albaranes
                AsignarAlbaranes
                
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
            
            If Data3.Recordset.RecordCount = 1 Then
                MsgBox "No podemos dejar la factura sin variedades. Elimine la factura.", vbExclamation
                Exit Sub
            End If
            
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar la Variedad?"
            Cad = Cad & vbCrLf & "Tipo: " & Data3.Recordset.Fields(0)
            Cad = Cad & vbCrLf & "Factura: " & Data3.Recordset.Fields(1)
            Cad = Cad & vbCrLf & "Fecha: " & Data3.Recordset.Fields(2)
            Cad = Cad & vbCrLf & "Variedad: " & Data3.Recordset.Fields(3)
            Cad = Cad & vbCrLf & "Campo: " & Data3.Recordset.Fields(5)
            
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Data3.Recordset.AbsolutePosition
                
                If Not EliminarLinea(Index) Then
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
                        CalcularGastos
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
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 1  'Añadir
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 8  ' Impresion de albaran
            mnImprimir_Click
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

    Me.CmdAceptar.visible = Not b
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

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid1"
            Opcion = 1
        Case "DataGrid2"
            Opcion = 2
        Case "DataGrid3" 'anticipos
            Opcion = 3
        Case "DataGrid4" 'Gastos de Pie de factura
            Opcion = 4
        Case "DataGrid5" 'Albaranes de almazara y de bodega
            Opcion = 5
        Case "DataGrid6" ' facturas varias descontadas
            Opcion = 6
    End Select
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
     b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
     vDataGrid.Enabled = Not b
    
   
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
        Case "DataGrid1" 'rfactsoc_calidad
'           SQL = "SELECT codtipom,numfactu,fecfactu,codsocio,codvarie,codcalid,kilosnet,precio,imporcal
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(4)|T|Codigo|1200|;"
            tots = tots & "S|txtAux(5)|T|Nombre Calidad|3900|;"
            tots = tots & "S|txtAux(6)|T|Kilos|2040|;"
            tots = tots & "S|txtAux(7)|T|Precio Cal.|1600|;"
            tots = tots & "S|txtAux(8)|T|Importe|1800|;"
            tots = tots & "S|txtAux(10)|T|Pr.Aplicado|1980|;"
            arregla tots, DataGrid1, Me, 350
'            DataGrid1.Columns(11).Alignment = dbgCenter
'            DataGrid1.Columns(12).Alignment = dbgRight
'            DataGrid1.Columns(13).Alignment = dbgRight
'            DataGrid1.Columns(14).Alignment = dbgRight
                       
         Case "DataGrid2" 'rfactsoc_variedad
'           SQL = "SELECT codtipom,numfactu,fecfactu,codsocio,codvarie,nomvarie,kilosnet,preciomed,imporvar "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(3)|T|Codigo|1200|;"
            tots = tots & "S|txtAux3(4)|T|Variedad|3000|;"
            tots = tots & "S|txtAux3(8)|T|Campo|1600|;"
            tots = tots & "S|txtAux3(5)|T|Kilos|1840|;S|txtAux3(6)|T|Pr.Medio|1400|;S|txtAux3(7)|T|Importe|1550|;S|txtAux3(9)|T|Imp.Gastos|1570|;"
            tots = tots & "N||||0|;S|chkAux(0)|CB|D|360|;"
            arregla tots, DataGrid2, Me, 350
            
'            DataGrid2.Columns(5).Alignment = dbgLeft
'            DataGrid2.Columns(7).Alignment = dbgLeft
'            DataGrid2.Columns(9).Alignment = dbgLeft
                     
         Case "DataGrid3" 'rfactsoc_anticipos
'           SQL = "SELECT codtipom,numfactu,fecfactu,codtipomanti,numfactuanti,fecfactuanti,codvarieanti,codvarieanti,nomvarie,codcampoanti,baseimpo "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux4(3)|T|Tipo Factura|2400|;"
            tots = tots & "S|txtAux4(4)|T|Factura|1000|;"
            tots = tots & "S|txtAux4(5)|T|F.Factura|1700|;"
            tots = tots & "S|txtAux4(6)|T|Codigo|900|;"
            tots = tots & "S|txtAux4(7)|T|Variedad|3100|;"
            tots = tots & "S|txtAux4(8)|T|Campo|1700|;"
            tots = tots & "S|txtAux4(9)|T|Importe|1700|;"
            arregla tots, DataGrid3, Me, 350
            
         Case "DataGrid4" 'rfactsoc_gastos
'           SQL = "SELECT codtipom,numfactu,fecfactu,codtipomanti,numfactuanti,fecfactuanti,codvarieanti,codvarieanti,nomvarie,codcampoanti,baseimpo "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(14)|T|Gasto|1100|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|txtAux3(15)|T|Descripción|8700|;"
            tots = tots & "S|txtAux3(16)|T|Importe|2750|;"
            arregla tots, DataGrid4, Me, 350
    
         Case "DataGrid5" 'rfactsoc_albaran
'           SQL = "SELECT codtipom,numfactu,fecfactu,codtipomanti,numfactuanti,fecfactuanti,codvarieanti,codvarieanti,nomvarie,codcampoanti,baseimpo "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux1(3)|T|Albaran|1000|;S|txtAux1(4)|T|Fecha|1400|;N||||0|;"
            tots = tots & "S|txtAux1(5)|T|Variedad|2100|;"
            tots = tots & "S|txtAux1(6)|T|Campo|1200|;S|txtAux1(7)|T|K.Neto|1500|;"
            tots = tots & "S|txtAux1(8)|T|Grado|1000|;S|txtAux1(9)|T|Precio|1000|;"
            tots = tots & "S|txtAux1(10)|T|Importe|1700|;S|txtAux1(11)|T|Imp.Gasto|1650|;"
            arregla tots, DataGrid5, Me, 350
    
            DataGrid5.Columns(6).Alignment = dbgLeft
    
         Case "DataGrid6" 'rfactsoc_fvarias
'           SQL = "SELECT codtipom,numfactu,fecfactu,codtipomfvar,numfactufvar,fecfactufvar,importeltotal "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux5(3)|T|Tipo Factura|2300|;S|txtAux5(4)|T|Factura|1300|;"
            tots = tots & "S|txtAux5(5)|T|Fecha|1400|;S|txtAux5(6)|T|Codigo|900|;S|txtAux5(7)|T|Forma de Pago|4400|;"
            tots = tots & "S|txtAux5(8)|T|Total Factura|2200|;"
            arregla tots, DataGrid6, Me, 350
    
            DataGrid6.Columns(7).Alignment = dbgLeft
            DataGrid6.Columns(8).Alignment = dbgRight
    
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub




Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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
        Case 4 ' calidad
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
                PonerFormatoEntero txtAux(Index)
                CmdAceptar.SetFocus
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
        
        
    Sql = " " & ObtenerWhereCP(True)
        
        
    '------------------------------------------------------------------------------
    '  LOG de acciones
    If Me.Check1(1).Value = 1 Then
        Set LOG = New cLOG
        
        LOG.Insertar 12, vUsu, "Elimina Factura: " & Text1(12).Text & "-" & Text1(0).Text & "-" & Text1(1).Text
    
        Set LOG = Nothing
    End If
    '-----------------------------------------------------------------------------
        
    'Eliminar en tablas de cabecera de factura
    '------------------------------------------
    
    'Lineas de calidades (rfactsoc_calidad)
    conn.Execute "Delete from rfactsoc_calidad " & Sql
    
    'Lineas de albaranes de bodega (rfactsoc_albaran)
    conn.Execute "Delete from rfactsoc_albaran " & Sql
    
    'Lineas de variedades (rfactsoc_variedad)
    conn.Execute "Delete from rfactsoc_variedad " & Sql
    
    'Antes de borrar anticipos desmarcarlos como liquidados
    conn.Execute "update rfactsoc_variedad set descontado = 0 where (codtipom, numfactu, fecfactu, codvarie, codcampo) in (select codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti from rfactsoc_anticipos " & Sql & ")"
    
    'Lineas de anticipos (rfactsoc_anticipos)
    conn.Execute "Delete from rfactsoc_anticipos " & Sql
    
    '[Monica]05/12/2011: anticipos de retirada de quatretonda
    'Antes de borrar retirada desmarcarlos como liquidados
    conn.Execute "update rfactsoc_variedad set descontado = 0 where (codtipom, numfactu, fecfactu, codvarie, codcampo) in (select codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti from rfactsoc_retirada " & Sql & ")"
        
    'Lineas de retirada (rfactsoc_retirada)
    conn.Execute "Delete from rfactsoc_retirada " & Sql
    
    'Lineas de gastos a pie de la factura (rfactsoc_gastos)
    conn.Execute "Delete from rfactsoc_gastos " & Sql
    
    
    '[Monica]15/04/2013: antes de descontar las facturas varias hemos de desmarcarlas como que han sido descontadas
    conn.Execute "update fvarcabfact set intliqui = 0 where (codtipom,numfactu,fecfactu) in (select codtipomfvar,numfactufvar,fecfactufvar from rfactsoc_fvarias " & Sql & ")"
    
    'facturas varias a descontar
    conn.Execute "Delete from rfactsoc_fvarias " & Sql
    
    '[Monica]10/04/2017: incrementos de factura en la liquidacion de Picassent y Coopic
    'incrementos de factura
    Dim vAux As String
    vAux = "insert into ringresos (codsocio,codvarie,concepto,importe) " & _
                " select aa.codsocio, rfactsoc_ingresos.codvarie, rfactsoc_ingresos.concepto, rfactsoc_ingresos.importe  from rfactsoc_ingresos , (select codsocio from rfactsoc  " & Sql & ") aa " & _
                "  " & Sql
     
    conn.Execute vAux
    
    conn.Execute "Delete from rfactsoc_ingresos " & Sql
    
    
    'Cabecera de factura (rfactsoc)
    conn.Execute "Delete from " & NombreTabla & Sql
    
    'Decrementar contador si borramos el ultima factura
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador Text1(12).Text, Val(Text1(0).Text)
    Set vTipoMov = Nothing
    
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
Dim cadena As String

    On Error GoTo FinEliminar
    
    Select Case Aux
        Case 0
            'Eliminar en tablas de variedades y calidades
            '------------------------------------------
            If Data3.Recordset.EOF Then Exit Function
                
            conn.BeginTrans
                
            Mens = ""
            
            '------------------------------------------------------------------------------
            '  LOG de acciones
            cadena = Data3.Recordset.Fields(0) & " " & Data3.Recordset.Fields(1) & " " & Data3.Recordset.Fields(2) & " " & Data3.Recordset.Fields(3) & " " & Data3.Recordset.Fields(5)
            
            Set LOG = New cLOG
            LOG.Insertar 12, vUsu, "Eliminar Linea variedad: " & cadena & vbCrLf
            Set LOG = Nothing
            '-----------------------------------------------------------------------------
            
            Sql = " where codtipom = '" & Data3.Recordset.Fields(0) & "'"
            Sql = Sql & " and numfactu = " & Data3.Recordset.Fields(1)
            Sql = Sql & " and fecfactu = " & DBSet(Data3.Recordset.Fields(2), "F")
            Sql = Sql & " and codvarie = " & Data3.Recordset.Fields(3)
            Sql = Sql & " and codcampo = " & Data3.Recordset.Fields(5)
            
            'Lineas de calidad (rfactsoc_calidad)
            conn.Execute "Delete from rfactsoc_calidad " & Sql
        
            'Lineas de variedades
            conn.Execute "Delete from rfactsoc_variedad " & Sql
    
        Case 1 ' linea de gastos a pie de pagina
            'Eliminar en tablas de gastos a pie
            '------------------------------------------
            If Data5.Recordset.EOF Then Exit Function
                
            conn.BeginTrans
                
            '------------------------------------------------------------------------------
            '  LOG de acciones
            cadena = Data3.Recordset.Fields(0) & " " & Data3.Recordset.Fields(1) & " " & Data3.Recordset.Fields(2) & " " & Data3.Recordset.Fields(3) & " " & Data3.Recordset.Fields(5)
            
            Set LOG = New cLOG
            LOG.Insertar 12, vUsu, "Eliminar Linea Gastos: " & cadena & vbCrLf
            Set LOG = Nothing
            '-----------------------------------------------------------------------------
            
            
            Mens = ""
            
            Sql = " where codtipom = '" & Data5.Recordset.Fields(0) & "'"
            Sql = Sql & " and numfactu = " & Data5.Recordset.Fields(1)
            Sql = Sql & " and fecfactu = " & DBSet(Data5.Recordset.Fields(2), "F")
            Sql = Sql & " and numlinea = " & DBSet(Data5.Recordset.Fields(3), "N")
            
            'Lineas de gastos (rfactsoc_gastos)
            conn.Execute "Delete from rfactsoc_gastos " & Sql
        
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

    CargaGrid DataGrid2, Data3, False 'variedades
    CargaGrid DataGrid1, Data2, False
    CargaGrid DataGrid3, Data4, False
    CargaGrid DataGrid4, Data5, False ' gastos de pie de la factura
    CargaGrid DataGrid5, Data6, False  ' albaranes de bodega y de almazara
    CargaGrid DataGrid6, Data7, False
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
    
    Sql = " codtipom= '" & Text1(12).Text & "'"
    Sql = Sql & " and numfactu = " & Text1(0).Text
    Sql = Sql & " and fecfactu = " & DBSet(Text1(1).Text, "F")

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
    Case 1  ' calidades
        Sql = "SELECT rfactsoc_calidad.codtipom, rfactsoc_calidad.numfactu, rfactsoc_calidad.fecfactu,"
        Sql = Sql & " rfactsoc_calidad.codvarie, rfactsoc_calidad.codcampo, rfactsoc_calidad.codcalid,"
        Sql = Sql & " rcalidad.nomcalid, rfactsoc_calidad.kilosnet, rfactsoc_calidad.precio, rfactsoc_calidad.imporcal, "
        Sql = Sql & " rfactsoc_calidad.preciocalidad "
        Sql = Sql & " FROM rfactsoc_calidad, rcalidad WHERE rfactsoc_calidad.codvarie = rcalidad.codvarie "
        Sql = Sql & " and rfactsoc_calidad.codcalid = rcalidad.codcalid "
    Case 2  ' variedades
        Sql = "SELECT rfactsoc_variedad.codtipom, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu, "
        Sql = Sql & " rfactsoc_variedad.codvarie, variedades.nomvarie, rfactsoc_variedad.codcampo, "
        Sql = Sql & " rfactsoc_variedad.kilosnet, rfactsoc_variedad.preciomed, rfactsoc_variedad.imporvar, rfactsoc_variedad.imporgasto,"
        Sql = Sql & " rfactsoc_variedad.descontado, IF(descontado=1,'*','') as ddescontado "
        Sql = Sql & " FROM rfactsoc_variedad, variedades " 'lineas de variedad de la factura socio
        Sql = Sql & " WHERE rfactsoc_variedad.codvarie = variedades.codvarie "
    Case 3  ' anticipos de venta campo
        Sql = "SELECT rfactsoc_anticipos.codtipom, rfactsoc_anticipos.numfactu, rfactsoc_anticipos.fecfactu, "
'        SQL = SQL & " CASE rfactsoc_anticipos.codtipomanti WHEN ""FAC"" THEN ""Anticipo V.Campo"" WHEN ""FAA"" THEN ""Anticipo"" END as codtipomanti,"
        Sql = Sql & " rfactsoc_anticipos.codtipomanti, "
        Sql = Sql & " rfactsoc_anticipos.numfactuanti, rfactsoc_anticipos.fecfactuanti, "
        Sql = Sql & " rfactsoc_anticipos.codvarieanti, variedades.nomvarie, rfactsoc_anticipos.codcampoanti, "
        Sql = Sql & " rfactsoc_anticipos.baseimpo "
        Sql = Sql & " FROM rfactsoc_anticipos, variedades " 'lineas de variedad de la factura socio
        Sql = Sql & " WHERE rfactsoc_anticipos.codvarieanti = variedades.codvarie "
    Case 4  ' gastos de pie de pagina
        Sql = "SELECT rfactsoc_gastos.codtipom, rfactsoc_gastos.numfactu, rfactsoc_gastos.fecfactu, "
        Sql = Sql & " rfactsoc_gastos.numlinea, rfactsoc_gastos.codgasto, rconcepgasto.nomgasto, "
        Sql = Sql & " rfactsoc_gastos.importe "
        Sql = Sql & " FROM rfactsoc_gastos, rconcepgasto " 'lineas de gastos de la factura socio
        Sql = Sql & " WHERE rfactsoc_gastos.codgasto = rconcepgasto.codgasto "
    Case 5  ' albaranes de almazara y de bodega
        Sql = "SELECT rfactsoc_albaran.codtipom, rfactsoc_albaran.numfactu, rfactsoc_albaran.fecfactu, "
        Sql = Sql & " rfactsoc_albaran.numalbar, rfactsoc_albaran.fecalbar, rfactsoc_albaran.codvarie, "
        Sql = Sql & " variedades.nomvarie, rfactsoc_albaran.codcampo, rfactsoc_albaran.kilosnet, "
        Sql = Sql & " rfactsoc_albaran.grado, rfactsoc_albaran.precio, rfactsoc_albaran.importe, "
        Sql = Sql & " rfactsoc_albaran.imporgasto "
        Sql = Sql & " FROM rfactsoc_albaran, variedades " 'lineas de albaranes de la factura socio
        Sql = Sql & " WHERE rfactsoc_albaran.codvarie = variedades.codvarie "
    Case 6 ' facturas varias descontadas
        Sql = "SELECT rfactsoc_fvarias.codtipom, rfactsoc_fvarias.numfactu, rfactsoc_fvarias.fecfactu, "
        Sql = Sql & " rfactsoc_fvarias.codtipomfvar, rfactsoc_fvarias.numfactufvar, rfactsoc_fvarias.fecfactufvar, "
        Sql = Sql & " fvarcabfact.codforpa, forpago.nomforpa, fvarcabfact.totalfac "
        Sql = Sql & " FROM rfactsoc_fvarias, fvarcabfact, forpago " 'lineas de facturas varias que se descuentan
        Sql = Sql & " WHERE rfactsoc_fvarias.codtipomfvar = fvarcabfact.codtipom and "
        Sql = Sql & " rfactsoc_fvarias.numfactufvar = fvarcabfact.numfactu and "
        Sql = Sql & " rfactsoc_fvarias.fecfactufvar = fvarcabfact.fecfactu and "
        Sql = Sql & " fvarcabfact.codforpa = forpago.codforpa "
    End Select
    
    If enlaza Then
        If Opcion = 6 Then
            Sql = Sql & " and rfactsoc_fvarias.codtipom= '" & Text1(12).Text & "'"
            Sql = Sql & " and rfactsoc_fvarias.numfactu = " & Text1(0).Text
            Sql = Sql & " and rfactsoc_fvarias.fecfactu = " & DBSet(Text1(1).Text, "F")
        Else
            Sql = Sql & " and " & ObtenerWhereCP(False)
        End If
        If Opcion = 1 Then
            Sql = Sql & " AND rfactsoc_calidad.codvarie=" & Data3.Recordset.Fields!codvarie
            Sql = Sql & " AND rfactsoc_calidad.codcampo=" & Data3.Recordset.Fields!codCampo
        End If
    Else
        If Opcion = 6 Then
            Sql = Sql & " and rfactsoc_fvarias.numfactu = -1 "
        Else
            Sql = Sql & " and numfactu = -1"
        End If
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
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    'Añadir
    Toolbar1.Buttons(1).Enabled = b
    Me.mnModificar.Enabled = b
    
    '[Monica]26/07/2013: quito la condicion de que si la factura esta contabilizarla no poder modificarla
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (hcoCodMovim = "") 'And Not (Check1(1).Value = 1)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    
    '[Monica]15/05/2012: quito la condicion de que si la factura esta impresa no poder eliminarla
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (hcoCodMovim = "") And Not (Check1(1).Value = 1) 'And Not (Check1(0).Value = 1)
    
    '[Monica]05/02/2014: para el caso de montifrut dejo eliminarla
    If vParamAplic.Cooperativa = 12 Then
        b = b Or Check1(1).Value = 1
    End If
    
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Impresión de albaran
    Toolbar1.Buttons(8).Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
    Me.mnImprimir.Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
    
    'Recepcion de nro de factura
    Toolbar5.Buttons(1).Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0 'And Data1.Recordset!pdtenrofact = 1
    Me.mnRecepcion.Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0 'And Data1.Recordset!pdtenrofact = 1

    If (Modo = 2) And Data1.Recordset.RecordCount > 0 Then
        Toolbar5.Buttons(1).Enabled = (CInt(Data1.Recordset!pdtenrofact) = 1)
        Me.mnRecepcion.Enabled = (CInt(Data1.Recordset!pdtenrofact) = 1)
    End If

    '
    
    
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta

    '[Monica]26/07/2013: quito la condicion de no poder modificarlo si esta contabilizado And Not Check1(1).Value = 1
    b = (Modo = 2) 'And Not Check1(1).Value = 1
    For i = 0 To 0
        ToolAux(i).Buttons(1).Enabled = b
        
        If b Then
            Select Case i
              Case 0
                bAux = (b And Me.Data3.Recordset.RecordCount > 0)
            End Select
        End If
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
    
    ' toolbar de gastos
    '[Monica]09/07/2013: añadida la condicion de que no pueda hacer nada si es montifrut
    b = ((Modo = 2) Or (Modo = 5)) And vParamAplic.Cooperativa <> 12 And vParamAplic.Cooperativa <> 18
    ToolAux(1).Buttons(1).Enabled = b
    
    bAux = False
    If b Then
        bAux = (b And Me.Data5.Recordset.RecordCount > 0) And vParamAplic.Cooperativa <> 12 And vParamAplic.Cooperativa <> 18
    End If
    ToolAux(1).Buttons(2).Enabled = bAux
    ToolAux(1).Buttons(3).Enabled = bAux


    ' toolbar de albaranes
    '[Monica]18/01/2019: añadida la cooperativa de frutas inma para poder añadir los albaranes
    b = ((Modo = 2) Or (Modo = 5)) And (vParamAplic.Cooperativa = 12 Or vParamAplic.Cooperativa = 18)
    ToolAux(2).Buttons(1).Enabled = b
    ToolAux(2).Buttons(2).Enabled = False
    ToolAux(2).Buttons(3).Enabled = False
    ToolAux(2).Buttons(2).visible = False
    ToolAux(2).Buttons(3).visible = False



End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NroCopias As Integer
Dim Tipo As Byte
Dim Sql5 As String
Dim EsComplemen As Byte

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Factura para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadselect = ""
    numParam = 0
    
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de albaran
    '---------------------------------------------------
    'Tipo de factura
    devuelve = "{" & NombreTabla & ".codtipom}='" & Mid(Combo1(0).Text, 1, 3) & "'"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "codtipom = '" & Mid(Combo1(0).Text, 1, 3) & "'"
    If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    
    Select Case Mid(Combo1(0).Text, 1, 3)
        Case "FRS" ' Impresion de facturas rectificativas
                   ' hacemos caso del codtipom que rectifica
                   
              Select Case Mid(Combo1(2).Text, 1, 3)
                    Case "FLI"
                        indRPT = 38 'Impresion de Factura Socio de Industria
                    Case Else
                        Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Mid(Combo1(0).Text, 1, 3), "T"))
                        If Tipo >= 7 And Tipo <= 10 Then
                            indRPT = 42 'Imporesion de Facturas de Bodega o Almazara
                        Else

'[Monica]07/02/2012: Hemos marcado las facturas que son complementarias, ya no hace falta esto
'
'                            '[Monica]07/02/2012: Si la factura que rectifico es complementaria le pasamos el parametro
'                            Sql5 = "select esliqcomplem from rfactsoc where (codtipom, numfactu, fecfactu) in (select rectific_codtipom, rectific_numfactu, rectific_fecfactu from rfactsoc where codtipom = " & DBSet(Mid(Combo1(0).Text, 1, 3), "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and fecfactu = " & DBSet(Text1(1).Text, "F") & ")"
'                            EsComplemen = DevuelveValor(Sql5)
'
'                            cadParam = cadParam & "pComplem=" & EsComplemen & "|"
'                            numParam = numParam + 1
                            
                            indRPT = 23 'Impresion de Factura Socio
                        End If
              End Select
        
        Case "FLI"
            indRPT = 38 'Impresion de Factura Socio de Industria
        Case Else
            Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Mid(Combo1(0).Text, 1, 3), "T"))
            If Tipo >= 7 And Tipo <= 10 Then
                indRPT = 42 'Imporesion de Facturas de Bodega o Almazara
            Else
                indRPT = 23 'Impresion de Factura Socio
                
'[Monica]07/02/2012: Hemos marcado las facturas que son complementarias, ya no hace falta esto
'
'                '[Monica]07/02/2012: enviamos si es o no una factura de liquidacion complementaria
'                cadParam = cadParam & "pComplem=" & Check1(5).Value & "|"
'                numParam = numParam + 1
            End If
    End Select
    
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    
    
    'Nº factura
    devuelve = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "numfactu = " & Val(Text1(0).Text)
    If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    
    'Fecha Factura
    devuelve = "{" & NombreTabla & ".fecfactu}=Date(" & Year(Text1(1).Text) & "," & Month(Text1(1).Text) & "," & Day(Text1(1).Text) & ")"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "fecfactu = " & DBSet(Text1(1).Text, "F")
    If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    
    cadParam = cadParam & "pDuplicado=1|"
    numParam = numParam + 1
    
    If indRPT = 23 And (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16) Then
        Dim PrecioApor As Double
        PrecioApor = DevuelveValor("select min(precio) from raporreparto")
        
        cadParam = cadParam & "pPrecioApor=""" & Replace(Format(PrecioApor, "#0.000000"), ",", ".") & """|"
        numParam = numParam + 1
    End If
    
    '[Monica]28/01/2014: preguntamos si quiere imprimir arrobas
    If vParamAplic.Cooperativa = 12 Or vParamAplic.Cooperativa = 18 Then
        If MsgBox("¿ Desea impresión con Arrobas ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            cadParam = cadParam & "pConArrobas=1|"
        Else
            cadParam = cadParam & "pConArrobas=0|"
        End If
        numParam = numParam + 1
    End If
    
    '[Monica]10/02/2016: preguntamos si quiere imprimir el detalle de los campos
    If vParamAplic.Cooperativa = 4 Then
        If Text1(12).Text = "FAA" Then
            If MsgBox("¿ Desea impresión detallada por campos ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                cadParam = cadParam & "pDetalle=1|"
            Else
                cadParam = cadParam & "pDetalle=0|"
            End If
            numParam = numParam + 1
        End If
    End If
    
    
    If Not HayRegParaInforme(NombreTabla, cadselect) Then Exit Sub
     
    With frmImprimir
          '[Monica]06/02/2012: añadido la siguientes 3 lineas para el envio por el outlook
            .outClaveNombreArchiv = Mid(Combo1(0).Text, 1, 3) & Format(Text1(0).Text, "0000000")
            .outCodigoCliProv = Text1(2).Text
            .outTipoDocumento = 100
    
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Factura de Socios"
            .ConSubInforme = True
            .Show vbModal
    End With

    If frmVisReport.EstaImpreso Then
        ActualizarRegistros "rfactsoc", cadselect
    End If
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
Dim cadMen As String

    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 14 ' codigo de concepto de gasto
            If txtAux3(Index) <> "" Then
                txtAux3(15).Text = DevuelveDesdeBDNew(cAgro, "rconcepgasto", "nomgasto", "codgasto", txtAux3(14), "N")
                If txtAux3(15).Text = "" Then
                    cadMen = "No existe el Concepto de Gasto: " & txtAux3(15).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmGas = New frmManConcepGasto
                        frmGas.DatosADevolverBusqueda = "0|1|"
                        frmGas.NuevoCodigo = txtAux3(14).Text
                        TerminaBloquear
                        frmGas.Show vbModal
                        Set frmGas = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        PonerFoco txtAux3(Index)
                    Else
                        txtAux3(Index).Text = ""
                    End If
                    PonerFoco txtAux3(Index)
                Else
                    '[Monica]20/07/2016: el gasto puede no ser de factura
                    If vParamAplic.Cooperativa <> 0 Then
                        If Not EsGastodeFactura(txtAux3(Index).Text) = True Then
                            MsgBox "Este concepto de gasto no es de factura. Reintroduzca.", vbExclamation
                            PonerFoco txtAux3(Index)
                        End If
                    End If
                End If
            Else
                txtAux3(15).Text = ""
            End If
    
        Case 16 ' importe
            If txtAux3(Index) <> "" Then
                If PonerFormatoDecimal(txtAux3(Index), 3) Then CmdAceptar.SetFocus
            End If
        
    End Select
    
    
    
End Sub

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de factura
    Sql = "select codtipom, nomtipom from usuarios.stipom where tipodocu > 0 and tipodocu < 12"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 1
    While Not Rs.EOF
'        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
        Sql = Rs.Fields(1).Value
        Sql = Rs.Fields(0).Value & " - " & Sql
        Combo1(0).AddItem Sql 'campo del codigo
        Combo1(0).ItemData(Combo1(0).NewIndex) = i
        i = i + 1
        Rs.MoveNext
    Wend
    
    'tipo de IRPF
    Combo1(1).AddItem "Módulos"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "E.D."
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Entidad"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    
    
    'estado de la factura con respecto al nro de factura
    Combo1(3).AddItem "Normal"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 0
    Combo1(3).AddItem "Pdte.Recibir"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 1
    Combo1(3).AddItem "Recibido"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 2
    
    
    

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
Dim Sql As String
Dim cadena As String


    On Error GoTo EModificarCab

    conn.BeginTrans
    
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    
    cadena = ""
    If Text1(1).Text <> FecAnterior Then cadena = cadena & " por " & Text1(1).Text
    If SocioAnterior <> Text1(2).Text Then cadena = cadena & " Soc " & SocioAnterior & " por " & Text1(2).Text
    If IRPFAnterior <> Combo1(1).ListIndex Then cadena = cadena & " IRPF " & IRPFAnterior & " por " & Combo1(1).ListIndex
    If IvaAnterior <> Text1(3).Text Then cadena = cadena & " Iva Ant " & IvaAnterior & " por " & Text1(3).Text
    If ObsAnterior <> Text1(23).Text Then cadena = cadena & " Obs Ant " & ObsAnterior & " por " & Text1(23).Text

    If cadena <> "" Then
        cadena = Text1(12).Text & " " & Text1(0).Text & " " & FecAnterior & cadena
    
        LOG.Insertar 12, vUsu, "Modificación Cabecera: " & cadena & vbCrLf
    End If
    
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    '[Monica]25/07/2013: cambio de la fecha de la factura
    If Text1(1).Text <> FecAnterior Then
        Sql = "update rfactsoc set fecfactu = " & DBSet(Text1(1).Text, "F")
        Sql = Sql & " where codtipom= '" & Text1(12).Text & "'"
        Sql = Sql & " and numfactu = " & Text1(0).Text
        Sql = Sql & " and fecfactu = " & DBSet(FecAnterior, "F")
    
        conn.Execute Sql
    End If
    
    b = ModificaDesdeFormulario2(Me, 2, "Frame2")
    If b Then b = ModificaDesdeFormulario2(Me, 2, "Frame6")

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
Dim Sql As String

    On Error GoTo EInsertarCab
    
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
                Set frmLFac = New frmManLinFactSocios
                
                frmLFac.ModoExt = 3
                frmLFac.tipoMov = Text1(12).Text
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


Private Sub InsertarLinea(Index As Integer)
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case Index
        Case 1: nomframe = "FrameGastos" 'clasificacion
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
            b = BloqueaRegistro("rfactsoc", "numfactu = " & Data1.Recordset!numfactu)
            CargaGrid DataGrid4, Data5, True
            
            CalcularGastos
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
       
    NumTabMto = Index
    If Index = 2 Then NumTabMto = 3
    
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case NumTabMto
        Case 1: vtabla = "rfactsoc_gastos"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case NumTabMto
        Case 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
'             *** canviar la clau primaria de les llínies,
'            pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
'             ***************************************************************

            AnyadirLinea DataGrid4, Data5

            anc = DataGrid4.Top
            If DataGrid4.Row < 0 Then
                anc = anc + 220
            Else
                anc = anc + DataGrid4.RowTop(DataGrid4.Row) + 5
            End If

            LLamaLineas ModificaLineas, anc, "DataGrid4"

            LimpiarCamposLin "FrameGastos"

            txtAux3(10).Text = Mid(Combo1(0).Text, 1, 3) 'tipo de movimiento
            txtAux3(11).Text = Text1(0).Text 'nro de factura
            txtAux3(12).Text = Text1(1).Text 'fecha de factura
            txtAux3(13).Text = NumF ' nro. de linea
            txtAux3(15).Text = ""

'            BloquearTxt txtAux(14), False
'                    BloquearTxt txtaux(12), False
            PonerFoco txtAux3(14)


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
            
            CalcularGastos
            PonerCampos
  
            V = Data5.Recordset.Fields(3) 'el 2 es el nº de llinia
            CargaGrid DataGrid4, Data5, True

            ' *** si n'hi han tabs ***
'                    SSTab1.Tab = 1

            DataGrid4.SetFocus
            Data5.Recordset.Find (Data5.Recordset.Fields(3).Name & " =" & V)

            LLamaLineas ModificaLineas, 0, "DataGrid4"
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
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function

    Select Case nomframe
        Case "FrameGastos"
            If txtAux3(16).Text = "" Then
                MsgBox "Debe introducir un importe. Reintroduzca.", vbExclamation
                PonerFoco txtAux3(16)
            End If
    
    End Select
    
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
    vWhere = vWhere & " codtipom = '" & Mid(Combo1(0).Text, 1, 3) & "' and numfactu = " & Val(Text1(0).Text) & " and fecfactu = " & DBSet(Text1(1).Text, "F")
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



Private Sub CalcularTotales()
Dim Importe  As Currency
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Anticipos As Currency

'15/01/2015: los anticipos van por variedad campo. CORREGIDO
    Sql = "select codvarie, codcampo, sum(imporvar) from rfactsoc_variedad where codtipom = " & DBSet(Data1.Recordset.Fields(0).Value, "T")
    Sql = Sql & " and numfactu = " & Data1.Recordset.Fields(1).Value
    Sql = Sql & " and fecfactu = " & DBSet(Data1.Recordset.Fields(2).Value, "F")
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Importe = 0
    While Not Rs.EOF
        Importe = Importe + DBLet(Rs.Fields(2).Value, "N") 'Solo es para saber que hay registros que mostrar
        
        Sql = "select sum(imporvar) from rfactsoc_variedad where (codtipom, numfactu, fecfactu) in (select codtipomanti,numfactuanti,fecfactuanti from rfactsoc_anticipos where codtipom = " & DBSet(Data1.Recordset.Fields(0).Value, "T")
        Sql = Sql & " and numfactu = " & Data1.Recordset.Fields(1).Value
        Sql = Sql & " and fecfactu = " & DBSet(Data1.Recordset.Fields(2).Value, "F") & ")"
        Sql = Sql & " and codvarie = " & DBSet(Rs.Fields(0).Value, "N")
        Sql = Sql & " and codcampo = " & DBSet(Rs.Fields(1).Value, "N")
    
        Anticipos = DevuelveValor(Sql)
        
        Importe = Importe - Anticipos
        
        Rs.MoveNext
    Wend
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
    End Select
    
    'solo en el caso de que estemos insertando y modificando y no haya % de retencion
    'le daremos el que hay en parametros
    If Text1(8).Text = "" And Combo1(1).ListIndex <> 2 And (Modo = 3 Or Modo = 4) Then
        Text1(8).Text = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
    End If
    
    ' calculo de la retencion
    PorRet = CCur(ComprobarCero(Text1(8).Text))
    ImpRet = Round2(BaseReten * PorRet / 100, 2)
    
    'solo si es liquidacion (normal o venta campo) o siniestro se calcula la aportación
    If EsFacturaLiquidacion(Text1(12).Text) Then
        ' calculo de la aportacion de fondo operativo
        BaseAFO = Base
        PorAFO = CCur(ComprobarCero(Text1(13).Text))
        ImpAFO = Round2(BaseAFO * PorAFO / 100, 2)
    
        TotFac = Base + impiva - ImpRet - ImpAFO
    Else
        TotFac = Base + impiva - ImpRet
    End If

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
    
    If BaseAFO = 0 Then
        Text1(14).Text = ""
    Else
        Text1(14).Text = Format(BaseAFO, "###,##0.00")
    End If
    
    If ImpAFO = 0 Then
        Text1(10).Text = ""
    Else
        Text1(10).Text = ImpAFO
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

Private Sub VisualizarAnticipos()
Dim vTipoMov As CTiposMov

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(Trim(Text1(12).Text)) Then
        Select Case vTipoMov.TipoDocu
            Case 1, 3, 5, 6, 7, 9
                ' 1=anticipos   3=anticipos vc  5=subvenciones   6=siniestros
                ' 7=ant.almaz   9=ant.bodega
                cmdAnticipos.Enabled = False
                cmdAnticipos.visible = False
                FrameAnticipos.visible = True
                cmdAnticipos_Click
'                Me.SSTab1.TabEnabled(3) = False

            Case 2, 4, 8, 10
                ' 2=liquidaciones, 4=liquidaciones vc, 8=liquidaciones almaz, 10=liquid.bodega
                cmdAnticipos.Enabled = True
                cmdAnticipos.visible = True
'                Me.SSTab1.TabEnabled(3) = True
        End Select
    End If
        
End Sub



Private Sub CalcularGastos()
Dim ImporteGastos  As Currency
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim ImporteFVarias As Currency

    Sql = "select sum(importe) from rfactsoc_gastos where codtipom = " & DBSet(Data1.Recordset.Fields(0).Value, "T")
    Sql = Sql & " and numfactu = " & Data1.Recordset.Fields(1).Value
    Sql = Sql & " and fecfactu = " & DBSet(Data1.Recordset.Fields(2).Value, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ImporteGastos = 0
    If Not Rs.EOF Then                 '[Monica]09/07/2013: añadida la condicion de montifrut pq los dtos tienen iva
        If Rs.Fields(0).Value <> 0 And vParamAplic.Cooperativa <> 12 Then ImporteGastos = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing


    '[Monica]26/02/2013: calculamos el importe de facturas varias que se van a descontar del importe de factura
    Sql = "select sum(totalfac) from fvarcabfact where (codtipom, numfactu, fecfactu) in "
    Sql = Sql & " (select codtipomfvar, numfactufvar, fecfactufvar from rfactsoc_fvarias where codtipom = " & DBSet(Data1.Recordset.Fields(0).Value, "T")
    Sql = Sql & " and numfactu = " & Data1.Recordset.Fields(1).Value
    Sql = Sql & " and fecfactu = " & DBSet(Data1.Recordset.Fields(2).Value, "F") & ")"
        
    ImporteFVarias = DevuelveValor(Sql)
    ImporteGastos = ImporteGastos + ImporteFVarias
    
    Text1(15).Text = Format(ImporteGastos, "###,##0.00")
    Text1(16).Text = Format(CCur(ImporteSinFormato(Text1(7).Text)) - CCur(ImporteSinFormato(Text1(15).Text)), "###,##0.00")
    
    DoEvents
    
End Sub



'******************************************
'********** FACTURAS VARIAS
'**********
Private Sub TxtAux5_GotFocus(Index As Integer)
    ConseguirFoco txtAux4(Index), Modo
End Sub

Private Sub TxtAux5_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux5_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux5_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux5(Index), Modo) Then Exit Sub
    
    
End Sub

Private Function AsignarAlbaranes() As Boolean
Dim Sql As String
Dim cadWhere As String


    On Error GoTo eAsignaAlbaranes

    AsignarAlbaranes = False

    Sql = "select * from rhisfruta where codsocio = " & DBSet(Text1(2).Text, "N")
    Sql = Sql & " and rhisfruta.codvarie in (select distinct codvarie from rfactsoc_variedad where codtipom = " & DBSet(Data1.Recordset!CodTipom, "T")
    Sql = Sql & " and numfactu = " & Data1.Recordset!numfactu & " and fecfactu = " & DBSet(Data1.Recordset!fecfactu, "F") & ")"
    
    cadWhere = "codsocio = " & DBSet(Text1(2).Text, "N")
    cadWhere = cadWhere & " and rhisfruta.codvarie in (select distinct codvarie from rfactsoc_variedad where codtipom = " & DBSet(Data1.Recordset!CodTipom, "T")
    cadWhere = cadWhere & " and numfactu = " & Data1.Recordset!numfactu & " and fecfactu = " & DBSet(Data1.Recordset!fecfactu, "F") & ")"
    
    If TotalRegistrosConsulta(Sql) <> 0 Then
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 47
        frmMens.cadWhere = cadWhere
        
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


Private Function InsertarAlbaranes(albaranes As String)
Dim Sql As String
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

    On Error GoTo eInsertarAlbaranes

    InsertarAlbaranes = False
    
    conn.BeginTrans
    
    vSQL = "select codvarie, sum(kilosnet) kilosnet from rhisfruta where numalbar in ( " & albaranes & ") group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    While Not Rs.EOF
        TotalKilos = DBLet(Rs!KilosNet, "N")
        ImporteVar = DevuelveValor("select imporvar from rfactsoc_variedad where " & ObtenerWhereCP(False) & " and codvarie = " & DBSet(Rs!codvarie, "N"))
    
        Sql2 = "select * from rhisfruta where numalbar in (" & albaranes & ") and codvarie = " & DBSet(Rs!codvarie, "N")
        Set Rs2 = New ADODB.Recordset
        
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs2.EOF
            ImporteAlb = Round2(ImporteVar * Rs2!KilosNet / TotalKilos, 2)
            PrecioAlb = Round2(ImporteAlb / Rs2!KilosNet, 4)
            
            CadValues = CadValues & "(" & DBSet(Text1(12).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "F") & ","
            CadValues = CadValues & DBSet(Rs2!NumAlbar, "N") & "," & DBSet(Rs2!Fecalbar, "F") & "," & DBSet(Rs2!codvarie, "N") & ",0," & DBSet(Rs2!KilosBru, "N") & ","
            CadValues = CadValues & DBSet(Rs2!KilosNet, "N") & ",0," & DBSet(PrecioAlb, "N") & "," & DBSet(ImporteAlb, "N") & ",0,0,0,0),"
            
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
        
        'actualizamos rfactsoc_variedad
        PrecioAlb = Round2(ImporteVar / TotalKilos, 4)
        
        Sql = "update rfactsoc_variedad set kilosnet = " & DBSet(TotalKilos, "N")
        Sql = Sql & ", preciomed = " & DBSet(PrecioAlb, "N")
        Sql = Sql & " where " & ObtenerWhereCP(False)
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        
        conn.Execute Sql
        
        'actualizamos rfactsoc_calidad
        Sql = "update rfactsoc_calidad set kilosnet = " & DBSet(TotalKilos, "N")
        Sql = Sql & ", precio = " & DBSet(PrecioAlb, "N")
        Sql = Sql & " where " & ObtenerWhereCP(False)
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        
        conn.Execute Sql
        
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
    
        ' igual que el insert pero reemplaza las columnas existentes
        Sql = "replace into rfactsoc_albaran (codtipom,numfactu,fecfactu,numalbar,fecalbar,codvarie,codcampo,kilosbru,"
        Sql = Sql & "kilosnet,grado,precio,importe,imporgasto,prretirada,prmoltura,prenvasado) values "
    
        conn.Execute Sql & CadValues
        
    End If
    conn.CommitTrans
    InsertarAlbaranes = True
    Exit Function
    
eInsertarAlbaranes:
    conn.RollbackTrans
    MuestraError Err.Number, "Insertando Albaranes", Err.Description
End Function




Private Function InsertarAlbaranesFrutasInma(albaranes As String) As Boolean
Dim Sql As String
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

    On Error GoTo eInsertarAlbaranes

    InsertarAlbaranesFrutasInma = False
    
    conn.BeginTrans
    
    vSQL = "select codvarie, codcampo, sum(kilostra) kilostra from rhisfruta where numalbar in ( " & albaranes & ") group by 1, 2 order by 1, 2"
    
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    While Not Rs.EOF
        TotalKilos = DBLet(Rs!KilosTra, "N")
        ImporteVar = DevuelveValor("select imporvar from rfactsoc_variedad where " & ObtenerWhereCP(False) & " and codvarie = " & DBSet(Rs!codvarie, "N") & " and codcampo = " & DBSet(Rs!codCampo, "N"))
    
        Sql2 = "select * from rhisfruta where numalbar in (" & albaranes & ") and codvarie = " & DBSet(Rs!codvarie, "N") & " and codcampo = " & DBSet(Rs!codCampo, "N")
        Set Rs2 = New ADODB.Recordset
        
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs2.EOF
            ImporteAlb = Round2(ImporteVar * Rs2!KilosTra / TotalKilos, 2)
            PrecioAlb = Round2(ImporteAlb / Rs2!KilosTra, 4)
            
            CadValues = CadValues & "(" & DBSet(Text1(12).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "F") & ","
            CadValues = CadValues & DBSet(Rs2!NumAlbar, "N") & "," & DBSet(Rs2!Fecalbar, "F") & "," & DBSet(Rs2!codvarie, "N") & ",0," & DBSet(Rs2!KilosBru, "N") & ","
            CadValues = CadValues & DBSet(Rs2!KilosTra, "N") & ",0," & DBSet(PrecioAlb, "N") & "," & DBSet(ImporteAlb, "N") & ",0,0,0,0),"
            
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
        
        'actualizamos rfactsoc_variedad
        PrecioAlb = Round2(ImporteVar / TotalKilos, 4)
        
        Sql = "update rfactsoc_variedad set kilosnet = " & DBSet(TotalKilos, "N")
        Sql = Sql & ", preciomed = " & DBSet(PrecioAlb, "N")
        Sql = Sql & " where " & ObtenerWhereCP(False)
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        Sql = Sql & " and codcampo = " & DBSet(Rs!codCampo, "N")
        
        conn.Execute Sql
        
        'faltaria introducir las calidades pq en la factura pq en frutas inma si que lo saben
        Sql = "replace rfactsoc_calidad (codtipom,numfactu,fecfactu,codvarie,codcampo,codcalid,kilosnet,precio,imporcal,preciocalidad,imporcalidad) "
        Sql = Sql & " select " & DBSet(Text1(12), "T") & "," & DBSet(Text1(0), "N") & "," & DBSet(Text1(1), "F") & ", cc.codvarie, cc.codcampo, ll.codcalid, sum(ll.kilostra), " & DBSet(PrecioAlb, "N") & ", round(sum(ll.kilostra) * " & DBSet(PrecioAlb, "N") & ",2),"
        Sql = Sql & DBSet(PrecioAlb, "N") & ", round(sum(ll.kilostra) * " & DBSet(PrecioAlb, "N") & ",2)"
        Sql = Sql & " from rhisfruta_clasif ll inner join rhisfruta cc on ll.numalbar = cc.numalbar "
        Sql = Sql & " where cc.numalbar in (" & albaranes & ") "
        Sql = Sql & " group by 1,2,3,4,5,6,8,10"
        Sql = Sql & " order by 1,2,3,4,5,6,8,10"
        conn.Execute Sql
        
        'actualizamos rfactsoc_calidad
'        Sql = "update rfactsoc_calidad set kilosnet = " & DBSet(TotalKilos, "N")
'        Sql = Sql & ", precio = " & DBSet(PrecioAlb, "N")
'        Sql = Sql & " where " & ObtenerWhereCP(False)
'        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
'
'        conn.Execute Sql
        
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
    
        Sql = "delete from rfactsoc_albaran where  " & ObtenerWhereCP(False)
        conn.Execute Sql
    
        ' igual que el insert pero reemplaza las columnas existentes
        Sql = "replace into rfactsoc_albaran (codtipom,numfactu,fecfactu,numalbar,fecalbar,codvarie,codcampo,kilosbru,"
        Sql = Sql & "kilosnet,grado,precio,importe,imporgasto,prretirada,prmoltura,prenvasado) values "
    
        conn.Execute Sql & CadValues
        
    End If
    conn.CommitTrans
    InsertarAlbaranesFrutasInma = True
    Exit Function
    
eInsertarAlbaranes:
    conn.RollbackTrans
    MuestraError Err.Number, "Insertando Albaranes", Err.Description
End Function




