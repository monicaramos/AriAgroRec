VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManFactTranspor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Transportistas"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   12225
   Icon            =   "frmManFactTranspor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   45
      TabIndex        =   68
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   69
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3690
      TabIndex        =   66
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   67
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
      Left            =   9945
      TabIndex        =   65
      Top             =   315
      Width           =   1605
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4875
      Left            =   45
      TabIndex        =   37
      Top             =   3915
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   8599
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
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
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   1080
         Width           =   3330
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
         Index           =   12
         Left            =   -71130
         MaxLength       =   7
         TabIndex        =   55
         Tag             =   "Nº Factura Rectificada|N|S|||rfacttra|rectif_numfactu|0000000|N|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   1110
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
         Left            =   -69765
         MaxLength       =   10
         TabIndex        =   54
         Tag             =   "Fecha Factura Rectificada|F|S|||rfacttra|rectif_fecfactu|dd/mm/yyyy|N|"
         Top             =   1080
         Width           =   1350
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
         Height          =   555
         Index           =   14
         Left            =   -74520
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   53
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
         Height          =   4125
         Left            =   60
         TabIndex        =   38
         Top             =   540
         Width           =   11880
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
            Left            =   4455
            MaskColor       =   &H00000000&
            TabIndex        =   40
            ToolTipText     =   "Buscar variedad"
            Top             =   1170
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux2 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
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
            Height          =   315
            Index           =   5
            Left            =   4545
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   71
            Text            =   "Text2"
            Top             =   1170
            Width           =   2235
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
            Left            =   1575
            MaxLength       =   7
            TabIndex        =   64
            Tag             =   "Transportista|T|N|||rfacttra_albaran|codtrans||S|"
            Text            =   "transpo"
            Top             =   1170
            Visible         =   0   'False
            Width           =   405
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
            Index           =   1
            Left            =   2835
            MaskColor       =   &H00000000&
            TabIndex        =   39
            ToolTipText     =   "Buscar albaran"
            Top             =   1170
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   9270
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
            BackColor       =   &H00FFFFFF&
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
            Left            =   8505
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
            BackColor       =   &H00FFFFFF&
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
            Left            =   7605
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
            BackColor       =   &H00FFFFFF&
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
            Left            =   6840
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
            BackColor       =   &H00FFFFFF&
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
            Left            =   2070
            MaxLength       =   7
            TabIndex        =   44
            Tag             =   "Albaran|N|N|||rfacttra_albaran|numalbar|0000000|S|"
            Text            =   "albaran"
            Top             =   1170
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   3060
            MaxLength       =   10
            TabIndex        =   45
            Tag             =   "Fecha Alb|F|N|||rfacttra_albaran|fecalbar|dd/mm/yyyy||"
            Text            =   "fecalbar"
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
            Index           =   0
            Left            =   315
            MaxLength       =   7
            TabIndex        =   43
            Tag             =   "Tipo Movim.|T|N|||rfacttra_albaran|codtipom||S|"
            Text            =   "tipof"
            Top             =   1170
            Visible         =   0   'False
            Width           =   345
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
            Left            =   720
            MaxLength       =   7
            TabIndex        =   42
            Tag             =   "Nº.Factura|N|N|||rfacttra_albaran|numfactu|0000000|S|"
            Text            =   "numfact"
            Top             =   1170
            Visible         =   0   'False
            Width           =   360
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
            Left            =   1170
            MaxLength       =   10
            TabIndex        =   41
            Tag             =   "Fecha Fact.|F|N|||rfacttra_albaran|fecfactu|dd/mm/yyyy|S|"
            Text            =   "FecFact"
            Top             =   1170
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   3870
            MaxLength       =   25
            TabIndex        =   46
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
            TabIndex        =   51
            Top             =   30
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
            Height          =   3495
            Left            =   180
            TabIndex        =   52
            Top             =   495
            Width           =   11625
            _ExtentX        =   20505
            _ExtentY        =   6165
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
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   -74250
         MaxLength       =   10
         TabIndex        =   63
         Tag             =   "Tipo Movimiento|T|S|||rfacttra|rectif_codtipom||N|"
         Top             =   1080
         Width           =   795
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
         Left            =   -74550
         TabIndex        =   60
         Top             =   840
         Width           =   2070
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
         Left            =   -69765
         TabIndex        =   59
         Top             =   840
         Width           =   1050
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
         Left            =   -71130
         TabIndex        =   58
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   -68700
         Picture         =   "frmManFactTranspor.frx":00BE
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
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
         Left            =   -74520
         TabIndex        =   57
         Top             =   1590
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   45
      TabIndex        =   20
      Top             =   900
      Width           =   12045
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
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   225
         Width           =   1980
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   10
         Left            =   300
         MaxLength       =   10
         TabIndex        =   62
         Tag             =   "Tipo Movimiento|T|N|||rfacttra|codtipom||S|"
         Top             =   240
         Width           =   855
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
         Left            =   9870
         TabIndex        =   6
         Tag             =   "Aridoc|N|N|0|1|rfacttra|pasaridoc|0||"
         Top             =   720
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
         Index           =   5
         Left            =   240
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Base Imponible|N|N|||rfacttra|baseimpo|###,##0.00||"
         Top             =   1770
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
         Left            =   9585
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "Total Factura|N|N|||rfacttra|totalfac|###,##0.00||"
         Top             =   1770
         Width           =   1830
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
         Left            =   6765
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Importe Iva|N|N|||rfacttra|imporiva|###,##0.00||"
         Top             =   1770
         Width           =   1830
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
         Left            =   1845
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
         Left            =   2475
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   1770
         Width           =   3450
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
         Left            =   5985
         MaxLength       =   6
         TabIndex        =   9
         Tag             =   "Porc.Iva|N|N|||rfacttra|porc_iva|##0.00||"
         Text            =   "123"
         Top             =   1770
         Width           =   645
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
         Left            =   5985
         MaxLength       =   6
         TabIndex        =   12
         Tag             =   "Porc.Retención|N|S|||rfacttra|porc_ret|##0.00||"
         Text            =   "123"
         Top             =   2340
         Width           =   645
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
         Left            =   6765
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Importe Retención|N|S|||rfacttra|impreten|#,##0.00||"
         Text            =   "123"
         Top             =   2340
         Width           =   1830
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
         TabIndex        =   11
         Tag             =   "Base Retención|N|S|||rfacttra|basereten|###,##0.00||"
         Top             =   2340
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
         Left            =   5985
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Tipo IRPF|N|N|0|3|rfacttra|tipoirpf|0|N|"
         Top             =   855
         Width           =   2580
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
         Left            =   9870
         TabIndex        =   5
         Tag             =   "Contabilizado|N|N|0|1|rfacttra|contabilizado|0||"
         Top             =   450
         Width           =   1890
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
         Left            =   150
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Cod.Transportista|T1|N|||rfacttra|codtrans||S|"
         Text            =   "Text1"
         Top             =   855
         Width           =   1290
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
         Index           =   0
         Left            =   2205
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Factura|N|S|||rfacttra|numfactu|0000000|S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1110
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
         Left            =   9870
         TabIndex        =   4
         Tag             =   "Impreso|N|N|0|1|rfacttra|impreso|0||"
         Top             =   180
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
         Index           =   1
         Left            =   3390
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||rfacttra|fecfactu|dd/mm/yyyy|S|"
         Top             =   240
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
         Left            =   1485
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   855
         Width           =   4455
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
         Left            =   45
         TabIndex        =   22
         Top             =   1245
         Width           =   11835
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
            TabIndex        =   36
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
            Left            =   9630
            TabIndex        =   35
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
            Left            =   6705
            TabIndex        =   34
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
            Left            =   1800
            TabIndex        =   33
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
            Left            =   5925
            TabIndex        =   32
            Top             =   270
            Width           =   630
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
            Left            =   5925
            TabIndex        =   31
            Top             =   855
            Width           =   630
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
            Left            =   6705
            TabIndex        =   30
            Top             =   855
            Width           =   1350
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
            TabIndex        =   29
            Top             =   840
            Width           =   1860
         End
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
         Left            =   5985
         TabIndex        =   27
         Top             =   585
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
         Left            =   165
         TabIndex        =   26
         Top             =   0
         Width           =   1350
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
         Left            =   3420
         TabIndex        =   25
         Top             =   0
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Transportista"
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
         Left            =   165
         TabIndex        =   24
         Top             =   585
         Width           =   1335
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         ToolTipText     =   "Buscar Socio"
         Top             =   615
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
         Index           =   28
         Left            =   2205
         TabIndex        =   23
         Top             =   0
         Width           =   1125
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   4500
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
      Top             =   8850
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
         TabIndex        =   19
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
      Left            =   11025
      TabIndex        =   16
      Top             =   8955
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
      Left            =   9855
      TabIndex        =   15
      Top             =   8940
      Width           =   1065
   End
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
      Left            =   11025
      TabIndex        =   17
      Top             =   8955
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3870
      Top             =   7785
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
   Begin MSAdodcLib.Adodc Data5 
      Height          =   360
      Left            =   900
      Top             =   8145
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
      Top             =   7290
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   11625
      TabIndex        =   70
      Top             =   210
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

Public NumAlbar As String  ' venimos de pedidos para insertar envases paletizacion

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBasico2 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmLFac As frmManLinFactTranspor 'Lineas de variedades de facturas transporte
Attribute frmLFac.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de iva de conta
Attribute frmTIva.VB_VarHelpID = -1

Private WithEvents frmVar As frmBasico2 'Form Mto de variedades de comercial
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

Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim vTrans As CTransportista ' clase de transportista



Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'variedad
            
            Set frmVar = New frmBasico2
            
            AyudaVariedad frmVar
            
            Set frmVar = Nothing
            
            PonerFoco txtAux1(5)
        
        Case 1 'albaranes del transportista pendientes de facturar
            Set frmMens = New frmMensajes
            
            frmMens.OpcionMensaje = 24
            frmMens.cadWhere = "codtrans = " & DBSet(Data1.Recordset!codTrans, "T")
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
Dim vTipoMov As CTiposMov

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
    Text1(8).Text = vParamAplic.PorcreteFacTra
    
    LimpiarDataGrids
    Combo1(0).ListIndex = 0
    Text1(10).Text = Mid(Trim(Combo1(0).List(0)), 1, 3)
'    Combo1(0).SetFocus
    
    PonerFoco Text1(1)
    
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
            anc = anc + 240 '440
        Else
            anc = anc + DataGrid5.RowTop(DataGrid5.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid5"
        
        
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

        For J = 0 To 2
            txtAux1(J).Text = DataGrid5.Columns(J).Text
        Next J

        txtAux1(10).Text = DataGrid5.Columns(11).Text
        

        For J = 3 To 5
            txtAux1(J).Text = DataGrid5.Columns(J).Text
        Next J
        txtAux2(5).Text = DataGrid5.Columns(6).Text
        For J = 6 To 9
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
            b = (xModo = 1 Or xModo = 2)
            For jj = 3 To 9
                txtAux1(jj).Height = DataGrid5.RowHeight - 10
                txtAux1(jj).Top = alto + 5
                txtAux1(jj).visible = b
            Next jj
            
            If xModo = 2 Then
                txtAux1(3).visible = False
                txtAux1(7).visible = False
            End If
            
            txtAux2(5).Height = DataGrid5.RowHeight - 10
            txtAux2(5).Top = alto + 5
            txtAux2(5).visible = b
            
            For jj = 0 To 1
                btnBuscar(jj).Height = DataGrid5.RowHeight - 10
                btnBuscar(jj).Top = alto + 5
                btnBuscar(jj).visible = b
            Next jj
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
'    btnPrimero = 13
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
'        .Buttons(10).Image = 11  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With
    
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


Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(10), CadenaSeleccion, 1)
        cadB = cadB & Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 2)
        cadB = cadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 3)
        cadB = cadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaSeleccion, 4)
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

' devolvemos la linea del datagrid en donde estabamos
Private Sub frmLFac_DatoSeleccionado(CadenaSeleccion As String)
'Dim vWhere As String
'
'   PonerCamposLineas
'
'   If CadenaSeleccion = "" Then Exit Sub
'
'   vWhere = "(codtipom = '" & RecuperaValor(CadenaSeleccion, 1) & "' and numfactu = " & RecuperaValor(CadenaSeleccion, 2) & " and fecfactu = " & RecuperaValor(CadenaSeleccion, 3)
'
'   SituarDataMULTI Data3, vWhere, "" ', Indicador
'
'   PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
'   PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
'

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
    txtAux1(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Variedad
    txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Variedad
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
                    
'                    '[Monica]14/02/2019: dejamos modificar el nro de factura
'                    If Modo = 3 Then
'                        '[Monica]14/02/2019: pueden tocar el contador
'                        If vParamAplic.TipoContadorTRA = 0 Then ' contador automatico de FTR facturas de transporte
'                            Set vTipoMov = New CTiposMov
'                            If vTipoMov.Leer(CodTipoMov) Then
'                                Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
'                            End If
'                        Else
'                            Set vTrans = New CTransportista
'                            If vTrans.LeerDatos(Text1(2).Text) Then
'                                Text1(0).Text = vTrans.ConseguirContador()
'                            End If
'                        End If
'                        PonerFoco Text1(0)
'                    End If
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

    '[Monica]09/01/2015: nuevo tipo de dato para busqueda sin asteriscos
    Text1(2).Tag = "Cod.Transportista|TT|N|||rfacttra|codtrans||S|"
    cadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    Text1(2).Tag = "Cod.Transportista|T|N|||rfacttra|codtrans||S|"

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select rfacttra.* from " & NombreTabla & " LEFT JOIN rfacttra_albaran ON rfacttra.codtipom=rfacttra_albaran.codtipom "
        CadenaConsulta = CadenaConsulta & " and rfacttra_albaran.numfactu = rfacttra.numfactu and rfacttra_albaran.fecfactu = rfacttra.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB & " GROUP BY rfacttra.codtipom, rfacttra.numfactu, rfacttra.fecfactu, rfacttra.codtrans " & Ordenacion
'        CadenaConsulta = "select palets.* from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY palets.numpalet " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
'    cad = ""
''    Cad = Cad & "Tipo|if(rfactsoc.codtipom='FAA','Anticipo','Liquidación') as a|T||10·"
'    cad = cad & "Tipo Factura|stipom.nomtipom|T||22·"
'    cad = cad & "Tipo|rfacttra.codtipom|T||7·"
'    cad = cad & "Nº.Factura|rfacttra.numfactu|N||10·"
'    cad = cad & "Fecha|rfacttra.fecfactu|F||15·"
'    cad = cad & "Transportista|rfacttra.codtrans|N||15·"
'    cad = cad & "Nombre Transportista|rtransporte.nomtrans|T||30·"
'    tabla = NombreTabla & " INNER JOIN rtransporte ON rfacttra.codtrans=rtransporte.codtrans "
'    tabla = "(" & tabla & ") INNER JOIN usuarios.stipom stipom ON rfacttra.codtipom = stipom.codtipom "
'    Titulo = "Facturas Transportistas"
'    devuelve = "1|2|3|"
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vtabla = tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = devuelve
'        frmB.vTitulo = Titulo
'        frmB.vSelElem = 0
''        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
''        If EsCabecera Then
''            PonerCadenaBusqueda
''            Text1(0).Text = Format(Text1(0).Text, "0000000")
''        End If
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If
'    Screen.MousePointer = vbDefault
'

    Set frmB = New frmBasico2
    
    AyudaFrasTransporte frmB, , cadB
    
    Set frmB = Nothing


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
'            Text1(0).BackColor = vblightblue
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
'        LLamaLineas Modo, 0, "DataGrid5"
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
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1

        
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
'    If Modo = 3 Then
'        Text1(0).Enabled = True
'    Else
        BloquearTxt Text1(0), b, True
'    End If
         
    BloquearTxt Text1(2), (Modo <> 1) And Modo <> 3 '<= 2)
    
    For i = 0 To txtAux1.Count - 1
         BloquearTxt txtAux1(i), True
         txtAux1(i).visible = False
    Next
    
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 3 To txtAux1.Count - 1
        If i <> 10 Then
            txtAux1(i).visible = ((Modo = 1) Or (Modo = 5))
            BloquearTxt txtAux1(i), Not (Modo = 1 Or Modo = 5)
        End If
    Next i
    txtAux2(5).visible = ((Modo = 1) Or (Modo = 5))
    
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
'    BloquearTxt txtAux1(3), b
'    BloquearTxt txtAux1(7), b
    
    For i = 0 To btnBuscar.Count - 1
        BloquearBtn Me.btnBuscar(i), Not (Modo = 5 Or Modo = 1)
        btnBuscar(i).visible = (Modo = 5 Or Modo = 1)
    Next i
    
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
       
       
'    Check1(1).Enabled = ((Modo = 3 Or Modo = 4) And vUsu.Nivel = 0)
       
    ' ***************************
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
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
'[Monica]17/01/2019: añado la condicion de que sea tb alzira
        If vParamAplic.Cooperativa = 12 Or vParamAplic.Cooperativa = 4 Then
            '[Monica]20/06/2017: control de fechas que antes no estaba, solo para Montifrut que cuando integra no coge la fecha de recepcion
            ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(1).Text))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then
                    MsgBox MensajeFechaOkConta, vbExclamation
'                    If MsgBox("¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
                    Exit Function
                End If
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
            Cad = Cad & vbCrLf & "Nota: " & Data6.Recordset!NumNotac
            Cad = Cad & vbCrLf & "Transportista: " & Data6.Recordset!codTrans
            
            
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
                        CalcularTotales
                    Else
                        LimpiarCampos
                        PonerModo 0
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

    On Error GoTo ECargaGrid

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
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
         Case "DataGrid5" 'rfactsoc_albaran
'           SQL = "SELECT codtipom,numfactu,fecfactu,codtipomanti,numfactuanti,fecfactuanti,codvarieanti,codvarieanti,nomvarie,codcampoanti,baseimpo "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux1(3)|T|Albaran|1100|;S|btnBuscar(1)|B|||;S|txtAux1(4)|T|Fecha|1400|;S|txtAux1(5)|T|Codigo|1200|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|txtAux2(5)|T|Variedad|2300|;"
            tots = tots & "S|txtAux1(6)|T|Campo|1300|;S|txtAux1(7)|T|Nota|1050|;"
            tots = tots & "S|txtAux1(8)|T|Kilos Neto|1300|;"
            tots = tots & "S|txtAux1(9)|T|Importe|1400|;N||||0|;"
            
            arregla tots, DataGrid5, Me, 350
    
            DataGrid5.Columns(3).Alignment = dbgLeft
            DataGrid5.Columns(6).Alignment = dbgLeft
    
    
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub




Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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
        Case 3 ' albaran
            PonerFormatoEntero txtAux1(Index)

        Case 4 ' fecha
            PonerFormatoFecha txtAux1(Index)
            
        Case 5 ' variedad
            If PonerFormatoEntero(txtAux1(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux1(Index), "variedades", "nomvarie")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & txtAux1(Index).Text & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco txtAux1(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
            
        Case 6 ' campo
            PonerFormatoEntero txtAux1(Index)
            
        Case 7 ' nro de nota
            PonerFormatoEntero txtAux1(Index)
            
        Case 8 ' peso neto
            PonerFormatoEntero txtAux1(Index)

        Case 9 ' importe
            If PonerFormatoDecimal(txtAux1(Index), 3) Then PonerFocoBtn cmdAceptar
            
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
    
    Dim vLog As cLOG
    Set vLog = New cLOG
    vLog.Insertar 1, vUsu, "Factura de transporte, Nro:" & Text1(0).Text & " Fec:" & Text1(1) & " Transpor: " & Text1(2)
    Set vLog = Nothing
    
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
            Sql = Sql & " and codtrans = " & DBSet(Data6.Recordset.Fields(11), "T")
            
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
        Sql = Sql & " rfacttra_albaran.kilosnet, rfacttra_albaran.importe, rfacttra_albaran.codtrans "
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
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    
    '[Monica]15/01/2019: tienen que poder insertar, tb frutas inma
    'Añadir
    Toolbar1.Buttons(1).Enabled = (vParamAplic.Cooperativa = 4) 'False 'B
    Me.mnModificar.Enabled = (vParamAplic.Cooperativa = 4) 'False 'B
    
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (hcoCodMovim = "") And Not (Check1(0).Value = 1) And Not (Check1(1).Value = 1)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b And vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16
    Me.mnModificar.Enabled = b And vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Impresión de albaran
    Toolbar1.Buttons(8).Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
    Me.mnImprimir.Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
    

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 2) And Not Check1(1).Value = 1 And (vParamAplic.Cooperativa = 4)
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
        Case "FTR"
            indRPT = 49
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
    
    'Transportista
    devuelve = "{" & NombreTabla & ".codtrans}=""" & Trim(Text1(2).Text) & """"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "codtrans = '" & Trim(Text1(2).Text) & "'"
    If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    
    
    cadParam = cadParam & "pDuplicado=1|"
    numParam = numParam + 1
    
    
    If Not HayRegParaInforme(NombreTabla, cadselect) Then Exit Sub
     
    With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Factura de Transportistas"
            .ConSubInforme = True
            .Show vbModal
    End With

    If frmVisReport.EstaImpreso Then
        ActualizarRegistros "rfacttra", cadselect
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

    Dim vLog As cLOG
    Set vLog = New cLOG
    vLog.Insertar 18, vUsu, "Factura de transporte, Nro:" & Text1(0).Text & " Fec:" & Text1(1) & " Transpor: " & Text1(2)
    Set vLog = Nothing



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
Dim Sql As String

    On Error GoTo EInsertarCab
    
    
    If vParamAplic.TipoContadorTRA = 0 Then ' contador automatico de FTR facturas de transporte
        
        Set vTipoMov = New CTiposMov
        If vTipoMov.Leer(CodTipoMov) Then
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            Sql = CadenaInsertarDesdeForm(Me)
            If Sql <> "" Then
                If InsertarOferta(Sql) Then
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    PonerModo 2
                    'Ponerse en Modo Insertar Lineas
'                    BotonMtoLineas 0, "Variedades"
                    BotonAnyadirLinea 0
'                    Set frmLFac = New frmManLinFactTranspor
'
'                    frmLFac.ModoExt = 3
'                    frmLFac.tipoMov = Text1(10).Text
'                    frmLFac.Factura = Text1(0).Text
'                    frmLFac.fecfactu = Text1(1).Text
'
'                    frmLFac.Show vbModal
'
'                    Set frmLFac = Nothing
                    
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
                If InsertarOferta2(Sql) Then
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
 '                   PonerModo 2
                    'Ponerse en Modo Insertar Lineas
'                    BotonMtoLineas 0, "Variedades"
                    BotonAnyadirLinea 0
                    
'[Monica]18/01/2019: la linea la agrego directamente
'                    Set frmLFac = New frmManLinFactTranspor
'
'                    frmLFac.ModoExt = 3
'                    frmLFac.tipoMov = Text1(10).Text
'                    frmLFac.Factura = Text1(0).Text
'                    frmLFac.fecfactu = Text1(1).Text
'
'                    frmLFac.Show vbModal
'
'                    Set frmLFac = Nothing
'
'                    CalcularTotales
                End If
            End If
            Text1(0).Text = Format(Text1(0).Text, "0000000")
        End If
'
'
        Set vTrans = Nothing
    
    End If
    
    
    
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function InsertarOferta(vSQL As String) As Boolean
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
    
    Dim vLog As cLOG
    Set vLog = New cLOG
    vLog.Insertar 17, vUsu, "Factura de transporte, Nro:" & Text1(0).Text & " Fec:" & Text1(1) & " Transpor: " & Text1(2)
    Set vLog = Nothing
    
    
    
    
    MenError = "Error al actualizar el contador de la Factura."
'    If Text1(0).Text = vTipoMov.Contador Then vTipoMov.IncrementarContador (CodTipoMov)
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


Private Function InsertarOferta2(vSQL As String) As Boolean
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
    
    Dim vLog As cLOG
    Set vLog = New cLOG
    vLog.Insertar 17, vUsu, "Factura de transporte, Nro:" & Text1(0).Text & " Fec:" & Text1(1) & " Transpor: " & Text1(2)
    Set vLog = Nothing
    
    MenError = "Error al actualizar el contador de la Factura."
'    If CLng(ComprobarCero(Text1(0).Text)) = vTrans.Contador + 1 Then vTrans.IncrementarContador
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
            CalcularTotales
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
                anc = anc + 240
            Else
                anc = anc + DataGrid5.RowTop(DataGrid5.Row) + 5
            End If

            LLamaLineas ModificaLineas, anc, "DataGrid5"

            LimpiarCamposLin "FrameAlbaranes"

            txtAux1(0).Text = Mid(Combo1(0).Text, 1, 3) 'tipo de movimiento
            txtAux1(1).Text = Text1(0).Text 'nro de factura
            txtAux1(2).Text = Text1(1).Text 'fecha de factura
            txtAux1(10).Text = Text1(2).Text 'transportista
            
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
    nomframe = "FrameAlbaranes" 'gastos
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            ' ******************************************************
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModificaLineas = 0
            
            CalcularTotales
            PonerCampos
  
            V = Data5.Recordset.Fields(3) 'el 2 es el nº de llinia
            CargaGrid DataGrid5, Data6, True

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

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function

    Select Case nomframe
        Case "FrameAlbaranes"
            If txtAux1(6).Text = "" Then
                MsgBox "Debe introducir un campo. Reintroduzca.", vbExclamation
                PonerFoco txtAux1(6)
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

