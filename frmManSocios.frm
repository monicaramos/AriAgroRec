VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Socios"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   13875
   Icon            =   "frmManSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   13875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
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
      Left            =   9060
      TabIndex        =   174
      Top             =   240
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5310
      TabIndex        =   172
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   173
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
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3900
      TabIndex        =   170
      Top             =   30
      Width           =   1335
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   171
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Baja de Socios"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir Fases"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   240
      TabIndex        =   167
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   168
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
               Object.Tag             =   "0"
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   66
      Top             =   780
      Width           =   13485
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
         Index           =   30
         Left            =   11250
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "C�digo Socio|N|S|||rsocios|nroasociado|000000||"
         Top             =   255
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
         Index           =   0
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "C�digo Socio|N|N|1|999999|rsocios|codsocio|000000|S|"
         Top             =   255
         Width           =   885
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
         Index           =   1
         Left            =   3510
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||rsocios|nomsocio|||"
         Top             =   255
         Width           =   5475
      End
      Begin VB.Label Label36 
         Caption         =   "Nro.Asociado"
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
         Left            =   9840
         TabIndex        =   165
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre "
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
         Left            =   2685
         TabIndex        =   68
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo"
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
         Left            =   360
         TabIndex        =   67
         Top             =   315
         Width           =   705
      End
   End
   Begin VB.TextBox text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   9630
      TabIndex        =   85
      Top             =   1020
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   63
      Top             =   7680
      Width           =   2865
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
         Left            =   120
         TabIndex        =   64
         Top             =   180
         Width           =   2655
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
      Left            =   12660
      TabIndex        =   38
      Top             =   7800
      Width           =   1035
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
      Left            =   11460
      TabIndex        =   36
      Top             =   7800
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6900
      Top             =   6570
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
      Left            =   12660
      TabIndex        =   70
      Top             =   7800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6060
      Left            =   240
      TabIndex        =   65
      Top             =   1590
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   10689
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmManSocios.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(1)=   "Label29"
      Tab(0).Control(2)=   "imgZoom(0)"
      Tab(0).Control(3)=   "imgMail(0)"
      Tab(0).Control(4)=   "Label19"
      Tab(0).Control(5)=   "imgFec(2)"
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(7)=   "Text1(2)"
      Tab(0).Control(8)=   "Text1(20)"
      Tab(0).Control(9)=   "FrameDatosDtoAdministracion"
      Tab(0).Control(10)=   "Text1(12)"
      Tab(0).Control(11)=   "Frame4"
      Tab(0).Control(12)=   "Text1(7)"
      Tab(0).Control(13)=   "Frame6"
      Tab(0).Control(14)=   "Frame7"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Secciones"
      TabPicture(1)   =   "frmManSocios.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameAux1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tel�fonos"
      TabPicture(2)   =   "frmManSocios.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux0"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Documentos"
      TabPicture(3)   =   "frmManSocios.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label16"
      Tab(3).Control(1)=   "Label17"
      Tab(3).Control(2)=   "imgFec(3)"
      Tab(3).Control(3)=   "Toolbar4"
      Tab(3).Control(4)=   "Frame5"
      Tab(3).Control(5)=   "Toolbar3"
      Tab(3).Control(6)=   "lw1"
      Tab(3).Control(7)=   "Toolbar2"
      Tab(3).Control(8)=   "Text3(0)"
      Tab(3).Control(9)=   "Frame8"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Pozos"
      TabPicture(4)   =   "frmManSocios.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameAux2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Miembros"
      TabPicture(5)   =   "frmManSocios.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FrameAux3"
      Tab(5).ControlCount=   1
      Begin VB.Frame FrameAux3 
         BorderStyle     =   0  'None
         Height          =   5430
         Left            =   -74955
         TabIndex        =   153
         Top             =   405
         Width           =   12360
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   1
            Left            =   720
            MaxLength       =   6
            TabIndex        =   161
            Tag             =   "Lineal|N|N|||rsocios_miembros|numlinea|#####0|S|"
            Text            =   "Linea"
            Top             =   3450
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   5
            Left            =   7020
            MaxLength       =   30
            TabIndex        =   160
            Tag             =   "Capital|N|S|||rsocios_miembros|capital|###,###,##0.00||"
            Text            =   "Capital"
            Top             =   3450
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   2
            Left            =   1440
            MaxLength       =   9
            TabIndex        =   155
            Tag             =   "NIF miembro|T|N|||rsocios_miembros|nifmiembro|||"
            Text            =   "Nif"
            Top             =   3450
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   0
            Left            =   150
            MaxLength       =   6
            TabIndex        =   154
            Tag             =   "C�digo Socio|N|N|1|999999|rsocios_miembros|codsocio|000000|S|"
            Text            =   "Socio"
            Top             =   3480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   4
            Left            =   5250
            MaxLength       =   30
            TabIndex        =   159
            Tag             =   "Votos|N|S|||rsocios_miembros|votos|###,###,##0.00||"
            Text            =   "Votos"
            Top             =   3450
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   3
            Left            =   1980
            MaxLength       =   40
            TabIndex        =   157
            Tag             =   "Nombre|T|N|||rsocios_miembros|nommiembro|||"
            Text            =   "Nombre"
            Top             =   3450
            Visible         =   0   'False
            Width           =   3165
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   3
            Left            =   45
            TabIndex        =   156
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
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
            Index           =   3
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
            Bindings        =   "frmManSocios.frx":00B4
            Height          =   4545
            Index           =   3
            Left            =   60
            TabIndex        =   158
            Top             =   510
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   8017
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   19
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
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   4395
         Left            =   -65100
         TabIndex        =   148
         Top             =   420
         Visible         =   0   'False
         Width           =   3465
         Begin VB.CommandButton cmdAccCRM 
            Height          =   375
            Index           =   0
            Left            =   0
            Picture         =   "frmManSocios.frx":00CC
            Style           =   1  'Graphical
            TabIndex        =   151
            ToolTipText     =   "Insertar Im�gen"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdAccCRM 
            Height          =   375
            Index           =   1
            Left            =   1080
            Picture         =   "frmManSocios.frx":0ACE
            Style           =   1  'Graphical
            TabIndex        =   150
            ToolTipText     =   "Ver Documento"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdAccCRM 
            Height          =   375
            Index           =   2
            Left            =   480
            Picture         =   "frmManSocios.frx":1058
            Style           =   1  'Graphical
            TabIndex        =   149
            ToolTipText     =   "Eliminar"
            Top             =   0
            Width           =   375
         End
         Begin VB.Image Image1 
            Height          =   3780
            Left            =   0
            Stretch         =   -1  'True
            Top             =   420
            Width           =   3405
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Domicilio Fiscal"
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
         Height          =   1485
         Left            =   -74850
         TabIndex        =   142
         Top             =   915
         Width           =   5985
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
            Index           =   3
            Left            =   1275
            MaxLength       =   35
            TabIndex        =   5
            Tag             =   "Domicilio|T|N|||rsocios|dirsocio|||"
            Top             =   240
            Width           =   4635
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
            Left            =   1275
            MaxLength       =   6
            TabIndex        =   6
            Tag             =   "C.Postal|T|N|||rsocios|codpostal|||"
            Top             =   660
            Width           =   735
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
            Left            =   2100
            MaxLength       =   30
            TabIndex        =   7
            Tag             =   "Poblaci�n|T|N|||rsocios|pobsocio|||"
            Top             =   660
            Width           =   3795
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
            Index           =   6
            Left            =   1275
            MaxLength       =   30
            TabIndex        =   8
            Tag             =   "Provincia|T|N|||rsocios|prosocio|||"
            Top             =   1065
            Width           =   4620
         End
         Begin VB.Label Label6 
            Caption         =   "Direcci�n"
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
            Left            =   240
            TabIndex        =   145
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label28 
            Caption         =   "Provincia"
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
            Left            =   240
            TabIndex        =   144
            Top             =   1125
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
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
            Left            =   240
            TabIndex        =   143
            Top             =   690
            Width           =   945
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Domicilio de Correo"
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
         Height          =   1485
         Left            =   -74850
         TabIndex        =   138
         Top             =   2460
         Width           =   5985
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
            Index           =   25
            Left            =   1275
            MaxLength       =   35
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||rsocios|dirsociocorreo|||"
            Top             =   270
            Width           =   4605
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
            Left            =   1275
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "C.Postal|T|N|||rsocios|codpostalcorreo|||"
            Top             =   660
            Width           =   735
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
            Index           =   23
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Poblaci�n|T|N|||rsocios|pobsociocorreo|||"
            Top             =   660
            Width           =   3825
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
            Index           =   22
            Left            =   1275
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||rsocios|prosociocorreo|||"
            Top             =   1065
            Width           =   4590
         End
         Begin VB.Label Label6 
            Caption         =   "Direcci�n"
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
            Left            =   240
            TabIndex        =   141
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label26 
            Caption         =   "Provincia"
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
            Left            =   240
            TabIndex        =   140
            Top             =   1095
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
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
            Left            =   240
            TabIndex        =   139
            Top             =   690
            Width           =   975
         End
      End
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         Height          =   5160
         Left            =   -74955
         TabIndex        =   131
         Top             =   405
         Width           =   12360
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   2
            Left            =   1290
            MaxLength       =   9
            TabIndex        =   133
            Tag             =   "Acciones|N|N|||rsocios_pozos|acciones|##0.00||"
            Text            =   "Acciones"
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux3 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   3
            Left            =   1950
            MaxLength       =   30
            TabIndex        =   134
            Tag             =   "Observaciones|T|S|||rsocios_pozos|observac|||"
            Text            =   "observaciones"
            Top             =   3420
            Visible         =   0   'False
            Width           =   4725
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   0
            Left            =   300
            MaxLength       =   6
            TabIndex        =   135
            Tag             =   "C�digo Socio|N|N|1|999999|rsocios_pozos|codsocio|000000|S|"
            Text            =   "Socio"
            Top             =   3420
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   1
            Left            =   750
            MaxLength       =   9
            TabIndex        =   132
            Tag             =   "Numero Fases|N|N|||rsocios_pozos|numfases|000|S|"
            Text            =   "Fases"
            Top             =   3405
            Visible         =   0   'False
            Width           =   465
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   2
            Left            =   45
            TabIndex        =   136
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
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
            Index           =   2
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
            Bindings        =   "frmManSocios.frx":1A5A
            Height          =   4545
            Index           =   2
            Left            =   60
            TabIndex        =   137
            Top             =   510
            Width           =   7995
            _ExtentX        =   14102
            _ExtentY        =   8017
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   19
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
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         BeginProperty Font 
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
         Left            =   -63660
         TabIndex        =   123
         Text            =   "Text4"
         Top             =   1050
         Width           =   1455
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
         Left            =   -70185
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha Nacimiento|F|S|||rsocios|fechanac|dd/mm/yyyy||"
         Top             =   525
         Width           =   1260
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   5370
         Left            =   45
         TabIndex        =   98
         Top             =   420
         Width           =   13165
         Begin VB.Frame Frame3 
            Height          =   735
            Left            =   30
            TabIndex        =   111
            Top             =   4530
            Width           =   12480
            Begin VB.TextBox txtAux2 
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
               Index           =   0
               Left            =   9885
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   114
               Text            =   "nomiva"
               Top             =   240
               Width           =   2475
            End
            Begin VB.TextBox txtAux2 
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
               Index           =   5
               Left            =   5655
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   113
               Text            =   "nomCuenta Proveedor"
               Top             =   240
               Width           =   3795
            End
            Begin VB.TextBox txtAux2 
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
               Index           =   4
               Left            =   1245
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   112
               Text            =   "nomCuenta Cliente"
               Top             =   240
               Width           =   3435
            End
            Begin VB.Label Label13 
               Caption         =   "Cta.Cliente"
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
               Left            =   90
               TabIndex        =   117
               Top             =   270
               Width           =   1095
            End
            Begin VB.Label Label12 
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
               Height          =   315
               Left            =   9540
               TabIndex        =   116
               Top             =   270
               Width           =   375
            End
            Begin VB.Label Label9 
               Caption         =   "Cta.Prov."
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
               Left            =   4725
               TabIndex        =   115
               Top             =   270
               Width           =   900
            End
         End
         Begin VB.CommandButton cmdAux 
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
            Height          =   330
            Index           =   5
            Left            =   9270
            TabIndex        =   110
            ToolTipText     =   "Buscar iva"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   6
            Left            =   8550
            MaxLength       =   2
            TabIndex        =   44
            Tag             =   "Cod.Iva|N|N|||rsocios_seccion|codiva|00||"
            Text            =   "iva"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtAux2 
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
            Height          =   330
            Index           =   1
            Left            =   1845
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   109
            Text            =   "Nombre seccion"
            Top             =   2925
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.CommandButton cmdAux 
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
            Height          =   330
            Index           =   4
            Left            =   1665
            TabIndex        =   108
            ToolTipText     =   "Buscar fecha"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton cmdAux 
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
            Height          =   330
            Index           =   1
            Left            =   6525
            TabIndex        =   106
            ToolTipText     =   "Buscar fecha"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton cmdAux 
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
            Height          =   330
            Index           =   0
            Left            =   5760
            TabIndex        =   105
            ToolTipText     =   "Buscar fecha"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   5
            Left            =   7605
            MaxLength       =   10
            TabIndex        =   43
            Tag             =   "Cta.Contable Proveedor|T|S|||rsocios_seccion|codmacpro|||"
            Text            =   "cta provee"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.CommandButton cmdAux 
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
            Height          =   330
            Index           =   2
            Left            =   7380
            TabIndex        =   103
            ToolTipText     =   "Buscar cuenta"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   4
            Left            =   6660
            MaxLength       =   10
            TabIndex        =   42
            Tag             =   "Cta.Contable Cliente|T|S|||rsocios_seccion|codmaccli|||"
            Text            =   "cta cliente"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.CommandButton cmdAux 
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
            Height          =   330
            Index           =   3
            Left            =   8325
            TabIndex        =   102
            ToolTipText     =   "Buscar cuenta"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   3
            Left            =   5940
            MaxLength       =   10
            TabIndex        =   41
            Tag             =   "Fecha Baja|F|S|||rsocios_seccion|fecbaja|dd/mm/yyyy||"
            Text            =   "fecha baja"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   2
            Left            =   5175
            MaxLength       =   10
            TabIndex        =   40
            Tag             =   "Fecha Alta|F|N|||rsocios_seccion|fecalta|dd/mm/yyyy||"
            Text            =   "fecha alta"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   1
            Left            =   990
            MaxLength       =   3
            TabIndex        =   39
            Tag             =   "Seccion|N|N|||rsocios_seccion|codsecci|000|S|"
            Text            =   "seccion"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   0
            Left            =   225
            MaxLength       =   6
            TabIndex        =   101
            Tag             =   "Nombre|N|N|||rsocios_seccion|codsocio|000000|S|"
            Text            =   "socio"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   45
            TabIndex        =   99
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
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
            Index           =   1
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
            Bindings        =   "frmManSocios.frx":1A72
            Height          =   3820
            Index           =   1
            Left            =   30
            TabIndex        =   100
            Top             =   510
            Width           =   12360
            _ExtentX        =   21802
            _ExtentY        =   6747
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   19
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
      Begin VB.Frame Frame4 
         Caption         =   "Tel�fonos"
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
         Height          =   1125
         Left            =   -74850
         TabIndex        =   93
         Top             =   3990
         Width           =   5985
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
            Index           =   11
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "M�vil|T|S|||rsocios|movsocio|||"
            Top             =   630
            Width           =   1545
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
            Index           =   9
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Telfno 2|T|S|||rsocios|telsoci2|||"
            Text            =   "123456789012345"
            Top             =   645
            Width           =   1545
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
            Index           =   10
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Tel�fono 3|T|S|||rsocios|telsoci3|||"
            Text            =   "0000000000"
            Top             =   240
            Width           =   1545
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
            Index           =   8
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Tel�fono 1|T|S|||rsocios|telsoci1|||"
            Text            =   "123456789012345"
            Top             =   225
            Width           =   1545
         End
         Begin VB.Label Label11 
            Caption         =   "N�mero 3"
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
            Left            =   3225
            TabIndex        =   97
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label Label20 
            Caption         =   "N�mero 2"
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
            Left            =   240
            TabIndex        =   96
            Top             =   690
            Width           =   990
         End
         Begin VB.Label Label25 
            Caption         =   "M�vil"
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
            Left            =   3210
            TabIndex        =   95
            Top             =   690
            Width           =   495
         End
         Begin VB.Label Label41 
            Caption         =   "N�mero 1"
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
            Left            =   225
            TabIndex        =   94
            Top             =   270
            Width           =   1005
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
         Index           =   12
         Left            =   -73575
         MaxLength       =   50
         TabIndex        =   17
         Tag             =   "E-mail|T|S|||rsocios|maisocio|||"
         Top             =   5385
         Width           =   4545
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   5490
         Left            =   -74955
         TabIndex        =   77
         Top             =   405
         Width           =   12360
         Begin VB.TextBox txtAux 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Index           =   16
            Left            =   7425
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   62
            Tag             =   "Observaciones|T|S|||rsocios_telefonos|observaciones|||"
            Top             =   4335
            Width           =   4755
         End
         Begin VB.TextBox txtAux 
            BeginProperty Font 
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
            Left            =   8595
            MaxLength       =   35
            TabIndex        =   51
            Tag             =   "Direccion|T|S|||rsocios_telefonos|direccion|||"
            Top             =   825
            Width           =   3540
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Inactivo"
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
            Left            =   11160
            TabIndex        =   55
            Tag             =   "Inactivo|N|N|||rsocios_telefonos|inactivo||N|"
            Top             =   2205
            Width           =   1155
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   1
            Left            =   360
            MaxLength       =   9
            TabIndex        =   46
            Tag             =   "C�digo telefono|T|N|||rsocios_telefono|idtelefono||S|"
            Text            =   "idtelefon"
            Top             =   3405
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   0
            Left            =   -90
            MaxLength       =   6
            TabIndex        =   45
            Tag             =   "C�digo Socio|N|N|1|999999|rsocios_telefonos|codsocio|000000|S|"
            Text            =   "cods"
            Top             =   3420
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   3
            Left            =   1575
            MaxLength       =   25
            TabIndex        =   48
            Tag             =   "IMEI|T|N|||rsocios_telefonos|imei|||"
            Text            =   "imei"
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   4
            Left            =   2925
            MaxLength       =   6
            TabIndex        =   49
            Tag             =   "Cod.Pobla|T|S|||rsocios_telefonos|codpostal|||"
            Text            =   "C.P."
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            BeginProperty Font 
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
            Left            =   8595
            MaxLength       =   30
            TabIndex        =   52
            Tag             =   "Poblacion|T|S|||rsocios_telefonos|poblacion|||"
            Top             =   1275
            Width           =   3525
         End
         Begin VB.TextBox txtAux 
            BeginProperty Font 
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
            Left            =   8595
            MaxLength       =   30
            TabIndex        =   53
            Tag             =   "Provincia|T|S|||rsocios_telefonos|provincia|||"
            Text            =   "Prov"
            Top             =   1740
            Width           =   3525
         End
         Begin VB.TextBox txtAux 
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
            Left            =   9315
            MaxLength       =   4
            TabIndex        =   59
            Tag             =   "Sucursal|N|S|0|9999|rsocios_telefonos|codsucur|0000||"
            Top             =   3630
            Width           =   630
         End
         Begin VB.TextBox txtAux 
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
            Left            =   10035
            MaxLength       =   2
            TabIndex        =   60
            Tag             =   "Digito Control|T|S|||rsocios_telefonos|digcontr|00||"
            Top             =   3630
            Width           =   495
         End
         Begin VB.TextBox txtAux 
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
            Index           =   15
            Left            =   10575
            MaxLength       =   10
            TabIndex        =   61
            Tag             =   "Cuenta Bancaria|T|S|||rsocios_telefonos|cuentaba|0000000000||"
            Top             =   3630
            Width           =   1575
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
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
            Height          =   290
            Index           =   2
            Left            =   900
            MaxLength       =   9
            TabIndex        =   47
            Tag             =   "NIF|T|S|||rsocios_telefonos|nif|||"
            Text            =   "nif"
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
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
            Left            =   8595
            MaxLength       =   4
            TabIndex        =   58
            Tag             =   "Banco|N|S|0|9999|rsocios_telefonos|codbanco|0000||"
            Top             =   3630
            Width           =   600
         End
         Begin VB.TextBox txtAux 
            BeginProperty Font 
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
            Left            =   8595
            MaxLength       =   40
            TabIndex        =   50
            Tag             =   "Nombre|T|S|||rsocios_telefonos|nombre|||"
            Top             =   375
            Width           =   3540
         End
         Begin VB.TextBox txtAux 
            BeginProperty Font 
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
            Left            =   8595
            MaxLength       =   50
            TabIndex        =   57
            Tag             =   "Mail|T|S|||rsocios_telefonos|mail|||"
            Top             =   3165
            Width           =   3540
         End
         Begin VB.TextBox txtAux 
            BeginProperty Font 
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
            Left            =   8595
            MaxLength       =   25
            TabIndex        =   56
            Tag             =   "SIM|T|N|||rsocios_telefonos|sim|||"
            Text            =   "1234567890123456789012345"
            Top             =   2715
            Width           =   3510
         End
         Begin VB.TextBox txtAux 
            BeginProperty Font 
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
            Left            =   8595
            MaxLength       =   10
            TabIndex        =   54
            Tag             =   "Tel�fono|T|S|||rsocios_telefonos|telefono1|||"
            Text            =   "1234567890"
            Top             =   2235
            Width           =   1410
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   45
            TabIndex        =   78
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
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
            Bindings        =   "frmManSocios.frx":1A8A
            Height          =   4545
            Index           =   0
            Left            =   45
            TabIndex        =   79
            Top             =   495
            Width           =   7350
            _ExtentX        =   12965
            _ExtentY        =   8017
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   19
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
               AllowFocus      =   0   'False
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label8 
            Caption         =   "Direcci�n"
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
            Left            =   7425
            TabIndex        =   107
            Top             =   870
            Width           =   1140
         End
         Begin VB.Image imgZoom 
            Height          =   240
            Index           =   1
            Left            =   9000
            Tag             =   "-1"
            ToolTipText     =   "Zoom descripci�n"
            Top             =   4050
            Width           =   240
         End
         Begin VB.Label Label39 
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
            Left            =   7425
            TabIndex        =   89
            Top             =   4110
            Width           =   1440
         End
         Begin VB.Label Label24 
            Caption         =   "Provincia"
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
            Left            =   7425
            TabIndex        =   88
            Top             =   1815
            Width           =   1230
         End
         Begin VB.Label Label2 
            Caption         =   "Poblaci�n"
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
            Left            =   7425
            TabIndex        =   87
            Top             =   1365
            Width           =   1230
         End
         Begin VB.Image imgMail 
            Height          =   240
            Index           =   1
            Left            =   8280
            Top             =   3210
            Width           =   240
         End
         Begin VB.Label Label38 
            Caption         =   "Banco"
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
            Left            =   7425
            TabIndex        =   84
            Top             =   3675
            Width           =   825
         End
         Begin VB.Label Label33 
            Caption         =   "Nombre"
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
            Left            =   7425
            TabIndex        =   83
            Top             =   420
            Width           =   1140
         End
         Begin VB.Label Label32 
            Caption         =   "Mail"
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
            Left            =   7425
            TabIndex        =   82
            Top             =   3210
            Width           =   780
         End
         Begin VB.Label Label30 
            Caption         =   "SIM"
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
            Left            =   7425
            TabIndex        =   81
            Top             =   2730
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "Tel�fono"
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
            Left            =   7425
            TabIndex        =   80
            Top             =   2280
            Width           =   975
         End
      End
      Begin VB.Frame FrameDatosDtoAdministracion 
         Caption         =   "Datos Relacionados Dto.Administraci�n"
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
         Height          =   4455
         Left            =   -68790
         TabIndex        =   72
         Top             =   420
         Width           =   7110
         Begin VB.CheckBox chkAbonos 
            Caption         =   "NO devolver aportaciones"
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
            Left            =   3030
            TabIndex        =   166
            Tag             =   "No Devolver Aportaciones|N|N|||rsocios|nodevolverapor||N|"
            Top             =   690
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
            Index           =   29
            Left            =   5370
            MaxLength       =   15
            TabIndex        =   29
            Tag             =   "Capital|N|S|||rsocios|capital|###,###,##0.00||"
            Top             =   3060
            Width           =   1605
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
            Left            =   1770
            MaxLength       =   15
            TabIndex        =   28
            Tag             =   "Votos|N|S|||rsocios|votos|###,###,##0.00||"
            Top             =   3060
            Width           =   1425
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
            Index           =   27
            Left            =   1770
            MaxLength       =   4
            TabIndex        =   30
            Tag             =   "IBAN|T|S|||rsocios|iban|||"
            Text            =   "Text1"
            Top             =   3495
            Width           =   645
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Emite Factura"
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
            Left            =   510
            TabIndex        =   25
            Tag             =   "Emite Factura|N|N|||rsocios|emitefact||N|"
            Top             =   2640
            Width           =   2025
         End
         Begin VB.TextBox text2 
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
            Index           =   26
            Left            =   2415
            TabIndex        =   147
            Top             =   3900
            Width           =   4560
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
            Left            =   1755
            MaxLength       =   4
            TabIndex        =   35
            Tag             =   "C�digo Globalgap|T|S|||rsocios|codigoggap|||"
            Top             =   3900
            Width           =   615
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Factura Interna ADV"
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
            Left            =   2550
            TabIndex        =   26
            Tag             =   "Fact.Interna ADV|N|N|0|1|rsocios|esfactadvinterna||N|"
            Top             =   2640
            Width           =   2625
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
            Left            =   4590
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Tag             =   "Tipo Relacion|N|N|0|2|rsocios|tiporelacion||N|"
            Top             =   2160
            Width           =   2400
         End
         Begin VB.TextBox text2 
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
            Index           =   21
            Left            =   2325
            TabIndex        =   118
            Top             =   1110
            Width           =   4680
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
            Left            =   1725
            MaxLength       =   2
            TabIndex        =   20
            Tag             =   "C�digo Cooperativa|N|N|0|99|rsocios|codcoope|00||"
            Top             =   1110
            Width           =   555
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Correo"
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
            Left            =   5370
            TabIndex        =   27
            Tag             =   "Correo|N|N|||rsocios|correo||N|"
            Top             =   2640
            Width           =   1485
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
            Left            =   2550
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Tag             =   "Tipo Produccion|N|N|0|3|rsocios|tipoprod||N|"
            Top             =   2160
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
            Index           =   0
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Tag             =   "Tipo IRPF|N|N|0|2|rsocios|tipoirpf||N|"
            Top             =   2160
            Width           =   1830
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
            Index           =   19
            Left            =   5370
            MaxLength       =   10
            TabIndex        =   34
            Tag             =   "Cuenta Bancaria|T|S|||rsocios|cuentaba|0000000000||"
            Text            =   "Text1"
            Top             =   3495
            Width           =   1605
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
            Left            =   4560
            MaxLength       =   2
            TabIndex        =   33
            Tag             =   "Digito Control|T|S|||rsocios|digcontr|00||"
            Text            =   "Text1"
            Top             =   3495
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
            Index           =   17
            Left            =   3615
            MaxLength       =   4
            TabIndex        =   32
            Tag             =   "Sucursal|N|S|0|9999|rsocios|codsucur|0000||"
            Text            =   "Text1"
            Top             =   3495
            Width           =   705
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
            Left            =   2670
            MaxLength       =   4
            TabIndex        =   31
            Tag             =   "Banco|N|S|0|9999|rsocios|codbanco|0000||"
            Text            =   "Text1"
            Top             =   3495
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
            Index           =   15
            Left            =   1725
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "C�digo Situacion|N|N|0|99|rsocios|codsitua|00||"
            Top             =   1500
            Width           =   555
         End
         Begin VB.TextBox text2 
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
            Index           =   15
            Left            =   2325
            TabIndex        =   73
            Top             =   1500
            Width           =   4680
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
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   19
            Tag             =   "Fecha Baja|F|S|||rsocios|fechabaja|dd/mm/yyyy||"
            Top             =   720
            Width           =   1260
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
            Index           =   13
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "Fecha Alta|F|N|||rsocios|fechaalta|dd/mm/yyyy||"
            Top             =   330
            Width           =   1260
         End
         Begin VB.Image imgDoc 
            Height          =   435
            Index           =   0
            Left            =   6480
            ToolTipText     =   "Aportaciones del Socio"
            Top             =   660
            Width           =   495
         End
         Begin VB.Image imgDoc 
            Height          =   465
            Index           =   1
            Left            =   6480
            ToolTipText     =   "Impresi�n Documento Alta/Baja"
            Top             =   180
            Width           =   510
         End
         Begin VB.Label Label14 
            Caption         =   "Documentos Alta/Baja/Transmisi�n"
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
            Left            =   3000
            TabIndex        =   164
            Top             =   330
            Width           =   3480
         End
         Begin VB.Label Label35 
            Caption         =   "Capital"
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
            Left            =   4560
            TabIndex        =   163
            Top             =   3090
            Width           =   705
         End
         Begin VB.Label Label34 
            Caption         =   "Votos"
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
            TabIndex        =   162
            Top             =   3120
            Width           =   705
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   180
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   2700
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1440
            ToolTipText     =   "Buscar globalgap"
            Top             =   3930
            Width           =   240
         End
         Begin VB.Label Label27 
            Caption         =   "Globalgap"
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
            TabIndex        =   146
            Top             =   3930
            Width           =   1005
         End
         Begin VB.Label Label21 
            Caption         =   "Relaci�n Cooperativa"
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
            Left            =   4590
            TabIndex        =   128
            Top             =   1920
            Width           =   2310
         End
         Begin VB.Label Label15 
            Caption         =   "Cooperativa"
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
            TabIndex        =   119
            Top             =   1170
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1410
            ToolTipText     =   "Buscar Cooperativa"
            Top             =   1140
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1410
            Picture         =   "frmManSocios.frx":1AA2
            ToolTipText     =   "Buscar fecha"
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1410
            Picture         =   "frmManSocios.frx":1B2D
            ToolTipText     =   "Buscar fecha"
            Top             =   330
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo Productor"
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
            Left            =   2520
            TabIndex        =   91
            Top             =   1920
            Width           =   1830
         End
         Begin VB.Label Label43 
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
            Left            =   180
            TabIndex        =   90
            Top             =   1920
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN Socio"
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
            Left            =   180
            TabIndex        =   86
            Top             =   3540
            Width           =   1005
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1410
            ToolTipText     =   "Buscar Situaci�n"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Label Label23 
            Caption         =   "Situaci�n"
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
            TabIndex        =   76
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label Label22 
            Caption         =   "Fecha Baja"
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
            Left            =   210
            TabIndex        =   75
            Top             =   750
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Fecha Alta"
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
            Left            =   210
            TabIndex        =   74
            Top             =   360
            Width           =   1215
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
         Index           =   20
         Left            =   -68760
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Tag             =   "Observaciones|T|S|||rsocios|observaciones|||"
         Top             =   5190
         Width           =   6915
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
         Left            =   -73590
         MaxLength       =   9
         TabIndex        =   3
         Tag             =   "NIF / CIF|T|N|||rsocios|nifsocio|||"
         Top             =   520
         Width           =   1290
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   3690
         Left            =   -74850
         TabIndex        =   120
         Top             =   450
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   6509
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas"
               Object.Tag             =   "0"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Entradas por Huerto"
               Object.Tag             =   "1"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Detalle de Entradas"
               Object.Tag             =   "2"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas Cliente"
               Object.Tag             =   "4"
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Impresi�n Documentos"
               Object.Tag             =   "3"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imagenes"
               Object.Tag             =   "5"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   3855
         Left            =   -74040
         TabIndex        =   121
         Top             =   420
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   2370
         Left            =   -74850
         TabIndex        =   127
         Top             =   480
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   4180
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas"
               Object.Tag             =   "0"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Entradas por Huerto"
               Object.Tag             =   "1"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Detalle Entradas"
               Object.Tag             =   "2"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas Clientes"
               Object.Tag             =   "3"
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   1005
         Left            =   -64980
         TabIndex        =   130
         Top             =   1440
         Width           =   2760
         Begin VB.CheckBox Check2 
            Caption         =   "Imprimir Facturas"
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
            Left            =   135
            TabIndex        =   125
            Top             =   720
            Width           =   2220
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Detalle Albaranes"
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
            Left            =   135
            TabIndex        =   124
            Top             =   270
            Width           =   2220
         End
      End
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   390
         Left            =   -64860
         TabIndex        =   129
         Top             =   1590
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir documentos"
               Object.Tag             =   "0"
               Style           =   2
               Value           =   1
            EndProperty
         EndProperty
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   -64140
         Picture         =   "frmManSocios.frx":1BB8
         ToolTipText     =   "Buscar fecha"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label17 
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
         Height          =   255
         Left            =   -64860
         TabIndex        =   126
         Top             =   1050
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   300
         Left            =   -64920
         TabIndex        =   122
         Top             =   510
         Width           =   2865
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Nacimiento"
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
         Left            =   -72225
         TabIndex        =   104
         Top             =   525
         Width           =   1770
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   -70440
         Picture         =   "frmManSocios.frx":1C43
         ToolTipText     =   "Buscar fecha"
         Top             =   525
         Width           =   240
      End
      Begin VB.Label Label19 
         Caption         =   "E-mail"
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
         Left            =   -74580
         TabIndex        =   92
         Top             =   5415
         Width           =   705
      End
      Begin VB.Image imgMail 
         Height          =   240
         Index           =   0
         Left            =   -73845
         Top             =   5415
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   -67140
         Tag             =   "-1"
         ToolTipText     =   "Zoom descripci�n"
         Top             =   4920
         Width           =   240
      End
      Begin VB.Label Label29 
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
         Left            =   -68760
         TabIndex        =   71
         Top             =   4890
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "NIF"
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
         Left            =   -74610
         TabIndex        =   69
         Top             =   525
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   6960
      Top             =   6330
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
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
      Left            =   13260
      TabIndex        =   169
      Top             =   150
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
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   3300
      TabIndex        =   152
      Top             =   7740
      Visible         =   0   'False
      Width           =   2715
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
      Begin VB.Menu mnBajaSocio 
         Caption         =   "Baja &de Socio"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnImprimirFases 
         Caption         =   "Impresion por &Fases"
         Shortcut        =   ^O
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
Attribute VB_Name = "frmManSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: C�SAR                    -+-+
' +-+- Men�: General-Clientes-Clientes -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps num�rics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => m�nim 1; si no PK => m�nim 0; m�xim => 99; format => 00)
' (si es DECIMAL; m�nim => 0; m�xim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

Private Const IdPrograma = 2004

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindr� el n� de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Socio As String


' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmSocPrev As frmManSociosPrev ' Socios vista previa
Attribute frmSocPrev.VB_VarHelpID = -1

Private WithEvents frmCtas As frmCtasConta 'cuenta contable
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de iva de conta
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmSit As frmManSituacion 'situaciones de socios
Attribute frmSit.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSeccion 'secciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmCoop As frmManCoope 'cooperativas
Attribute frmCoop.VB_VarHelpID = -1
Private WithEvents frmDoc As frmDocAltaBaja 'documentos de alta/baja socios/campos
Attribute frmDoc.VB_VarHelpID = -1
Private WithEvents frmFac As frmManFactSocios ' mantenimiento de facturas de socios
Attribute frmFac.VB_VarHelpID = -1
Private WithEvents frmHco As frmManHcoFruta ' mantenimiento de hco de fruta
Attribute frmHco.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCampos ' mantenimiento de campos
Attribute frmCam.VB_VarHelpID = -1
Private WithEvents frmGlo As frmManGlobalGap ' mantenimiento basico para globalgap
Attribute frmGlo.VB_VarHelpID = -1
Private WithEvents frmFacPOZ  As frmPOZRecibos  ' mantenimiento de recibos de pozos
Attribute frmFacPOZ.VB_VarHelpID = -1
' *****************************************************
Private frmDocs   As frmFichaTecIMG  'frmDocImgs  mto de imagenes
Private WithEvents frmMens  As frmMensajes  ' para ver la imagen del documento
Attribute frmMens.VB_VarHelpID = -1


Private WithEvents frmApo  As frmAPOAportacion  ' mantenimiento de aportaciones
Attribute frmApo.VB_VarHelpID = -1


Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Ll�nies

Dim NumTabMto As Integer 'Indica quin n� de Tab est� en modo Mantenimient
Dim TituloLinea As String 'Descripci� de la ll�nia que est� en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de ll�nies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de b�squeda posar el valor de poblaci� seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el n� del Bot� PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de ll�nies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim vSeccion As CSeccion
Dim b As Boolean

Private BuscaChekc As String
Private Const CarpetaIMG = "ImgFicFT"

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean

'Cambio en cuentas de la contabilidad
Dim IbanAnt As String
Dim NombreAnt As String
Dim BancoAnt  As String
Dim SucurAnt As String
Dim DigitoAnt As String
Dim CuentaAnt As String

Dim DirecAnt As String
Dim cPostalAnt As String
Dim PoblaAnt As String
Dim ProviAnt As String
Dim NifAnt As String
Dim EMaiAnt As String

Dim FecBajaAnt As String



Private Sub chkAbonos_GotFocus(Index As Integer)
    PonerFocoChk Me.chkAbonos(Index)
End Sub

Private Sub chkAbonos_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAbonos(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAbonos(" & Index & ")|"
    End If
End Sub

Private Sub chkAbonos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAbonos_LostFocus(Index As Integer)
    If Index = 1 And (Modo = 3 Or Modo = 4) Then
        If chkAbonos(Index).Value = 1 Then Text1(25).Text = ""
    End If
End Sub

Private Sub cmdAccCRM_Click(Index As Integer)

    Select Case Index
        Case 0
            Set frmDocs = New frmFichaTecIMG

            frmDocs.vDatos = Text1(0).Text & "|" & Text1(1).Text & "|"
            frmDocs.Opcion = Index
            frmDocs.Show vbModal

            Set frmDocs = Nothing
            
            CargaDatosLW
            
        Case 1 'Impresi�n del documento
            If Me.lw1.SelectedItem Is Nothing Then Exit Sub
        
            ImprimirImagen
            
        Case 2 'Eliminar
            If Me.lw1.SelectedItem Is Nothing Then Exit Sub
            
            EliminarImagen
    End Select
    
End Sub

Private Sub EliminarImagen()
Dim Sql As String
Dim Mens As String
    
    On Error GoTo eEliminarImagen

    Mens = "Va a proceder a eliminar la im�gen de la lista correspondiente al socio. " & vbCrLf & vbCrLf & "� Desea continuar ?" & vbCrLf & vbCrLf
    
    If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Sql = "delete from rfichdocs where codsocio = " & DBSet(Text1(0).Text, "N") & " and codigo = " & Me.lw1.SelectedItem.SubItems(3)
        conn.Execute Sql
        
        CargaDatosLW
        
    End If
    Exit Sub

eEliminarImagen:
    MuestraError Err.Number, "Eliminar im�gen", Err.Description
End Sub



Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'B�SQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                
                    '------------------------------------------------------------------------------
                    '  LOG de acciones
                    ValorAnterior = ""
                    
                    Set LOG = New cLOG
                    LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-I", "rsocios", ObtenerWhereCab(False)
                    Set LOG = Nothing
                    '-----------------------------------------------------------------------------
                
                    CargarUnSocio CLng(Text1(0).Text), "I"
                    ' *** canviar o llevar el WHERE, repasar codEmpre ****
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    'Data1.RecordSource = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
                    ' ***************************************************************
                    TerminaBloquear
                    PosicionarData
                    PonerCampos
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    TerminaBloquear
                    CargarUnSocio CLng(Text1(0).Text), "U"
                    
                    '------------------------------------------------------------------------------
                    '  LOG de acciones
                    Set LOG = New cLOG
                    LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-U", "rsocios", ObtenerWhereCab(False)
                    Set LOG = Nothing
                    '-----------------------------------------------------------------------------
                    
                    
                    
                    '[Monica]10/07/2013: Si han cambiado nombre o CCC pregunto si quieren cambiar los datos de la cuenta en la seccion de horto
                    ModificarDatosCuentaContable
                    
                    PosicionarData
                    PonerCampos
                End If
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han ll�nies ***
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    InsertarLinea
                Case 2 'modificar ll�nies
                    ModificarLinea
                    PosicionarData
            End Select
        ' **************************
    
    End Select
    Screen.MousePointer = vbDefault
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub cmdAux_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 4 'Secciones
            Set frmSec = New frmManSeccion
            frmSec.DatosADevolverBusqueda = "0|1|"
            frmSec.CodigoActual = txtAux(1).Text
            frmSec.Show vbModal
            Set frmSec = Nothing
            PonerFoco txtAux(1)
        
        Case 0, 1 'fecha de alta y fecha de baja
           If ModoLineas = 0 Then Exit Sub
           Screen.MousePointer = vbHourglass
           
           Dim esq As Long
           Dim dalt As Long
           Dim menu As Long
           Dim obj As Object
        
           Set frmC = New frmCal
            
           esq = cmdAux(Index).Left
           dalt = cmdAux(Index).Top
            
           Set obj = cmdAux(Index).Container
        
           While cmdAux(Index).Parent.Name <> obj.Name
                esq = esq + obj.Left
                dalt = dalt + obj.Top
                Set obj = obj.Container
           Wend
            
           menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar
        
           frmC.Left = esq + cmdAux(Index).Parent.Left + 30
           frmC.Top = dalt + cmdAux(Index).Parent.Top + cmdAux(Index).Height + menu - 40
        
           
           frmC.NovaData = Now
           indice = Index + 2
           Me.cmdAux(0).Tag = Index
           
           PonerFormatoFecha txtaux1(indice)
           If txtaux1(indice).Text <> "" Then frmC.NovaData = CDate(txtaux1(indice).Text)
        
           Screen.MousePointer = vbDefault
           frmC.Show vbModal
           Set frmC = Nothing
           PonerFoco txtaux1(indice)
        
        Case 2, 3 'cuentas contables de cliente y proveedor
            If vSeccion Is Nothing Then Exit Sub
            
            indice = Index + 2
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtaux1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtaux1(indice)
        
        
        Case 5 'codigo de iva
            Set frmTIva = New frmTipIVAConta
            frmTIva.DeConsulta = True
            frmTIva.DatosADevolverBusqueda = "0|1|"
            frmTIva.CodigoActual = txtaux1(6).Text
            frmTIva.Show vbModal
            Set frmTIva = Nothing
            PonerFoco txtaux1(6)

    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

' *** si n'hi han combos a la cap�alera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    
    If Index = 0 And (Modo = 3 Or Modo = 4) Then
        '[Monica]12/03/2015: solo si no es Alzira lo dejamos como estaba, para el caso de Alzira siempre abierto
        If vParamAplic.Cooperativa <> 4 Then
            '[Monica]26/03/2014: a�adimos que puedan ser tb ED (Combo1(0).ListIndex = 1) antes solo entidad
            chkAbonos(3).Enabled = (Combo1(0).ListIndex = 2 Or Combo1(0).ListIndex = 1)
            If chkAbonos(3).Enabled = False Then chkAbonos(3).Value = 0
        Else
            chkAbonos(3).Enabled = True
        End If

'        Text1(28).Enabled = (Combo1(0).ListIndex = 2)
'        Text1(29).Enabled = (Combo1(0).ListIndex = 2)
'        If Not Text1(28).Enabled Then
'            Text1(28).Text = ""
'            Text1(29).Text = ""
'        End If
        
    End If
    
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
        
        If Socio = "" Then ProcesarCarpetaImagenes
        
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
    
    If DatosADevolverBusqueda = "" Then
        Set dbAriagro = Nothing
    End If
    
End Sub

Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 18 'index del bot� "primero"
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
        .Buttons(1).Image = 20  'baja de socio
        .Buttons(2).Image = 25  'imprimir fases
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
    
    ' ******* si n'hi han ll�nies *******
    'ICONETS DE LES BARRES ALS TABS DE LL�NIA
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
    'La nevegacion para entradas, facturas....
    ImagenesNavegacion
   'Ponemos los datos del listview
    imgFec(3).Tag = vParam.FecIniCam
    Check1.Value = 1
    Check2.Value = 1
    CargaColumnas 0


    '[Monica]20/03/2014: si es escalona utxera ponemos fecha de revision
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Label10.Caption = "Fecha Revisi�n"
        Text1(7).Tag = "Fecha Revisi�n|F|S|||rsocios|fechanac|dd/mm/yyyy||"
    End If


    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i
    
    '[Monica]18/01/2016: para poder ver las aportaciones del socio, solo las que no son utxera, escalona o mallaes
    Me.imgDoc(0).Picture = frmPpal.imgListPpal.ListImages(9).Picture
    
    
    'fichero-add
    Me.imgDoc(1).Picture = frmPpal.imgListPpal.ListImages(24).Picture
'    'fichero-delete
'    Me.imgDoc(2).Picture = frmPpal.imgListPpal.ListImages(27).Picture
    
    'carga IMAGES de mail
    For i = 0 To Me.imgMail.Count - 1
        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Next i
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han ll�nies *******
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "rsocios"
    Ordenacion = " ORDER BY codsocio"
    '************************************************
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la cap�alera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codsocio=-1"
    Data1.Refresh
       
    ModoLineas = 0
       
    ' **** si n'hi ha algun frame que no te datagrids ***
'    CargaFrame 3, False
    ' *************************************************
         
    CargaCombo
    
    '[Monica]08/04/2015: s�lo Catadau ve el nro de asociado
    Label36.visible = (vParamAplic.Cooperativa = 0)
    Text1(30).visible = (vParamAplic.Cooperativa = 0)
    Text1(30).Enabled = (vParamAplic.Cooperativa = 0)
    
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'b�squeda
        ' *** posar de groc els camps visibles de la clau primaria de la cap�alera ***
        Text1(0).BackColor = vbYellow 'codclien
        ' ****************************************************************************
    End If

    ' Para el chivato
    If DatosADevolverBusqueda = "" Then
        Set dbAriagro = New BaseDatos
        dbAriagro.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, vConfig.User, vConfig.password
    End If

'    If Dir(App.Path & "\ficadobe.dat") = "" Then
'        Toolbar2.Buttons(11).Enabled = False
'        Toolbar2.Buttons(11).visible = False
'        Exit Sub
'    End If
    
    '[Monica]30/04/2013: solo si venimos de frmContRecFact (facturacion de contratos de montifrut)
    If Socio <> "" Then
        Me.chkVistaPrevia.Value = 0
        CadB = "codsocio= " & Socio
        Text1(0).Text = Socio
        
        HacerBusqueda
        SSTab1.Tab = 3
        Toolbar2_ButtonClick Toolbar2.Buttons(11)
    End If



End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    Timer1.Enabled = False
    Label31.visible = False
    
    
    limpiar Me   'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
    Next i
    For i = 0 To chkAbonos.Count - 1
        Me.chkAbonos(i).Value = 0
    Next i
    ' *** si n'hi han combos a la cap�alera ***
    ' *****************************************

    Me.Label31.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funci� del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    BuscaChekc = ""
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de b�squeda, el bot� "Regresar" quan n'hi han datos
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
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Nom�s es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    CmdCancelar.visible = b
    CmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a m�s neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    For i = 0 To 4
        BloquearChk Me.chkAbonos(i), (Modo = 0 Or Modo = 2 Or Modo = 5)
    Next i
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la cap�alera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la cap�alera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    For i = 0 To imgFec.Count - 2
        BloquearImgFec Me, i, Modo
    Next i
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    
    
    'El campo 3(0) NUNCA se puede escribir en el
    Text3(0).Enabled = True
    Text3(0).Text = Me.imgFec(3).Tag
    
    
    ' solo si tenemos registro cargado podemos imprimir documentos
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    Me.imgDoc(1).visible = b
    Me.imgDoc(1).Enabled = b
    Me.Label14.visible = b
    Me.Refresh
    
'    Me.imgDoc(2).visible = b
'    Me.imgDoc(2).Enabled = b
        
    ' ********************************************************
    ' *** si n'hi han ll�nies i imagens de buscar que no estiguen als grids ******
    'Ll�nies Departaments
    b = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
'    BloquearImage imgBuscar(3), Not b
'    BloquearImage imgBuscar(4), Not b
'    BloquearImage imgBuscar(7), Not b
'    imgBuscar(3).Enabled = b
'    imgBuscar(3).visible = b
    ' ****************************************************************************
            
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
        CargaGrid 2, False
        CargaGrid 3, False
    End If
    
    Me.imgDoc(0).Enabled = (Modo = 2 And vParamAplic.Cooperativa = 0)
    Me.imgDoc(0).visible = (Modo = 2 And vParamAplic.Cooperativa = 0)
    
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = b
    DataGridAux(2).Enabled = b
    DataGridAux(3).Enabled = b
'    ' ****** si n'hi han combos a la cap�alera ***********************
'    If (Modo = 0) Or (Modo = 2) Or (Modo = 4) Or (Modo = 5) Then
'        Combo1(0).Enabled = False
'        Combo1(0).BackColor = &H80000018 'groc
'    ElseIf (Modo = 1) Or (Modo = 3) Then
'        Combo1(0).Enabled = True
'        Combo1(0).BackColor = &H80000005 'blanc
'    End If
'    ' ****************************************************************
    
    ' *** si n'hi han ll�nies i alg�n tab que no te datagrid ***
'    BloquearFrameAux Me, "FrameAux3", Modo, NumTabMto
'    BloquearFrameAux2 Me, "FrameAux3", (Modo <> 5) Or (Modo = 5 And indFrame <> 3) 'frame datos viaje indiv.
    ' ***************************
        
    'telefonos
    b = (Modo = 5) And (NumTabMto = 0) 'And (ModoLineas <> 3)
    For i = 1 To 4
        BloquearTxt txtAux(i), Not b
    Next i
    For i = 5 To txtAux.Count - 1
        BloquearTxt txtAux(i), Not b
    Next i
    Me.chkAbonos(1).Enabled = b
    b = (Modo = 5) And (NumTabMto = 0) And ModoLineas = 2
    BloquearTxt txtAux(1), b
    
    'secciones
    b = (Modo = 5) And (NumTabMto = 1) 'And (ModoLineas <> 3)
    For i = 1 To txtaux1.Count - 1
        BloquearTxt txtaux1(i), Not b
    Next i
    b = (Modo = 5) And (NumTabMto = 1) And ModoLineas = 2
    BloquearTxt txtaux1(1), b
    BloquearBtn cmdAux(4), b
    
    'pozos
    b = (Modo = 5) And NumTabMto = 2
    For i = 1 To 3
        BloquearTxt txtaux3(i), Not b
    Next i
    b = (Modo = 5) And NumTabMto = 2 And ModoLineas = 2
    BloquearTxt txtaux3(1), b
    
    'miembros
    b = (Modo = 5) And NumTabMto = 3
    For i = 2 To 5
        BloquearTxt txtaux4(i), Not b
    Next i
    
'    ' capital y votos enabled solo si el socio es entidad
'    Text1(28).Enabled = ((Modo = 1) Or ((Combo1(0).ListIndex = 2) And (Modo = 3 Or Modo = 4)))
'    Text1(29).Enabled = ((Modo = 1) Or ((Combo1(0).ListIndex = 2) And (Modo = 3 Or Modo = 4)))
     
     '-----------------------------
    PonerModoOpcionesMenu (Modo) 'Activar opcions men� seg�n modo
    PonerOpcionesMenu   'Activar opcions de men� seg�n nivell
                        'de permisos de l'usuari

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Men� i Toolbar seg�n el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAP�ALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = b And DatosADevolverBusqueda = ""
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And DatosADevolverBusqueda = "" 'And Not DeConsulta and DatosADevolverBusqueda = ""
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'dar de baja un socio
    Toolbar5.Buttons(1).Enabled = b
    Me.mnBajaSocio.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(8).Enabled = b
'    Toolbar1.Buttons(14).Enabled = b
       
    ' *** si n'hi han ll�nies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And DatosADevolverBusqueda = ""
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    ' ****************************************
    
'    ' *** si n'hi han tabs que no tenen grids ***
'    i = 3
'    If AdoAux(i).Recordset.EOF Then
'        ToolAux(i).Buttons(1).Enabled = b
'        ToolAux(i).Buttons(2).Enabled = False
'        ToolAux(i).Buttons(3).Enabled = False
'    Else
'        ToolAux(i).Buttons(1).Enabled = False
'        ToolAux(i).Buttons(2).Enabled = b
'    End If
    ' *******************************************
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botons de Despla�ament; per a despla�ar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informaci� proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enla�a en el data1
'           -> Si no el carreguem sense enlla�ar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 'telefonos
            tabla = "rsocios_telefonos"
            Sql = "SELECT rsocios_telefonos.codsocio, rsocios_telefonos.idtelefono, rsocios_telefonos.nif, "
            Sql = Sql & " rsocios_telefonos.imei, rsocios_telefonos.codpostal, rsocios_telefonos.nombre, "
            Sql = Sql & " rsocios_telefonos.direccion, rsocios_telefonos.poblacion, rsocios_telefonos.provincia, "
            Sql = Sql & " rsocios_telefonos.telefono1, rsocios_telefonos.sim, rsocios_telefonos.mail, rsocios_telefonos.codbanco, "
            Sql = Sql & " rsocios_telefonos.codsucur, rsocios_telefonos.digcontr, rsocios_telefonos.cuentaba, "
            Sql = Sql & " rsocios_telefonos.observaciones,  rsocios_telefonos.inactivo "
            Sql = Sql & " FROM " & tabla
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE codsocio = -1"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".idtelefono "
            
            
       Case 1 ' secciones
            tabla = "rsocios_seccion"
             Sql = "SELECT rsocios_seccion.codsocio, rsocios_seccion.codsecci, rseccion.nomsecci, rsocios_seccion.fecalta, "
             Sql = Sql & " rsocios_seccion.fecbaja, rsocios_seccion.codmaccli, rsocios_seccion.codmacpro, rsocios_seccion.codiva "
            Sql = Sql & " FROM " & tabla & " INNER JOIN rseccion ON rsocios_seccion.codsecci = rseccion.codsecci "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE codsocio = -1"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".codsecci "
            
       Case 2 ' pozos
            tabla = "rsocios_pozos"
            Sql = "SELECT rsocios_pozos.codsocio, rsocios_pozos.numfases, rsocios_pozos.acciones, rsocios_pozos.observac "
            Sql = Sql & " FROM " & tabla
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE codsocio = -1"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".numfases "
            
       Case 3 ' miembros de entidades
            tabla = "rsocios_miembros"
            Sql = "SELECT rsocios_miembros.codsocio, rsocios_miembros.numlinea, rsocios_miembros.nifmiembro, rsocios_miembros.nommiembro, rsocios_miembros.votos, rsocios_miembros.capital "
            Sql = Sql & " FROM " & tabla
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE codsocio = -1"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".numlinea "
            
            
            
    End Select
    ' ********************************************************************************
    
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
        '   Com la clau principal es �nica, en posar el sql apuntant
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
    indice = CByte(Me.cmdAux(0).Tag + 2)
    txtaux1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag)
    If indice = 0 Then
        Text3(indice).Text = Format(vFecha, "dd/mm/yyyy")
    Else
        Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
    End If
End Sub

Private Sub frmCoop_DatoSeleccionado(CadenaSeleccion As String)
    Text1(21).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo cooperativa
    FormateaCampo Text1(21)
    Text2(21).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre cooperativa
End Sub

Private Sub frmGlo_DatoSeleccionado(CadenaSeleccion As String)
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo globalgap
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre globalgap
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    txtaux1(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codseccion
    FormateaCampo txtaux1(1)
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nomseccion
    
'    Set vSeccion = New CSeccion
'    If vSeccion.LeerDatos(txtaux1(1).Text) Then
'        b = vSeccion.AbrirConta
'    Else
'        Set vSeccion = Nothing
'    End If
End Sub

Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
    Text1(15).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo situacion
    FormateaCampo Text1(15)
    Text2(15).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre situacion
End Sub

Private Sub frmSocPrev_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "codsocio = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "N")
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Tipo de iva
    txtaux1(6).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo txtaux1(6)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomiva
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub Image1_DblClick()
Dim L As Long
Dim c As String
    
    
'    L = Me.lw1.SelectedItem.SubItems(2) '  .Recordset!Codigo
'    If InStr(1, lw1.SelectedItem.SubItems(2), ".pdf") <> 0 Then
    
        ImprimirImagen
        Exit Sub
'    Else
'        C = App.Path & "\" & CarpetaIMG & "\" & L
'    End If
'
'
'
'    Set frmMens = New frmMensajes
'
'    frmMens.OpcionMensaje = 46
'    frmMens.Cadena = C
'    frmMens.Show vbModal
'
'    Set frmMens = Nothing
    
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            '[Monica]12/03/2015: a Alzira le dejamos en cualquier tipo de socio
           If vParamAplic.Cooperativa <> 4 Then
                vCadena = "S�lo est� activo si el socio es una entidad o ED. En este caso, cuando " & vbCrLf & _
                          "est� marcado todas las facturas que se generen para este socio, se" & vbCrLf & _
                          "generar�n como contabilizadas con el n�mero que le corresponda del" & vbCrLf & _
                          "tipo de movimiento." & vbCrLf & vbCrLf & _
                          "Cuando se reciba el documento se podr� cambiar el nro de factura y " & vbCrLf & _
                          "contabilizar en un proceso espec�fico. " & vbCrLf & vbCrLf
            Else
                vCadena = "Cuando est� marcado todas las facturas que se generen para este socio," & vbCrLf & _
                          "se generar�n como contabilizadas con el n�mero que le corresponda " & vbCrLf & _
                          "del tipo de movimiento." & vbCrLf & vbCrLf & _
                          "Cuando se reciba el documento se podr� cambiar el nro de factura y " & vbCrLf & _
                          "contabilizar en un proceso espec�fico. " & vbCrLf & vbCrLf
            End If
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripci�n de Ayuda"
    
End Sub

Private Sub imgDoc_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 ' aportaciones del socio
            Set frmApo = New frmAPOAportacion
            frmApo.CodigoActual = Text1(0).Text
            frmApo.Show vbModal
            Set frmApo = Nothing
        
        Case 1 'documentos de alta baja de socios/campos
            Set frmDoc = New frmDocAltaBaja
            frmDoc.NumCod = Text1(0).Text
            frmDoc.Show vbModal
            Set frmDoc = Nothing
    End Select
    
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
        
       menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar
    
       frmC1.Left = esq + imgFec(Index).Parent.Left + 30
       frmC1.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
       
       frmC1.NovaData = Now
       Select Case Index
            Case 0, 1
                indice = Index + 13
            Case 2
                indice = Index + 5
            Case 3
                indice = 0
       End Select
       
       Me.imgFec(0).Tag = indice
       
       If Index <> 3 Then
           PonerFormatoFecha Text1(indice)
           If Text1(indice).Text <> "" Then frmC1.NovaData = CDate(Text1(indice).Text)
        
           Screen.MousePointer = vbDefault
           frmC1.Show vbModal
           Set frmC1 = Nothing
           PonerFoco Text1(indice)
       Else
           PonerFormatoFecha Text3(indice)
           If Text3(indice).Text <> "" Then frmC1.NovaData = CDate(Text3(indice).Text)
        
           Screen.MousePointer = vbDefault
           frmC1.Show vbModal
           Set frmC1 = Nothing
           PonerFoco Text3(indice)
           
       End If
      'Para la fecha de la navegacion
       If Index = 3 And Text3(0).Text <> "" Then
            imgFec(3).Tag = Text3(0).Text
            CargaDatosLW
       End If
    
End Sub


Private Sub imgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

'    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0: dirMail = Text1(12).Text
        Case 1: dirMail = txtAux(11).Text
    End Select

    If LanzaMailGnral(dirMail) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            indice = 20
            frmZ.pTitulo = "Observaciones del Socio"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(indice)
        Case 1
            indice = 16
            frmZ.pTitulo = "Observaciones del Tel�fono"
            frmZ.pValor = txtAux(indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco txtAux(indice)
    End Select
            
End Sub


Private Sub lw1_Click()

    '[Monica]20/06/2013: a�adida la condicion, pq fallaba
    If CByte(RecuperaValor(lw1.Tag, 1)) = 5 Then
        CargarIMG lw1.SelectedItem.SubItems(2)
    End If
    
End Sub

Private Sub lw1_DblClick()
Dim Seleccionado As Long
Dim cadena As String

    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub

    If Me.DatosADevolverBusqueda <> "" And Socio = "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un socio. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0
        'Facturas
        Set frmFac = New frmManFactSocios
        frmFac.hcoCodMovim = lw1.SelectedItem.SubItems(1)
        frmFac.hcoCodTipoM = lw1.SelectedItem.Text
        frmFac.hcoFechaMov = lw1.SelectedItem.SubItems(2)
        frmFac.Show vbModal
        Set frmFac = Nothing
        
    Case 1
        'Entradas por huerto
        Set frmCam = New frmManCampos
        frmCam.NroCampo = lw1.SelectedItem.Text
        frmCam.Show vbModal
        Set frmCam = Nothing
        
    Case 2
        'Detalle de entradas
        Set frmHco = New frmManHcoFruta
        frmHco.NroAlbaran = lw1.SelectedItem.Text
        frmHco.Show vbModal
        Set frmHco = Nothing
    
    Case 4
        'facturas de clientes
        'Facturas de Pozos
        If lw1.SelectedItem.Text = "RCP" Or lw1.SelectedItem.Text = "RMP" Or lw1.SelectedItem.Text = "TAL" Or lw1.SelectedItem.Text = "RVP" Or lw1.SelectedItem.Text = "RMT" Then
            Set frmFacPOZ = New frmPOZRecibos
            frmFacPOZ.hcoCodMovim = lw1.SelectedItem.SubItems(1)
            frmFacPOZ.hcoCodTipoM = lw1.SelectedItem.Text
            frmFacPOZ.hcoFechaMov = lw1.SelectedItem.SubItems(2)
            frmFacPOZ.Show vbModal
            Set frmFacPOZ = Nothing
        End If
        
    Case 5
        ImprimirImagen
        Screen.MousePointer = vbDefault
        Exit Sub
    
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    lw1.SetFocus
    Seleccionado = lw1.SelectedItem.Index
    CargaDatosLW
    lw1.SelectedItem.Selected = False
    Set lw1.SelectedItem = Nothing
    If lw1.ListItems.Count >= Seleccionado Then
        lw1.ListItems(Seleccionado).Selected = True
        lw1.ListItems(Seleccionado).EnsureVisible
    End If

    Screen.MousePointer = vbDefault
    
End Sub


Private Sub ImprimirImagen()
Dim NFic As Long
Dim vAdobe As String
                
'   If InStr(1, Me.lw1.SelectedItem.SubItems(2), ".pdf") <> "0" Then
'
'        NFic = FreeFile
'        If Dir(App.Path & "\ficadobe.dat") = "" Then
'            MsgBox "Falta fichero de Configuracion. Llame a Ariadna.", vbExclamation
'            Exit Sub
'        End If
'        Open App.Path & "\ficadobe.dat" For Input As #NFic
'        Line Input #NFic, vAdobe
'        Close #NFic
'        Shell vAdobe & " " & Me.lw1.SelectedItem.SubItems(2), vbMaximizedFocus
'
'   Else
'        With frmImprimir
'            .FormulaSeleccion = "{rsocios.codsocio}=" & DBSet(Text1(0).Text, "N") & " and {rfichdocs.codigo} = " & Me.lw1.SelectedItem.SubItems(3)
'            .OtrosParametros = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
'            .Titulo = "Im�gen " & Me.lw1.SelectedItem.SubItems(1)
'            .NumeroParametros = 1
'            .SoloImprimir = False
'            .EnvioEMail = False
'            .NombreRPT = "rImgDocs.rpt"
'
'            .Opcion = 2015
'            .Show vbModal
'        End With
'   End If

   Call ShellExecute(Me.hWnd, "Open", Me.lw1.SelectedItem.SubItems(2), "", "", 1)
   
End Sub



Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnImprimirFases_Click()
    AbrirListado (31)
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnBajaSocio_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonBajaSocio
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

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Select Case Index
        Case 0
            PonerFormatoFecha Text3(Index)
      
            If Text3(0).Text <> "" Then
                imgFec(3).Tag = Text3(0).Text
                CargaDatosLW
            End If
    End Select
End Sub

Private Sub Timer1_Timer()
    Label31.visible = Not Label31.visible
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 5  'B�scar
           mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 11 ' Baja de socios
            mnBajaSocio_Click
        Case 8 'Imprimir
            mnImprimir_Click
        Case 14 'Imprimir
            mnImprimirFases_Click
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la cap�alera no es Text1(0), canviar-ho en <=== *****
    Timer1.Enabled = False
    Label31.visible = False
    
    
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la cap�alera ***
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

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la cap�alera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'    Dim cad As String
'
'    'Cridem al form
'    ' **************** arreglar-ho per a vore lo que es desije ****************
'    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
'    cad = ""
'    cad = cad & ParaGrid(Text1(1), 45, "Nombre")
'    cad = cad & ParaGrid(Text1(0), 10, "C�d.")
'    cad = cad & ParaGrid(Text1(2), 15, "NIF")
'    cad = cad & ParaGrid(Text1(8), 15, "Tel�fono")
'    cad = cad & ParaGrid(Text1(11), 15, "M�vil")
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vtabla = NombreTabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = "1|" '*** els camps que volen que torne ***
'        frmB.vTitulo = "Socios" ' ***** repasa a��: t�tol de BuscaGrid *****
'        frmB.vSelElem = 0
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha posat valors i tenim que es formulari de b�squeda llavors
'        'tindrem que tancar el form llan�ant l'event
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha retornat datos, es a decir NO ha retornat datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If


    
    Set frmSocPrev = New frmManSociosPrev
    frmSocPrev.cWhere = CadB
    frmSocPrev.DatosADevolverBusqueda = "0|1|2|"
    frmSocPrev.Show vbModal
    
    Set frmSocPrev = Nothing

End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
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
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
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
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        ' *** canviar o llevar, si cal, el WHERE; repasar codEmpre ***
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        'CadenaConsulta = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
        ' ******************************************
        PonerCadenaBusqueda
        ' *** si n'hi han ll�nies sense grids ***
'        CargaFrame 0, True
        ' ************************************
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la cap�alera *******
    Text1(0).Text = SugerirCodigoSiguienteStr("rsocios", "codsocio")
    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    ' ***********************************************************
    Combo1_LostFocus (0)
    
    
    ' *** si n'hi han camps de descripci� a la cap�alera ***
    PosarDescripcions
    ' ******************************************************

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
    ' ********************************************
End Sub


Private Sub BotonModificar()

    PonerModo 4

    '[Monica]10/07/2013:me guardo los valores de nombre y CCC por si cambian
    NombreAnt = Text1(1).Text
    IbanAnt = Text1(27).Text
    BancoAnt = Text1(16).Text
    SucurAnt = Text1(17).Text
    DigitoAnt = Text1(18).Text
    CuentaAnt = Text1(19).Text
    
    DirecAnt = Text1(3).Text
    cPostalAnt = Text1(4).Text
    PoblaAnt = Text1(5).Text
    ProviAnt = Text1(6).Text
    NifAnt = Text1(2).Text
    EMaiAnt = Text1(12).Text

    FecBajaAnt = Text1(14).Text

    CargarValoresAnteriores Me, 1

    ' *** bloquejar els camps visibles de la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    Combo1_LostFocus (0)
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
    ' *********************************************************
End Sub

Private Sub BotonBajaSocio()

    '[Monica]19/12/2012: damos aviso si hay entradas esta campa�a
    If HayEntradasSocio(Text1(0).Text) Then
        If MsgBox("Este socio tiene entradas esta campa�a. � Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If

    '[Monica]21/09/2016:
    If Text1(14).Text <> "" Then
        MsgBox "Este socio ya est� dado de baja.", vbExclamation
        Exit Sub
    End If


    Screen.MousePointer = vbHourglass
    Text1(0).Text = Data1.Recordset!Codsocio
    
    frmListado.NumCod = Text1(0).Text
    frmListado.OpcionListado = 23
    frmListado.Show vbModal
    
    TerminaBloquear
    PonerCampos
    Screen.MousePointer = vbDefault
End Sub





Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "�Seguro que desea eliminar el Socio?"
    Cad = Cad & vbCrLf & "C�digo: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset!nomsocio 'Data1.Recordset.Fields(1)
    ' **************************************************************************
    
    'borrem
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ' **** repasar si es diu Data1 l'adodc de la cap�alera ***
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
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
    For i = 0 To DataGridAux.Count - 1
        If i <> 4 Then
            CargaGrid i, True
            If Not AdoAux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
        End If
    Next i
    ' *******************************************

    ' *** si n'hi han ll�nies sense datagrid ***
    ' ***************************************
    
    ' ************* configurar els camps de les descripcions de la cap�alera *************
'    txtAux2(22).Text = PonerNombreDeCod(txtAux(22), "poblacio", "despobla", "codpobla", "N")

'    SSTab1.TabEnabled(5) = (Data1.Recordset!TipoIRPF = 2)
'    SSTab1.TabVisible(5) = (Data1.Recordset!TipoIRPF = 2)
    PosarDescripcions
    
    '[Monica]21/05/2013: indicamos la situacion de bloqueo
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Label31.Caption = Text2(15).Text
        Label31.visible = False
        Label31.visible = (ComprobarCero(Text1(15).Text) >= 1)
        If Label31.visible Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        
        End If
    Else
        Label31.Caption = Text2(15).Text
        Label31.visible = SituacionBloqueo(Text1(15).Text)
    End If
    
    '[Monica]23/10/2013: Solo si es Escalona o Utxera (o de momento montifrut) damos mensaje de que el socio tiene pagos pendientes
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Or vParamAplic.Cooperativa = 12 Then
        '[Monica]15/05/2013: Visualizamos los cobros pendientes del socio
        ComprobarCobrosSocio CStr(Data1.Recordset!Codsocio), ""
    End If

    lblIndicador.Caption = "Datos navegacion"
    Me.Refresh
    DoEvents
    CargaDatosLW

    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu   'Activar opcions de men� seg�n nivell
                        'de permisos de l'usuari
    
End Sub

Private Function SituacionBloqueo(Situ As String) As Boolean
Dim Sql As String

    Sql = "select bloqueo from rsituacion where codsitua = " & DBSet(Situ, "N")
    SituacionBloqueo = (DevuelveValor(Sql) = 1)

End Function


Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V

    Select Case Modo
        Case 1, 3 'B�squeda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la cap�alera ***
                PonerFoco Text1(0)
                ' ***************************************************

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la cap�alera ***
                PonerFoco Text1(0)
                ' *******************************************
        
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    ModoLineas = 0
                    ' *** les ll�nies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Or NumTabMto = 3 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la cap�alera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        ' ********************************************************
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripci� dins del grid, els neteje ***
                        'txtAux2(2).text = ""
                        ' *****************************************************************

                        ' ***  bloquejar i huidar els camps que estan fora del datagrid ***
                        Select Case NumTabMto
                            Case 0 'cuentas bancarias
                                'BotonModificar
'                                BloquearTxt txtaux(11), True
'                                BloquearTxt txtaux(12), True
                            Case 1 'secciones
                                For i = 0 To txtaux1.Count - 1
                                    txtaux1(i).Text = ""
                                    BloquearTxt txtaux1(i), True
                                Next i
                                txtAux2(1).Text = ""
                                txtAux2(4).Text = ""
                                txtAux2(5).Text = ""
                                BloquearTxt txtAux2(1), True
                                BloquearTxt txtAux2(4), True
                                BloquearTxt txtAux2(5), True
                            Case 2 'telefonos
                                For i = 0 To txtAux.Count
                                    BloquearTxt txtAux(i), True
                                Next i
                            Case 3 'miembros
                                For i = 0 To txtaux4.Count
                                    BloquearTxt txtaux4(i), True
                                Next i
                        End Select
                    ' *** els tabs que no tenen datagrid ***
                    ElseIf NumTabMto = 4 Then
                        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        CargaFrame 3, True
                    End If
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ************************
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar ll�nies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ***********************
                    
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de ll�nies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

                    ' ***  bloquejar els camps fora dels grids ***

                    ' ***  bloquejar els camps fora dels grids ***
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select
            
            If NumTabMto = 1 Then
                If Not vSeccion Is Nothing Then
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                End If
            End If
            
            TerminaBloquear

            PosicionarData
            
            ' *** si n'hi han ll�nies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
            ' *********************************************************
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
Dim cta As String
Dim cadMen As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    
'++[Monica] 05/10/2009 comprobamos que la cuenta CCC sea correcta
    If b And (Modo = 3 Or Modo = 4) Then
        If Text1(16).Text = "" Or Text1(17).Text = "" Or Text1(18).Text = "" Or Text1(19).Text = "" Then
            Text1(16).Text = ""
            Text1(17).Text = ""
            Text1(18).Text = ""
            Text1(19).Text = ""
        Else
            cta = Format(Text1(16).Text, "0000") & Format(Text1(17).Text, "0000") & Format(Text1(18).Text, "00") & Format(Text1(19).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El socio no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria del socio no es correcta. � Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(16)
                    b = False
                End If
            Else
'                '[Monica]20/11/2013: a�adimos el tema de la comprobacion del IBAN
'                If Not Comprueba_CC_IBAN(cta, Text1(42).Text) Then
'                    cadMen = "La cuenta IBAN del cliente no es correcta. � Desea continuar ?."
'                    If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                        b = True
'                    Else
'                        PonerFoco Text1(42)
'                        b = False
'                    End If
'                End If

'       sustituido por lo de David
                BuscaChekc = ""
                If Me.Text1(27).Text <> "" Then BuscaChekc = Mid(Text1(27).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.Text1(27).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(27).Text = BuscaChekc & cta
                    Else
                        If Mid(Text1(27).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.Text1(27).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(27)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
                
        End If
        
        If b Then
            If Text1(26).Text <> "" Then
                Sql = "select count(*) from rsocios where codsocio <> " & DBSet(Text1(0).Text, "N") & " and codigoggap = " & DBSet(Text1(26).Text, "T")
                
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Este c�digo de GlobalGap ya est� asignado a otro socio. Revise.", vbExclamation
                    PonerFoco Text1(26)
                End If
            End If
        End If
        
        '[Monica]29/09/2014: comprobamos si vamos a dar de baja que no tenga fecha de alta en programa operativo en ninguno de sus campos
        If b Then
            If Text1(14).Text <> "" Then
                Sql = "select count(*) from rcampos where codsocio = " & DBSet(Text1(0).Text, "N") & " and not fecaltapropera is null "
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Este socio tiene alg�n campo dado de alta en el programa operativo." & vbCrLf & vbCrLf & "No puede darlo de baja. Revise.", vbExclamation
                    PonerFoco Text1(14)
                    b = False
                End If
            End If
        End If
        
    End If
'++

    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
    Cad = "(codsocio=" & Text1(0).Text & ")"
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
    ' ***********************************************************************************
End Sub


Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la cap�alera, repasar codEmpre *******
    vWhere = " WHERE codsocio=" & Data1.Recordset!Codsocio
        ' ***********************************************************************
        
    ' ***** elimina les ll�nies ****
    conn.Execute "DELETE FROM rsocios_seccion " & vWhere
        
    conn.Execute "DELETE FROM rsocios_telefonos " & vWhere
        
    ' *******************************
        
    CargarUnSocio Data1.Recordset!Codsocio, "D"
        
    'Eliminar la CAP�ALERA
    vWhere = " WHERE codsocio=" & Data1.Recordset!Codsocio
    conn.Execute "Delete from " & NombreTabla & vWhere
       
    CadenaCambio = "DELETE FROM rsocios " & vWhere
    
    '------------------------------------------------------------------------------
    '  LOG de acciones
    ValorAnterior = ""
    Set LOG = New cLOG
    LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-D", "rsocios", ObtenerWhereCab(False)
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
       
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

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Sql As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Select Case Index
        Case 0 'cod socio
            PonerFormatoEntero Text1(0)

        Case 1 'NOMBRE
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 2 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
                
                
        Case 15 'situacion
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsituacion", "nomsitua")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Situaci�n: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSit = New frmManSituacion
                        frmSit.DatosADevolverBusqueda = "0|1|"
                        frmSit.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmSit.Show vbModal
                        Set frmSit = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        
        Case 21 'cooperativa
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rcoope", "nomcoope")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Cooperativa: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCoop = New frmManCoope
                        frmCoop.DatosADevolverBusqueda = "0|1|"
                        frmCoop.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCoop.Show vbModal
                        Set frmCoop = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        
'        Case 20 ' Tipo de Iva
'            If PonerFormatoEntero(Text1(Index)) Then
''cuando abra la conexion de ariconta
'
''                text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiposiva", "porceiva", "codigiva", "N", cConta)
''                If text2(Index).Text = "" Then
''                    MsgBox "No existe el Tipo de Iva. Reintroduzca.", vbExclamation
''                    Text1(Index).Text = ""
''                    PonerFoco Text1(Index)
''                End If
'            End If
            
        
        Case 3 ' direccion
            If Modo = 3 Then Text1(25).Text = Text1(Index).Text
            
        Case 4 ' codpostal
            If Modo = 3 Then Text1(24).Text = Text1(Index).Text
            
        Case 5 ' poblacion
            If Modo = 3 Then Text1(23).Text = Text1(Index).Text
            
        Case 6 ' provincia
            If Modo = 3 Then Text1(22).Text = Text1(Index).Text
        
        Case 7 'Fecha no comprobaremos que est� dentro de campa�a
            If Modo = 1 Then Exit Sub
            '[Monica]24/10/2013: no tenia que dar el mensaje de dentro de campa�a
            PonerFormatoFecha Text1(Index) ', True
            
        Case 13, 14 'Fechas
            If Modo = 1 Then Exit Sub
            '[Monica]24/10/2013: aqui si que debe dar el mensaje de dentro de campa�a
            PonerFormatoFecha Text1(Index), True
            
        Case 25 'tipo de movimiento
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
          
        Case 16, 17 'ENTIDAD Y SUCURSAL BANCARIA
            PonerFormatoEntero Text1(Index)
          
        Case 26
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = DevuelveDesdeBDNew(cAgro, "rglobalgap", "descripcion", "codigo", Text1(Index).Text, "T")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el c�digo de GlobalGap: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        
                        Set frmGlo = New frmManGlobalGap
                        
                        frmGlo.DatosADevolverBusqueda = "0|1|"
                        frmGlo.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        
                        frmGlo.Show vbModal
                        
                        Set frmGlo = Nothing
                        
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    If Modo = 3 Or Modo = 4 Then
                        Sql = "select count(*) from rsocios where codsocio <> " & DBSet(Text1(0).Text, "N") & " and codigoggap = " & DBSet(Text1(Index).Text, "T")
                        
                        If TotalRegistros(Sql) <> 0 Then
                            MsgBox "Este c�digo de GlobalGap ya est� asignado a otro socio. Revise.", vbExclamation
                            PonerFoco Text1(Index)
                        End If
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
    
        Case 27 ' codigo de iban
            Text1(Index).Text = UCase(Text1(Index).Text)
            
        Case 28, 29 ' votos y capital
            PonerFormatoDecimal Text1(Index), 1
            
    End Select
    
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 16 Or Index = 17 Or Index = 18 Or Index = 19 Then
        Dim cta As String
        Dim CC As String
        If Text1(16).Text <> "" And Text1(17).Text <> "" And Text1(18).Text <> "" And Text1(19).Text <> "" Then
            
            cta = Format(Text1(16).Text, "0000") & Format(Text1(17).Text, "0000") & Format(Text1(18).Text, "00") & Format(Text1(19).Text, "0000000000")
            If Len(cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If Text1(27).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Text1(27).Text = "ES" & cta
                Else
                    CC = CStr(Mid(Text1(27).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(Text1(27).Text, 3) <> cta Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
                
            End If
        End If
    End If
    
    
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 20 Then
        If KeyAscii = teclaBuscar Then
            If Modo = 1 Or Modo = 3 Or Modo = 4 Then
                Select Case Index
                    Case 7: KEYFecha KeyAscii, 2 ' fecha baja
                    Case 13: KEYFecha KeyAscii, 0 ' fecha alta
                    Case 14: KEYFecha KeyAscii, 1 ' fecha baja
                    Case 15: KEYBusqueda KeyAscii, 0 'situacion
                    Case 21: KEYBusqueda KeyAscii, 1 'cooperativa
                    Case 26: KEYBusqueda KeyAscii, 2 'codigo de globalgap
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    Else
        If Text1(Index) = "" And KeyAscii = teclaBuscar Then
            imgZoom_Click (Index)
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvan�ar/Retrocedir els camps en les fleches de despla�ament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub



' **** si n'hi han camps de descripci� a la cap�alera ****
Private Sub PosarDescripcions()
Dim NomEmple As String

    On Error GoTo EPosarDescripcions

    Text2(15).Text = PonerNombreDeCod(Text1(15), "rsituacion", "nomsitua", "codsitua", "N")
    Text2(21).Text = PonerNombreDeCod(Text1(21), "rcoope", "nomcoope", "codcoope", "N")
    Text2(26).Text = DevuelveDesdeBDNew(cAgro, "rglobalgap", "descripcion", "codigo", Text1(26).Text, "T")
    If vParamAplic.NumeroConta <> 0 Then
'        lo hemos pasado a lineas
'        Text2(20).Text = PonerNombreDeCod(Text1(20), "tiposiva", "porceiva", "codigiva", "N", cConta)
    End If
    
EPosarDescripcions:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo descripciones", Err.Description
End Sub
' ************************************************************


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

    ModoLineas = 3 'Posem Modo Eliminar Ll�nia
    
    If Modo = 4 Then 'Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'telefonos
            Sql = "�Seguro que desea eliminar el telefono?"
            Sql = Sql & vbCrLf & "Tel�fono: " & AdoAux(Index).Recordset!idtelefono & " - " & AdoAux(Index).Recordset!imei
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rsocios_telefonos"
                Sql = Sql & vWhere & " AND idtelefono= " & DBLet(AdoAux(Index).Recordset!idtelefono, "T")
                
                
                CadenaCambio = "DELETE FROM rsocios_telefonos " & vWhere
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-D", "rsocios_telefonos", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
                
            End If
        Case 1 'secciones
            Sql = "�Seguro que desea eliminar la secci�n?"
            Sql = Sql & vbCrLf & "Secci�n: " & AdoAux(Index).Recordset!codsecci & " - " & AdoAux(Index).Recordset!nomsecci
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rsocios_seccion"
                Sql = Sql & vWhere & " AND codsecci= " & DBLet(AdoAux(Index).Recordset!codsecci, "N")
            
                CadenaCambio = "DELETE FROM rsocios_seccion " & vWhere
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-D", "rsocios_seccion", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            
            End If
        Case 2 'pozos
            Sql = "�Seguro que desea eliminar el registro?"
            Sql = Sql & vbCrLf & "Numero Fase: " & AdoAux(Index).Recordset!numfases
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rsocios_pozos"
                Sql = Sql & vWhere & " AND numfases= " & DBLet(AdoAux(Index).Recordset!numfases, "N")
                
                CadenaCambio = "DELETE FROM rsocios_pozos " & vWhere
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-D", "rsocios_pozos", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            End If
        Case 3 'miembros
            Sql = "�Seguro que desea eliminar el miembro?"
            Sql = Sql & vbCrLf & "NIF: " & AdoAux(Index).Recordset!nifmiembro
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rsocios_miembros"
                Sql = Sql & vWhere & " AND numlinea= " & DBLet(AdoAux(Index).Recordset!numlinea, "N")
                
                CadenaCambio = "DELETE FROM rsocios_miembros " & vWhere
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-D", "rsocios_miembros", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 4 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        ' *** si n'hi han tabs sense datagrid ***
        If Index = 4 Then CargaFrame 3, True
        ' ***************************************
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
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
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Ll�nia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de ll�nies ***
    Select Case Index
        Case 0: vtabla = "rsocios_telefonos"
        Case 1: vtabla = "rsocios_seccion"
        Case 2: vtabla = "rsocios_pozos"
        Case 3: vtabla = "rsocios_miembros"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0 ' *** pose els index dels tabs de ll�nies que tenen datagrid ***
            ' *** canviar la clau primaria de les ll�nies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
                NumF = SugerirCodigoSiguienteStr(vtabla, "idtelefono", vWhere)
            Else
                NumF = ""
            End If
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'cuentas
                    For i = 0 To txtAux.Count - 1
                        txtAux(i).Text = ""
                    Next i
                    txtAux(0).Text = Text1(0).Text 'codsocio
                    txtAux(1).Text = NumF 'idtelefono
                    PonerFoco txtAux(1)
                    
            End Select
         
         Case 1   'secciones
            ' *** canviar la clau primaria de les ll�nies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
'                NumF = SugerirCodigoSiguienteStr(vTabla, "codsecci", vWhere)
'            Else
                NumF = ""
'            End If
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 1 'secciones
                    For i = 0 To txtaux1.Count - 1
                        txtaux1(i).Text = ""
                    Next i
                    txtaux1(0).Text = Text1(0).Text 'codsocio
                    txtaux1(1).Text = NumF 'codseccion
                    txtAux2(1).Text = ""
                    txtAux2(4).Text = ""
                    txtAux2(5).Text = ""
                    txtAux2(0).Text = ""
                    PonerFoco txtaux1(1)
                    
            End Select
         
        Case 2
            ' *** canviar la clau primaria de les ll�nies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
'                NumF = SugerirCodigoSiguienteStr(vTabla, "numfases", vWhere)
'            Else
'                NumF = ""
'            End If
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            For i = 0 To txtaux3.Count - 1
                txtaux3(i).Text = ""
            Next i
            
            txtaux3(0).Text = Text1(0).Text 'codsocio
            txtaux3(1).Text = NumF 'numero de fase
            PonerFoco txtaux3(1)
        
        Case 3
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
                NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            Else
                NumF = ""
            End If
            ' ***************************************************************
            
            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            For i = 0 To txtaux4.Count - 1
                txtaux4(i).Text = ""
            Next i
            
            txtaux4(0).Text = Text1(0).Text 'codsocio
            txtaux4(1).Text = NumF 'numero de fase
            PonerFoco txtaux4(2) 'nif
            
            
    End Select
End Sub


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar ll�nia
       
    If Modo = 4 Then 'Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    ' *********************************
  
    Select Case Index
        Case 0, 1, 2, 3 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
        
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 0 'telefonos
            For i = 0 To 16
                txtAux(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
            
            CargarValoresAnteriores Me, 2, "FrameAux0"
            
        Case 1 'secciones
            For i = 0 To 1
                txtaux1(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
            txtAux2(1).Text = DataGridAux(Index).Columns(2).Text
            For i = 3 To 7
                txtaux1(i - 1).Text = DataGridAux(Index).Columns(i).Text
            Next i
        
            CargarValoresAnteriores Me, 2, "FrameAux1"
        
        
        Case 2 'pozos
            For i = 0 To 3
                txtaux3(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
        
            CargarValoresAnteriores Me, 2, "FrameAux2"
        
        Case 3 'miembros
            For i = 0 To 5
                txtaux4(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
        
            CargarValoresAnteriores Me, 2, "FrameAux3"
        
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'telefonos
            PonerFoco txtAux(2)
        Case 1 'secciones
            PonerFoco txtaux1(2)
            If txtaux1(1).Text <> "" Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(txtaux1(1)) Then
                    If vSeccion.AbrirConta Then
                        If txtaux1(4).Text <> "" Then
                            txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtaux1(4).Text, "T")
                        End If
                        If txtaux1(5).Text <> "" Then
                            txtAux2(5).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtaux1(5).Text, "T")
                        End If
                        If txtaux1(6).Text <> "" Then
                            txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtaux1(6).Text, "N")
                        End If
                    End If
                End If
            End If
        Case 2 ' pozos
            PonerFoco txtaux3(2)
        Case 3 ' miembros
            PonerFoco txtaux4(2)
    End Select
    ' ***************************************************************************************
End Sub

' ***** Si n'hi han combos *****
' per a seleccionar la opcio del combo quan estem modificant; nom�s per a "si" i "no"
'Private Sub SelComboBool(valor As Integer, combo As ComboBox)
'Private Sub SelComboBool(valor, combo As ComboBox)
'    Dim i As Integer
'    Dim j As Integer
'
'    i = valor
'    For j = 0 To combo.ListCount - 1
'        If combo.ItemData(j) = i Then
'            combo.ListIndex = j
'            Exit For
'        End If
'    Next j
'End Sub
' ********************************


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
    Select Case Index
        Case 0 'telefonos
            For jj = 1 To 4
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
        Case 1 'secciones
            For jj = 1 To txtaux1.Count - 1
                txtaux1(jj).visible = b
                txtaux1(jj).Top = alto
            Next jj
            txtAux2(1).visible = b
            txtAux2(1).Top = alto
        
            For jj = 0 To cmdAux.Count - 1
                cmdAux(jj).visible = b
                cmdAux(jj).Top = txtaux1(3).Top
                cmdAux(jj).Height = txtaux1(3).Height
            Next jj
            
        Case 2 ' pozos
            For jj = 1 To 3
                txtaux3(jj).visible = b
                txtaux3(jj).Top = alto
            Next jj
    
        Case 3 ' miembros
            For jj = 2 To 5
                txtaux4(jj).visible = b
                txtaux4(jj).Top = alto
            Next jj
    
    End Select
End Sub




Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
    ImprimirDocumentos
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' Baja de socios
            mnBajaSocio_Click
        Case 2 'Imprimir fases
            mnImprimirFases_Click
    End Select
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
        Case 2 'NIF
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            ValidarNIF txtAux(Index).Text
        
        Case 5 'NOMBRE
            If txtAux(Index).Text <> "" Then txtAux(Index).Text = UCase(txtAux(Index).Text)
    
        Case 12, 13 'ENTIDAD Y SUCURSAL BANCARIA
            PonerFormatoEntero txtAux(Index)
            
        Case 16
            CmdAceptar.SetFocus
    End Select
    
    ' ******************************************************************************
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Not txtAux(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
                Select Case Index
                    Case 4: KEYBusqueda KeyAscii, 7 'pais
                    Case 10: KEYBusqueda KeyAscii, 3 'mercado
                    Case 11: KEYBusqueda KeyAscii, 4 'cadena
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
'    End If
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
    
    
    If b And NumTabMto = 2 And ModoLineas = 1 Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "rsocios_pozos", "acciones", "codsocio", txtaux3(0).Text, "N", , "numfases", txtaux3(1).Text, "N")
        If Sql <> "" Then
            MsgBox "El n�mero de fase ya existe. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtaux3(1)
        End If
    End If
    
    If b And NumTabMto = 3 And ModoLineas = 1 Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "rsocios_miembros", "nifmiembro", "codsocio", txtaux4(0).Text, "N")
        If Sql <> "" Then
            MsgBox "El nif ya existe para este socio. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtaux4(2)
        End If
    End If
    
    
    
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

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

' *** si n'hi han formularis de buscar codi a les ll�nies ***
Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'situacion
            Set frmSit = New frmManSituacion
            frmSit.DatosADevolverBusqueda = "0|1|"
            frmSit.CodigoActual = Text1(15).Text
            frmSit.Show vbModal
            Set frmSit = Nothing
            PonerFoco Text1(15)
        
       Case 1 'cooperativa
            Set frmCoop = New frmManCoope
            frmCoop.DeConsulta = True
            frmCoop.DatosADevolverBusqueda = "0|1|"
            frmCoop.CodigoActual = Text1(21).Text
            frmCoop.Show vbModal
            Set frmCoop = Nothing
            PonerFoco Text1(21)
    
        Case 2 ' codigo globalgap
            indice = 26
            '[Monica]25/04/2012
'            Set frmGlo = New frmBasico
'            AyudaGlobalGap frmGlo, Text1(indice)
            Set frmGlo = New frmManGlobalGap
            
            frmGlo.DeConsulta = True
            frmGlo.DatosADevolverBusqueda = "0|1|"
            frmGlo.CodigoActual = Text1(indice).Text
            frmGlo.Show vbModal
            
            Set frmGlo = Nothing

            Set frmGlo = Nothing
            PonerFoco Text1(indice)
           
    
    End Select
    
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtaux1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo txtaux1(indice)
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nomiva
End Sub

' *********************************************************************************
Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'telefonos
                If DataGridAux(Index).Columns.Count > 2 Then
                    For i = 5 To txtAux.Count - 1
                        txtAux(i).Text = DataGridAux(Index).Columns(i).Text
                    Next i
                    Me.chkAbonos(1).Value = DataGridAux(Index).Columns(17).Text
                    
                End If
            Case 1 'secciones
                If DataGridAux(Index).Columns.Count > 2 Then
                    txtAux2(4).Text = ""
                    txtAux2(5).Text = ""
                    txtAux2(0).Text = ""
                    Set vSeccion = New CSeccion
                    If vSeccion.LeerDatos(AdoAux(1).Recordset!codsecci) Then
                        If vSeccion.AbrirConta Then
                            If DBLet(AdoAux(1).Recordset!codmaccli, "T") <> "" Then
                                txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", AdoAux(1).Recordset!codmaccli, "T")
                            End If
                            If DBLet(AdoAux(1).Recordset!codmacpro, "T") <> "" Then
                                txtAux2(5).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", AdoAux(1).Recordset!codmacpro, "T")
                            End If
                            
                            txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", CStr(AdoAux(1).Recordset!CodIva), "N")
                            
                            vSeccion.CerrarConta
                        End If
                    End If
                    Set vSeccion = Nothing
                End If
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    If numTab = 0 Then
        SSTab1.Tab = 2
    ElseIf numTab = 1 Then
        SSTab1.Tab = 1
    ElseIf numTab = 2 Then
        SSTab1.Tab = 4
    ElseIf numTab = 3 Then
        SSTab1.Tab = 5
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
Dim tip As Integer
Dim i As Byte

    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    AdoAux(Index).Refresh
    
    If Not AdoAux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, AdoAux(Index), 2, "FrameAux" & Index
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        If (Index = 4) Then 'datos facturacion
            tip = AdoAux(Index).Recordset!tipclien
            If (tip = 1) Then 'persona
                txtAux2(27).Text = AdoAux(Index).Recordset!ape_raso & "," & AdoAux(Index).Recordset!Nom_Come
            ElseIf (tip = 2) Then 'empresa
                txtAux2(27).Text = AdoAux(Index).Recordset!Nom_Come
            End If
            txtAux2(28).Text = DBLet(AdoAux(Index).Recordset!desforpa, "T")
            txtAux2(29).Text = DBLet(AdoAux(Index).Recordset!desrutas, "T")
            'txtAux2(31).Text = DBLet(AdoAux(Index).Recordset!comision, "T") & " %"
            txtAux2(32).Text = DBLet(AdoAux(Index).Recordset!nomrapel, "T")
            'Descripcion cuentas contables de la Contabilidad
            For i = 35 To 38
                txtAux2(i).Text = PonerNombreDeCod(txtAux(i), "cuentas", "nommacta", "codmacta", , cConta)
            Next i
        End If
        ' ************************************************************************
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja nom�s lo que te TAG
        txtAux2(0).Text = ""
        txtAux2(1).Text = ""
        
'        txtaux2(27).Text = ""
'        txtaux2(28).Text = ""
'        txtaux2(29).Text = ""
        'txtAux2(31).Text = ""
'        txtaux2(32).Text = ""
'        For i = 35 To 38
'            txtaux2(i).Text = ""
'        Next i
        ' **********************************************************************
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
' ****************************************


Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    'b = DataGridAux(Index).Enabled
    'DataGridAux(Index).Enabled = False
    
'    AdoAux(Index).ConnectionString = Conn
'    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
'    AdoAux(Index).CursorType = adOpenDynamic
'    AdoAux(Index).LockType = adLockPessimistic
'    DataGridAux(Index).ScrollBars = dbgNone
'    AdoAux(Index).Refresh
'    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
'    DataGridAux(Index).AllowRowSizing = False
'    DataGridAux(Index).RowHeight = 290
'    If PrimeraVez Then
'        DataGridAux(Index).ClearFields
'        DataGridAux(Index).ReBind
'        DataGridAux(Index).Refresh
'    End If
'
'    For i = 0 To DataGridAux(Index).Columns.Count - 1
'        DataGridAux(Index).Columns(i).AllowSizing = False
'    Next i
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
    'DataGridAux(Index).Enabled = b
'    PrimeraVez = False
    
    Select Case Index
        Case 0 'telefonos
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtAux(1)|T|Telefono|1400|;" 'codsocio,idtelefono
            tots = tots & "S|txtAux(2)|T|NIF|1400|;"
            tots = tots & "S|txtAux(3)|T|IMEI|3070|;"
            tots = tots & "S|txtAux(4)|T|C.P|900|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(Index).Columns(2).Alignment = dbgLeft
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
            DataGridAux(Index).Columns(4).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'            BloquearTxt txtAux(14), Not b
'            BloquearTxt txtAux(15), Not b

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
                For i = 5 To 16
                    txtAux(i).Text = DataGridAux(Index).Columns(i).Text
                Next i
            Else
                For i = 0 To 16
                    txtAux(i).Text = ""
                Next i
            End If
        
        Case 1 'secciones
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtaux1(1)|T|C�digo|800|;S|cmdAux(4)|B|||;" 'codsocio,codsecci
            tots = tots & "S|txtAux2(1)|T|Nombre|4400|;"
            tots = tots & "S|txtaux1(2)|T|F.Alta|1400|;S|cmdAux(0)|B|||;"
            tots = tots & "S|txtaux1(3)|T|F.Baja|1400|;S|cmdAux(1)|B|||;"
            tots = tots & "S|txtaux1(4)|T|Cta.Cliente|1500|;S|cmdAux(2)|B|||;"
            tots = tots & "S|txtaux1(5)|T|Cta.Proveedor|1500|;S|cmdAux(3)|B|||;"
            tots = tots & "S|txtaux1(6)|T|Iva|800|;S|cmdAux(5)|B|||;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(Index).Columns(2).Alignment = dbgLeft
            DataGridAux(Index).Columns(5).Alignment = dbgLeft
            DataGridAux(Index).Columns(6).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'            BloquearTxt txtAux(14), Not b
'            BloquearTxt txtAux(15), Not b

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), Modo)
'                txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), Modo)
'                txtAux2(0).Text = PonerNombreDeCod(txtaux1(6), "tiposiva", "nombriva", "codigiva", "N", cConta)
            Else
                For i = 0 To 6
                    txtaux1(i).Text = ""
                Next i
                txtAux2(0).Text = ""
                txtAux2(1).Text = ""
                txtAux2(4).Text = ""
                txtAux2(5).Text = ""
            End If
        
        Case 2 'pozos
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtAux3(1)|T|Fases|900|;" 'codsocio,numfase
            tots = tots & "S|txtAux3(2)|T|Acciones|1200|;"
            tots = tots & "S|txtAux3(3)|T|Observaciones|5280|;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

        Case 3 'miembros de entidades
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;S|txtAux4(2)|T|NIF|1250|;" 'codsocio,numfase
            tots = tots & "S|txtAux4(3)|T|Nombre|6450|;"
            tots = tots & "S|txtAux4(4)|T|Votos|1400|;"
            tots = tots & "S|txtAux4(5)|T|Acciones|1500|;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))



    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
        LimpiarCamposFrame Index
    End If
    ' **********************************************************
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub


Private Sub InsertarLinea()
'Inserta registre en les taules de Ll�nies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'telefonos
        Case 1: nomframe = "FrameAux1" 'secciones
        Case 2: nomframe = "FrameAux2" 'pozos
        Case 3: nomframe = "FrameAux3" 'miembros
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            ' *************************************************
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
'            Select Case NumTabMto
'                Case 0: TablaAux = "rsocios_telefonos" 'telefonos
'                Case 1: TablaAux = "rsocios_seccion" 'secciones
'                Case 2: TablaAux = "rsocios_pozos" 'pozos
'            End Select
'
'            '------------------------------------------------------------------------------
'            '  LOG de acciones
'            Set LOG = New cLOG
'            LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-I", TablaAux, ObtenerWhereCab(False)
'            Set LOG = Nothing
'            '-----------------------------------------------------------------------------
           
            
            '++monica: en caso de estar insertando seccion y que no existan las
            'cuentas contables hacemos esto para que las inserte en contabilidad.
            If NumTabMto = 1 Then
               txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 3, Text1(0))
               txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 3, Text1(0))
            End If
            
            If NumTabMto = 1 Then
                If Not vSeccion Is Nothing Then
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                End If
            End If
            
            
            Select Case NumTabMto
                Case 0, 1, 2, 3 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
'                Case 3 ' *** els index dels tabs que NO tenen grid ***
'                    CargaFrame 3, True
'                    If b Then BotonModificar
'                    ModoLineas = 0
'                LLamaLineas NumTabMto, 0
            End Select
           
            SituarTab (NumTabMto)
            
            
        End If
    End If
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Ll�nies
Dim nomframe As String
Dim V As Integer
Dim Cad As String
Dim TablaAux As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'telefonos
        Case 1: nomframe = "FrameAux1" 'secciones
        Case 2: nomframe = "FrameAux2" 'pozos
        Case 3: nomframe = "FrameAux3" 'miembros
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
            End If
            
            Select Case NumTabMto
                Case 0: TablaAux = "rsocios_telefonos" 'telefonos
                Case 1: TablaAux = "rsocios_seccion" 'secciones
                Case 2: TablaAux = "rsocios_pozos" 'pozos
            End Select
    
            '------------------------------------------------------------------------------
            '  LOG de acciones
            Set LOG = New cLOG
            LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-U", TablaAux, ObtenerWhereCab(False)
            Set LOG = Nothing
            '-----------------------------------------------------------------------------
                    
            '[Monica]10/10/2016: si modificamos las ctas las tiene que crear si no existen
            '++monica: en caso de estar modificando seccion y que no existan las
            'cuentas contables hacemos esto para que las inserte en contabilidad.
            If vParamAplic.Cooperativa = 12 Then
                If NumTabMto = 1 Then
                   txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 4, Text1(0))
                   txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 4, Text1(0))
                End If
            End If
            
            If NumTabMto = 1 Then
                If Not vSeccion Is Nothing Then
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                End If
            End If
            
            ' ******************************************************
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModoLineas = 0

            If NumTabMto <> 4 Then
                V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto)

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



Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " codsocio=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' *** neteja els camps dels tabs de grid que
'estan fora d'este, i els camps de descripci� ***
Private Sub LimpiarCamposFrame(Index As Integer)
Dim i As Integer
    On Error Resume Next

    Select Case Index
        Case 0 'telefonos
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
            Next i
        Case 1 'secciones
            For i = 0 To txtaux1.Count - 1
                txtaux1(i).Text = ""
            Next i
        Case 2 'pozos
            For i = 0 To txtaux3.Count - 1
                txtaux3(i).Text = ""
            Next i
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' ***********************************************

Private Sub printNou()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim nTabla As String
Dim ConSubInforme As Boolean

    indRPT = 52 'Impresion de facturas de socios
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    
    ConSubInforme = False
    If InStr(1, nomDocu, "Tur") Then ConSubInforme = True
    
    
    With frmImprimir2
        .cadTabla2 = "rsocios"
        .Informe2 = nomDocu ' "rManSocios.rpt"
        If CadB <> "" Then
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(Data1, Me)
        '[Monica]13/07/2012: falla si hay un solo registro seleccionado y apretamos registros buscados
        If Data1.Recordset.RecordCount = 1 Then .cadRegSelec = .cadRegActua
        .cadTodosReg = ""
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|pOrden={clientes.ape_raso}|"
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|"
        .NumeroParametros2 = 1
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = ConSubInforme 'False
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

' ********* si n'hi han combos a la cap�alera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo irpf
    Combo1(0).AddItem "M�dulos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "E.D."
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Entidad"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    'tipo de produccion
    Combo1(1).AddItem "Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Tercero"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Otra OPA"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    Combo1(1).AddItem "Aportacionista"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    Combo1(1).AddItem "No productor"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 4
   
    'tipo de relacion con la cooperativa
    Combo1(2).AddItem "Socio"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Asociado"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Tercero"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
  
End Sub


Private Sub txtAux1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim cadena As String
    
    If Not PerderFocoGnral(txtaux1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
        Case 1 ' seccion
                If PonerFormatoEntero(txtaux1(Index)) Then
                    Set vSeccion = New CSeccion
                    If vSeccion.LeerDatos(txtaux1(Index)) Then
                        txtAux2(Index).Text = vSeccion.Nombre
                        If vSeccion.AbrirConta Then
                        
                            ' si estamos insertando montamos las cuentas contables con las raices
                            ' y el codigo
                            
                            cadena = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
                            
                            '18/09/2009
                            txtaux1(4).Text = vSeccion.RaizSocio & Format(txtaux1(0).Text, cadena)
                            txtaux1(5).Text = vSeccion.RaizProv & Format(txtaux1(0).Text, cadena)
                        End If
                    Else
                        Set vSeccion = Nothing
                        cadMen = "No existe la Secci�n: " & txtaux1(Index).Text & vbCrLf
                        cadMen = cadMen & "�Desea crearla?" & vbCrLf
                        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                            Set frmSec = New frmManSeccion
                            frmSec.DatosADevolverBusqueda = "0|1|"
                            frmSec.NuevoCodigo = txtaux1(Index).Text
                            txtaux1(Index).Text = ""
                            TerminaBloquear
                            frmSec.Show vbModal
                            Set frmSec = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            txtaux1(Index).Text = ""
                        End If
                    End If
                Else
                    txtaux1(Index).Text = ""
                End If
        
            
        Case 2, 3 'fecha de alta y de baja
            PonerFormatoFecha txtaux1(Index), True
            
        Case 4, 5 'cta Cliente y Proveedor
            If txtaux1(Index).Text = "" Then Exit Sub
            
            If Not vSeccion Is Nothing Then
                txtAux2(Index).Text = PonerNombreCuenta(txtaux1(Index), Modo)
                If txtaux1(Index).Text <> "" Then
                    If Not vSeccion.CtaConRaizCorrecta(txtaux1(Index).Text, Index - 4) Then
                        MsgBox "La cuenta no tiene la raiz correcta. Revise.", vbExclamation
                    Else
                        ' si la cuenta es correcta y no existe la insertamos en contabilidad
                        txtAux2(Index).Text = PonerNombreCuenta(txtaux1(Index), 3, Text1(0))
                    End If
                End If
            End If

        Case 6 'codigo iva
            If txtaux1(Index).Text = "" Then Exit Sub
            
            If Not vSeccion Is Nothing Then
                  txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtaux1(Index).Text, "N")
            End If
            CmdAceptar.SetFocus

    End Select
    
    ' ******************************************************************************
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
   If Not txtaux1(Index).MultiLine Then ConseguirFocoLin txtaux1(Index)
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtaux1(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 1:  KEYImage KeyAscii, 4
            Case 2:  KEYImage KeyAscii, 0
            Case 3:  KEYImage KeyAscii, 1
            Case 4:  KEYImage KeyAscii, 2
            Case 5:  KEYImage KeyAscii, 3
            Case 6:  KEYImage KeyAscii, 5
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYImage(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    cmdAux_Click (indice)
End Sub



'??????????????????????????
Private Sub TxtAux3_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim cadena As String
    
    If Not PerderFocoGnral(txtaux3(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
        Case 1 ' numfases
            PonerFormatoEntero txtaux3(Index)
            
        Case 2
            PonerFormatoDecimal txtaux3(Index), 10
        
        Case 3 'observaciones
            CmdAceptar.SetFocus

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

'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'
'
'       El listview tendra los datos de albaranes, facturas... que tenga el cliente
'       Con lo cual, a partir de un click tendremos que ser capaces de situarnos en
'       el formulario correspondiente
'
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------

Private Sub ImagenesNavegacion()
    With Me.Toolbar2
'        .ImageList = frmPpal.imgListImages16
        .ImageList = frmPpal.imgListPpal
        .Buttons(1).Image = 23
        .Buttons(3).Image = 30
        .Buttons(5).Image = 25
        .Buttons(7).Image = 22
        .Buttons(9).Image = 38
        .Buttons(11).Image = 24
        
    End With
    With Me.Toolbar3
        .ImageList = frmPpal.imgListImages16
        .Buttons(1).Image = 5
        .Buttons(3).Image = 7
        .Buttons(5).Image = 6
        .Buttons(7).Image = 8
    End With
    With Me.Toolbar4
        .ImageList = frmPpal.imgListComun16
        
'        .ImageList = frmPpal.imgListPpal
        .Buttons(1).Image = 10
    End With
    
'    Set lw1.SmallIcons = frmPpal.imgListPpal
    Set lw1.SmallIcons = frmPpal.imgListImages16
    
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Tag = "" Then Exit Sub
    
    If Button.Index = 9 Then
        ImprimirDocumentos
        Exit Sub
    End If
    
    
    Label16.Caption = ""
    'Levantamos todos los botones y dejamos pulsado el de ahora
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Tag <> "" Then
            If Toolbar2.Buttons(NumRegElim).Index <> Button.Index Then Toolbar2.Buttons(NumRegElim).Value = tbrUnpressed
        End If
    Next NumRegElim
    CargaColumnas CByte(Button.Tag)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLW
End Sub





Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim c As ColumnHeader

    Frame8.visible = False

    Select Case OpcionList
    Case 0
        'Facturas
        Label16.Caption = "Facturas"
        Columnas = "Tipo|Numero|Fecha|Importe|"
        Ancho = "1000|2000|1200|3600|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|"
        'Formatos
        Formato = "|00000000|dd/mm/yyyy|" & FormatoImporte & "|"
        Ncol = 4
               
    Case 1
        'Entradas por Huerto
        Label16.Caption = "Entradas por Huerto"
        Columnas = "Huerto|Nro.Orden|Partida|Variedad|Kilos Netos|"
        Ancho = "1200|1200|2400|1900|1700|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|"
        'Formatos
        Formato = "|0|0|0|###,###,##0|"
        Ncol = 5
        
    Case 2
        'Detalles de Entradas
        Label16.Caption = "Detalles de Entradas"
        If vParamAplic.Cooperativa = 12 Then
            Columnas = "Albar�n|Fecha|Variedad|Cajas|Kilos Netos|"
            Ancho = "1400|1400|2000|1500|2000|"
            'vwColumnRight =1  left=0   center=2
            Alinea = "0|0|0|1|1|"
            'Formatos
            Formato = "|dd/mm/yyyy|0|###,##0|###,###,##0|"
            Ncol = 5
        Else
            Columnas = "Albar�n|Fecha|Campo|N.Orden|Partida|Variedad|Kilos Netos|"
            Ancho = "1100|1400|1100|1000|1300|1300|1200|"
            'vwColumnRight =1  left=0   center=2
            Alinea = "0|0|0|0|0|0|1|"
            'Formatos
            Formato = "|dd/mm/yyyy|0|0|0|0|###,###,##0|"
            Ncol = 7
        End If
        
    Case 4
        'Facturas adv , retirada almazara y bodega , recibos de pozos
        Label16.Caption = "Facturas Cliente"
        Columnas = "Tipo|Factura|Fecha|Importe|Cobrado|"
        Ancho = "1000|2000|1200|2600|1000|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|0|"
        'Formatos
        Formato = "|00000000|dd/mm/yyyy|" & FormatoImporte & "||"
        Ncol = 5
        
    Case 5
        ' Documentos
        Label16.Caption = "Im�genes"
        Columnas = "C�digo|Nombre|Documento|Id|Tipo|"
        Ancho = "1000|6000|0|0|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "|||||"
        Ncol = 5
        
        Frame8.visible = True
        
    End Select
    
    
'    'Fecha incio busquedas
'    Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
    'Guardo la opcion en el tag
    lw1.Tag = OpcionList & "|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set c = lw1.ColumnHeaders.Add()
         c.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         c.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         c.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         c.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
End Sub


Private Sub CargaDatosLW()
Dim c As String
Dim bs As Byte
    bs = Screen.MousePointer
    c = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & Label16.Caption
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = c
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLW2()
Dim Cad As String
Dim Rs As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Orden As String
Dim c As String


    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar3.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    'Fecha incio busquedas
    Text3(0).Text = Format(imgFec(3).Tag, "dd/mm/yyyy")
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0
        'FACTURAS
        Cad = "select h.codtipom,h.numfactu,h.fecfactu,h.totalfac from rfactsoc h WHERE 1=1"
        Cad = Cad & " and h.codsocio=" & Data1.Recordset!Codsocio
        GroupBy = "1,2,3"
        BuscaChekc = "h.fecfactu"
        'La fecha
        If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
        
    Case 1
        'ENTRADAS POR HUERTO
        Cad = "select c.codcampo,c.nrocampo,p.nomparti,v.nomvarie,sum(h.kilosnet) "
        Cad = Cad & " from ((rcampos c left join rhisfruta h on c.codcampo = h.codcampo "
        Cad = Cad & " and c.codsocio = h.codsocio and c.codvarie = h.codvarie "
        
        'La fecha
        If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
        
        Cad = Cad & " )"
        Cad = Cad & " inner join rpartida p on c.codparti = p.codparti) "
        Cad = Cad & " inner join variedades v on c.codvarie = v.codvarie "
        Cad = Cad & " where c.codsocio=" & Data1.Recordset!Codsocio
        Cad = Cad & " and c.fecbajas is null"
        
        GroupBy = "1,2,3,4"
        BuscaChekc = "h.fecalbar"
        Orden = "c.codcampo"
    Case 2
        'DETALLE DE ENTRADAS
        If vParamAplic.Cooperativa = 12 Then
            Cad = "select h.numalbar,h.fecalbar,v.nomvarie,h.numcajon, h.kilosnet from rhisfruta h, variedades v WHERE "
            Cad = Cad & " h.codvarie = v.codvarie and "
            Cad = Cad & " h.codsocio=" & Data1.Recordset!Codsocio
            
            BuscaChekc = "h.fecalbar"
            'La fecha
            If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
        
            GroupBy = "1,2,3,4,5"

        Else
            Cad = "select h.numalbar,h.fecalbar,h.codcampo,c.nrocampo,p.nomparti,v.nomvarie,h.kilosnet from rhisfruta h, rcampos c, rpartida p, variedades v WHERE "
            Cad = Cad & " h.codcampo=c.codcampo and h.codsocio=c.codsocio and h.codvarie=c.codvarie and "
            Cad = Cad & " h.codvarie = v.codvarie and "
            Cad = Cad & " c.codparti=p.codparti "
            Cad = Cad & " and h.codsocio=" & Data1.Recordset!Codsocio
            
            BuscaChekc = "h.fecalbar"
            'La fecha
            If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
            
            GroupBy = "1,2,3,4,5,6,7"
            
        End If
        
    
    Case 4
        'FACTURAS de cliente (advfacturas, rbodfacturas, rrecibpozos)
        'advfacturas
        Dim ContaADV As Integer
        ContaADV = DevuelveValor("select empresa_conta from rseccion where codsecci = " & vParamAplic.SeccionADV)
        
        If ContaADV <> 0 Then
        
        
            Cad = "select h.codtipom,h.numfactu,h.fecfactu,h.totalfac totalfac, if(sum(coalesce(c1.impvenci,0)+coalesce(c1.gastos,0)-coalesce(c1.impcobro,0))=0,'Cobrado','')  "
            Cad = Cad & " from (advfacturas h inner join usuarios.stipom u1 on h.codtipom = u1.codtipom) left join conta" & ContaADV & ".scobro c1 on c1.numserie = u1.letraser "
            Cad = Cad & " and c1.codfaccl = h.numfactu "
            Cad = Cad & " and c1.fecfaccl = h.fecfactu "
            Cad = Cad & " WHERE 1=1"
            Cad = Cad & " and h.codsocio=" & Data1.Recordset!Codsocio
            BuscaChekc = "h.fecfactu"
            'La fecha
            If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
            Cad = Cad & " GROUP BY 1,2,3,4 "
            
            GroupBy = "1,2,3,4"
            
            
        
        End If
        
        'rbodfacturas
        Dim ContaBOD As Integer
        ContaBOD = DevuelveValor("select empresa_conta from rseccion where codsecci = " & vParamAplic.SeccionBodega)
        
        If ContaBOD <> 0 Then
        
            If Cad <> "" Then Cad = Cad & " union "
        
        
            Cad = Cad & "select i.codtipom,i.numfactu,i.fecfactu,i.totalfac totalfac, if(sum(coalesce(c2.impvenci,0)+coalesce(c2.gastos,0)-coalesce(c2.impcobro,0))=0,'Cobrado','')  "
            Cad = Cad & " from (rbodfacturas i inner join usuarios.stipom u2 on i.codtipom = u2.codtipom) left join conta" & ContaBOD & ".scobro c2 on c2.numserie = u2.letraser "
            Cad = Cad & " and c2.codfaccl = i.numfactu "
            Cad = Cad & " and c2.fecfaccl = i.fecfactu "
            Cad = Cad & " WHERE 1=1"
            Cad = Cad & " and i.codsocio=" & Data1.Recordset!Codsocio
            BuscaChekc = "i.fecfactu"
            'La fecha
            If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
            
            Cad = Cad & " GROUP BY 1,2,3,4"
            
            GroupBy = "1,2,3,4"
            
    
        End If
    
        'rrecibpozos
        Dim ContaPOZOS As Integer
        ContaPOZOS = DevuelveValor("select empresa_conta from rseccion where codsecci = " & vParamAplic.SeccionPOZOS)
        
        If ContaPOZOS <> 0 Then
        
            If Cad <> "" Then Cad = Cad & " union "
        
        
        
            Cad = Cad & "select j.codtipom,j.numfactu,j.fecfactu,j.totalfact totalfac, if(sum(coalesce(c.impvenci,0)+coalesce(c.gastos,0)-coalesce(c.impcobro,0))=0,'Cobrado','') "
            Cad = Cad & " from (rrecibpozos j inner join usuarios.stipom u on j.codtipom = u.codtipom)  left join conta" & ContaPOZOS & ".scobro c on c.numserie = u.letraser "
            Cad = Cad & " and c.codfaccl = j.numfactu "
            Cad = Cad & " and c.fecfaccl = j.fecfactu "
            Cad = Cad & " WHERE 1=1"
            Cad = Cad & " and j.codsocio=" & Data1.Recordset!Codsocio
            BuscaChekc = "j.fecfactu"
            'La fecha
            If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
            
            GroupBy = "1,2,3,4"

        End If

        BuscaChekc = "1,2,3"
        
    Case 5 ' imagenes
        Cad = "select h.orden, h.descripfich, h.campo, h.codigo, h.docum from rfichdocs h WHERE "
        Cad = Cad & " codsocio=" & Data1.Recordset!Codsocio
        GroupBy = ""
        BuscaChekc = "orden"
        
    End Select
    
    
    'El group by
    If GroupBy <> "" Then Cad = Cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
    If CByte(RecuperaValor(lw1.Tag, 1)) = 1 Then BuscaChekc = Orden
    
    'BuscaChekc="" si es la opcion de precios especiales
    If Cad <> "" Then Cad = Cad & " ORDER BY " & BuscaChekc & " DESC"
    BuscaChekc = ""
    
    lw1.ListItems.Clear
    
    
    If CByte(RecuperaValor(lw1.Tag, 1)) = 5 Then
        
        CargarArchivos
    
    Else
        If Cad <> "" Then
            Set Rs = New ADODB.Recordset
            Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                Set It = lw1.ListItems.Add()
                If lw1.ColumnHeaders(1).Tag <> "" Then
                    It.Text = Format(Rs.Fields(0), lw1.ColumnHeaders(1).Tag)
                Else
                    It.Text = Rs.Fields(0)
                End If
                'El resto de cmpos
                For NumRegElim = 2 To CInt(RecuperaValor(lw1.Tag, 2))
                    If IsNull(Rs.Fields(NumRegElim - 1)) Then
                        It.SubItems(NumRegElim - 1) = " "
                    Else
                        If lw1.ColumnHeaders(NumRegElim).Tag <> "" Then
                            It.SubItems(NumRegElim - 1) = Format(Rs.Fields(NumRegElim - 1), lw1.ColumnHeaders(NumRegElim).Tag)
                        Else
                            It.SubItems(NumRegElim - 1) = Rs.Fields(NumRegElim - 1)
                        End If
                    End If
                    
                Next
                It.SmallIcon = ElIcono
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        End If
    End If
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set Rs = Nothing
    
End Sub

Private Sub CargarArchivos()
Dim c As String
Dim L As Long
Dim Rs As ADODB.Recordset
Dim nFile As Long


    ProcesarCarpetaImagenes


    c = "Select * from rfichdocs where codsocio=" & DBSet(Text1(0).Text, "N") & " ORDER BY orden"
'    Me.lblCarga2.Caption = "Leyendo desde BD "
'    Me.lblCarga2.Refresh
    adodc1.ConnectionString = conn
    adodc1.RecordSource = c
    adodc1.Refresh

    If adodc1.Recordset.EOF Then
        'NO HAY NINGUNA
        CargarIMG ""
    Else
        'LEEMOS LAS IMAGENES
'        InsertandoImg = True
        While Not adodc1.Recordset.EOF
            L = adodc1.Recordset!Codigo
'            Me.lblCarga2.Caption = "Leyendo desde BD " & L & "       " & adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
'            lblCarga2.Refresh
            c = App.Path & "\" & CarpetaIMG & "\" & L
            If DBLet(adodc1.Recordset!Docum) <> "0" Then
                c = App.Path & "\" & CarpetaIMG & "\" & adodc1.Recordset!Docum
            End If
            If Dir(c) <> "" Then
                AnyadirAlListview c, True
            Else
                If LeerBinary(adodc1.Recordset!campo, c) Then
                    AnyadirAlListview c, True
                End If
            End If
            adodc1.Recordset.MoveNext
        Wend
    
    
        
'        InsertandoImg = False
        If lw1.ListItems.Count > 0 Then CargarIMG lw1.ListItems(1).SubItems(2)
    End If

    Set adodc1.Recordset = Nothing
End Sub

Private Sub AnyadirAlListview(vpaz As String, DesdeBD As Boolean)
Dim J As Integer
Dim Aux As String
Dim It As ListItem
Dim Contador As Integer
    If Dir(vpaz, vbArchive) = "" Then
        MsgBox "No existe el archivo: " & vpaz, vbExclamation
    Else
        'List1.AddItem vpaz
        Set It = lw1.ListItems.Add()
'        It.SmallIcon = 23
        
'        If DesdeBD Then
'            J = InStrRev(vpaz, "\") + 1
'            Aux = Mid(vpaz, J)
'            It.Text = "C�digo " & Aux
'            If Not IsNumeric(Aux) Then It.SmallIcon = 9
'            It.SubItems(2) = Aux
'
'        Else
'            Contador = Contador + 1
            It.Text = Me.adodc1.Recordset!Orden '"Nuevo " & Contador
'        End If
        
        It.SubItems(1) = Me.adodc1.Recordset.Fields(3)  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
        It.SubItems(2) = vpaz
        It.SubItems(3) = Me.adodc1.Recordset.Fields(0)
        Set It = Nothing
    End If
End Sub


Private Function CargarIMG(Archivo As String) As Boolean
    On Error Resume Next
    Screen.MousePointer = vbHourglass
'    lblCarga2.Caption = "Cargando ..."
'    lblCarga2.Refresh
    CargarIMG = False
    
    If InStr(1, Archivo, ".pdf") <> 0 Then
        Me.Image1.Picture = LoadPicture(App.Path & "\pdf.dat")
    Else
        If InStr(1, Archivo, ".tif") <> 0 Then
            Me.Image1.Picture = LoadPicture(App.Path & "\tif.dat")
        Else
            If InStr(1, Archivo, ".png") Then
                Me.Image1.Picture = LoadPicture(App.Path & "\png.dat")
            Else
                Me.Image1.Picture = LoadPicture(Archivo)
            End If
        End If
    End If

    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
    Else
        CargarIMG = True
    End If
'    lblCarga2.Caption = lblCarga2.Tag
    Screen.MousePointer = vbDefault
End Function



Private Sub ImprimirDocumentos()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim Industria As Boolean
Dim Cad As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
'    cadparam = cadparam & "pSocio=" & Data1.Recordset!codsocio & "|"
'    numParam = numParam + 1
'
    CadParam = CadParam & "Fecha=Date(""" & Text3(0).Text & """)|"
    numParam = numParam + 1
    
    CadParam = CadParam & "pDetalleAlb=" & Check1.Value & "|"
    numParam = numParam + 1
    
    CadParam = CadParam & "pImpFactura=" & Check2.Value & "|"
    numParam = numParam + 1
    
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Cad = "select count(*) "
        Cad = Cad & " from rsocios p "
        Cad = Cad & " where p.codsocio=" & Data1.Recordset!Codsocio
        Cad = Cad & " and p.fechabaja is null"
'        If Text3(0).Text <> "" Then Cad = Cad & " and h.fecfactu >='" & Format(Text3(0).Text, FormatoFecha) & "'"
    
    Else
        Cad = "select count(*) "
        Cad = Cad & " from ((rcampos c left join rhisfruta h on c.codcampo = h.codcampo "
        Cad = Cad & " and c.codsocio = h.codsocio and c.codvarie = h.codvarie "
        If Text3(0).Text <> "" Then Cad = Cad & " and h.fecalbar >='" & Format(Text3(0).Text, FormatoFecha) & "'"
        Cad = Cad & " )"
        Cad = Cad & " inner join rpartida p on c.codparti = p.codparti) "
        Cad = Cad & " inner join variedades v on c.codvarie = v.codvarie "
        Cad = Cad & " where c.codsocio=" & Data1.Recordset!Codsocio
        Cad = Cad & " and c.fecbajas is null"
    End If
        
    If TotalRegistros(Cad) = 0 Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        Exit Sub
    End If
        
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        If Not AnyadirAFormula(cadFormula, "{rsocios.codsocio}=" & Data1.Recordset!Codsocio) Then Exit Sub
    Else
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.codsocio}=" & Data1.Recordset!Codsocio) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rhisfruta.fecalbar}>=Date(""" & Text3(0).Text & """)") Then Exit Sub
    End If
    indRPT = 40 'Impresion de Factura Socio
    ConSubInforme = True
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    cadTitulo = "Resumen de Documentos Socio"
        
    LlamarImprimir
End Sub

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .ConSubInforme = ConSubInforme
        .Show vbModal
    End With
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub

Private Sub ProcesarCarpetaImagenes()
Dim c As String
Dim MiNombre As String

    On Error GoTo EProcesarCarpetaImagenes
    c = App.Path & "\" & CarpetaIMG
    If Dir(c, vbDirectory) = "" Then
        MkDir c
    Else
        On Error Resume Next
        If Dir(c & "\*.*", vbArchive) <> "" Then 'Kill c & "\*.*"
            MiNombre = Dir(c & "\*.*")   ' Recupera la primera entrada.
            Do While MiNombre <> ""   ' Inicia el bucle.
               ' Ignora el directorio actual y el que lo abarca.
               If MiNombre <> "." And MiNombre <> ".." Then
                    Kill c & "\" & MiNombre
               End If
               MiNombre = Dir   ' Obtiene siguiente entrada.
            Loop
        End If
        On Error GoTo EProcesarCarpetaImagenes
    
    End If
    
    Exit Sub
EProcesarCarpetaImagenes:
    MuestraError Err.Number, "ProcesarCarpetaImagenes"
End Sub



Private Sub ModificarDatosCuentaContable()
Dim Sql As String
Dim Cad As String
Dim vSeccion As CSeccion
Dim vSocio As cSocio
Dim Cuentas As String
Dim Sql1 As String
Dim Sql2 As String
Dim Nregs As Long
Dim Rs As ADODB.Recordset
Dim vIban As String

    On Error GoTo eModificarDatosCuentaContable


    If Text1(1).Text <> NombreAnt Or Text1(16).Text <> BancoAnt Or Text1(17).Text <> SucurAnt Or Text1(18).Text <> DigitoAnt Or Text1(19).Text <> CuentaAnt Or _
       DirecAnt <> Text1(3).Text Or cPostalAnt <> Text1(4).Text Or PoblaAnt <> Text1(5).Text Or ProviAnt <> Text1(6).Text Or NifAnt <> Text1(2).Text Or _
       EMaiAnt <> Text1(12).Text Or _
       IbanAnt <> Text1(27).Text Then
        
        Cad = "Se han producido cambios en los siguientes datos del Socio: " & vbCrLf
        
        If NombreAnt <> Text1(1).Text Then Cad = Cad & " Nombre,"
        If DirecAnt <> Text1(3).Text Then Cad = Cad & " Direccion,"
        If cPostalAnt <> Text1(4).Text Then Cad = Cad & " CPostal,"
        If PoblaAnt <> Text1(5).Text Then Cad = Cad & " Poblaci�n,"
        If ProviAnt <> Text1(6).Text Then Cad = Cad & " Provincia,"
        If NifAnt <> Text1(2).Text Then Cad = Cad & " NIF,"
        If EMaiAnt <> Text1(12).Text Then Cad = Cad & " EMail,"
        
        If IbanAnt <> Text1(27).Text Then Cad = Cad & " Iban,"
        
        If BancoAnt <> Text1(16).Text Then Cad = Cad & " Banco,"
        If SucurAnt <> Text1(17).Text Then Cad = Cad & " Sucursal,"
        If DigitoAnt <> Text1(18).Text Then Cad = Cad & " Dig.Control,"
        If CuentaAnt <> Text1(19).Text Then Cad = Cad & " Cuenta banco,"
        
        Cad = Mid(Cad, 1, Len(Cad) - 1)
        
        Cad = Cad & vbCrLf & vbCrLf & "� Desea actualizar los datos en la Contabilidad de la Secci�n Horto ?" & vbCrLf & vbCrLf
        
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Set vSocio = New cSocio
            If vSocio.LeerDatosSeccion(Text1(0).Text, vParamAplic.Seccionhorto) Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                    If vSeccion.AbrirConta Then
                        ConnConta.BeginTrans
                        
                        Sql = "update cuentas set nommacta = " & DBSet(Trim(Text1(1).Text), "T")
                        Sql = Sql & ", razosoci = " & DBSet(Trim(Text1(1).Text), "T")
                        Sql = Sql & ", dirdatos = " & DBSet(Trim(Text1(3).Text), "T")
                        Sql = Sql & ", codposta = " & DBSet(Trim(Text1(4).Text), "T")
                        Sql = Sql & ", despobla = " & DBSet(Trim(Text1(5).Text), "T")
                        Sql = Sql & ", desprovi = " & DBSet(Trim(Text1(6).Text), "T")
                        Sql = Sql & ", nifdatos = " & DBSet(Trim(Text1(2).Text), "T")
                        Sql = Sql & ", maidatos = " & DBSet(Trim(Text1(12).Text), "T")
                        
                        If Not vParamAplic.ContabilidadNueva Then
                            Sql = Sql & ", entidad = " & DBSet(Trim(Text1(16).Text), "T", "S")
                            Sql = Sql & ", oficina = " & DBSet(Trim(Text1(17).Text), "T", "S")
                            Sql = Sql & ", cc = " & DBSet(Trim(Text1(18).Text), "T", "S")
                            Sql = Sql & ", cuentaba = " & DBSet(Trim(Text1(19).Text), "T", "S")
                            '[Monica]22/11/2013: tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                Sql = Sql & ", iban = " & DBSet(Trim(Text1(27).Text), "T", "S")
                            End If
                        Else
                            vIban = MiFormat(Text1(27).Text, "") & MiFormat(Text1(16).Text, "0000") & MiFormat(Text1(17).Text, "0000") & MiFormat(Text1(18).Text, "00") & MiFormat(Text1(19).Text, "0000000000")
                        
                            Sql = Sql & ", iban = " & DBSet(vIban, "T")
                            Sql = Sql & ", codpais = 'ES' "
                        
                        End If
                        Sql = Sql & " where codmacta = "
                        
                        Cuentas = ""
                        
                        If vSocio.CtaClien <> "" Then
                            ConnConta.Execute Sql & DBSet(Trim(vSocio.CtaClien), "T")
                            Cuentas = Cuentas & DBSet(Trim(vSocio.CtaClien), "T") & ","
                        End If
                        If vSocio.CtaProv <> "" Then
                            ConnConta.Execute Sql & DBSet(Trim(vSocio.CtaProv), "T")
                            Cuentas = Cuentas & DBSet(Trim(vSocio.CtaProv), "T") & ","
                        End If
                        
                        'quitamos la ultima coma de las cuentas contables que hemos de modificar
                        If Cuentas <> "" Then Cuentas = Mid(Cuentas, 1, Len(Cuentas) - 1)
                        
                        '[Monica]30/08/2013: si han cambiado los datos de la cuenta del banco y hay cobros/pagos pendientes
                        '                    pregunto si quieren cambiarlos en tesoreria
                        If (BancoAnt <> Text1(16).Text Or SucurAnt <> Text1(17).Text Or DigitoAnt <> Text1(18).Text Or CuentaAnt <> Text1(19).Text Or _
                            IbanAnt <> Text1(27).Text) _
                           And Cuentas <> "" Then
                           
                            If Not vParamAplic.ContabilidadNueva Then
                            
                                Sql1 = "select sum(total) from ("
                                Sql1 = Sql1 & "select  count(*) total "
                                Sql1 = Sql1 & " from scobro  cc "
                                '[Monica]14/11/2014: si es escalona o utxera no miramos la situacion de la remesa
                                If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                                    Sql1 = Sql1 & " where cc.codmacta  in (" & Cuentas & ")  "
                                Else
                                    '[Monica]09/09/2014: la remesa no tiene que estar contabilizada. Lo cambiamos
                                    'Sql1 = Sql1 & " where cc.codmacta  in (" & Cuentas & ")  and (cc.codrem is null or cc.codrem = 0) and (cc.transfer is null or cc.transfer = 0)"
                                    Sql1 = Sql1 & " where cc.codmacta  in (" & Cuentas & ")  and (cc.siturem is null or cc.siturem <= 'Q')"
                                End If
                                Sql1 = Sql1 & " union "
                                Sql1 = Sql1 & " select count(*) total "
                                Sql1 = Sql1 & " from spagop pp "
                                '[Monica]09/09/2014: en los pagos cuando se contabiliza la transferencia desaparece el pago. Lo cambiamos
                                'Sql1 = Sql1 & " where pp.ctaprove in (" & Cuentas & ") and  (pp.transfer is null or pp.transfer = 0)"
                                Sql1 = Sql1 & " where pp.ctaprove in (" & Cuentas & ") "
                                Sql1 = Sql1 & ") aaaaaa "

                            Else
                                Sql1 = "select sum(total) from ("
                                Sql1 = Sql1 & "select  count(*) total "
                                Sql1 = Sql1 & " from cobros  cc "
                                '[Monica]14/11/2014: si es escalona o utxera no miramos la situacion de la remesa
                                If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                                    Sql1 = Sql1 & " where cc.codmacta  in (" & Cuentas & ")  "
                                Else
                                    '[Monica]09/09/2014: la remesa no tiene que estar contabilizada. Lo cambiamos
                                    'Sql1 = Sql1 & " where cc.codmacta  in (" & Cuentas & ")  and (cc.codrem is null or cc.codrem = 0) and (cc.transfer is null or cc.transfer = 0)"
                                    Sql1 = Sql1 & " where cc.codmacta  in (" & Cuentas & ")  and (cc.siturem is null or cc.siturem <= 'Q')"
                                End If
                                Sql1 = Sql1 & " union "
                                Sql1 = Sql1 & " select count(*) total "
                                Sql1 = Sql1 & " from pagos pp "
                                '[Monica]09/09/2014: en los pagos cuando se contabiliza la transferencia desaparece el pago. Lo cambiamos
                                'Sql1 = Sql1 & " where pp.ctaprove in (" & Cuentas & ") and  (pp.transfer is null or pp.transfer = 0)"
                                Sql1 = Sql1 & " where pp.codmacta in (" & Cuentas & ") "
                                Sql1 = Sql1 & ") aaaaaa "
                            End If

                            Nregs = 0

                            Set Rs = New ADODB.Recordset
                            Rs.Open Sql1, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                            If Not Rs.EOF Then
                                If DBLet(Rs.Fields(0).Value) <> 0 Then Nregs = Rs.Fields(0).Value
                            End If
                            Set Rs = Nothing
                            
                            If Nregs <> 0 Then
                                Cad = "Se han producido cambios en la Cta.Bancaria del Socio."
                                Cad = Cad & vbCrLf & vbCrLf & "� Desea actualizar los Cobros y Pagos pendientes en Tesoreria ?" & vbCrLf & vbCrLf
                                
                                If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                                    If Not vParamAplic.ContabilidadNueva Then
                                         Sql2 = "update scobro set codbanco = " & DBSet(Text1(16).Text, "N", "S") & ", codsucur = " & DBSet(Text1(17).Text, "N", "S")
                                         Sql2 = Sql2 & ", digcontr = " & DBSet(Text1(18).Text, "T", "S") & ", cuentaba = " & DBSet(Text1(19).Text, "T", "S")
                                         
                                         '[Monica]22/11/2013: tema iban
                                         If vEmpresa.HayNorma19_34Nueva = 1 Then
                                            Sql2 = Sql2 & ", iban = " & DBSet(Text1(27).Text, "T", "S")
                                         End If
                                         
                                         Sql2 = Sql2 & " where codmacta in (" & Cuentas & ") "
                                        
                                        '[Monica]14/11/2014: si es escalona o utxera no miramos la situacion de la remesa
                                         If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                                         
                                         Else
                                            '[Monica]09/09/2014: la remesa no tiene que estar contabilizada. Lo cambiamos
                                            'Sql2 = Sql2 & " and (codrem is null or codrem = 0) and (transfer is null or transfer = 0)"
                                            Sql2 = Sql2 & " and (siturem is null or siturem <= 'Q')"
                                         End If
                                         
                                         ConnConta.Execute Sql2
                                         
                                         Sql2 = "update spagop set entidad = " & DBSet(Text1(16).Text, "T", "S") & ", oficina = " & DBSet(Text1(17).Text, "T", "S")
                                         Sql2 = Sql2 & ", cc = " & DBSet(Text1(18).Text, "T", "S") & ", cuentaba = " & DBSet(Text1(19).Text, "T", "S")
                                         
                                         '[Monica]22/11/2013: tema iban
                                         If vEmpresa.HayNorma19_34Nueva = 1 Then
                                            Sql2 = Sql2 & ", iban = " & DBSet(Text1(27).Text, "T", "S")
                                         End If
                                         
                                         Sql2 = Sql2 & " where ctaprove in (" & Cuentas & ") "
                                         '[Monica]09/09/2014: la remesa no tiene que estar contabilizada. Lo cambiamos
                                         'Sql2 = Sql2 & " and (transfer is null or transfer = 0)"
                                        
                                         ConnConta.Execute Sql2
                                     Else
                                         vIban = MiFormat(Text1(27).Text, "") & MiFormat(Text1(16).Text, "0000") & MiFormat(Text1(17).Text, "0000") & MiFormat(Text1(18).Text, "00") & MiFormat(Text1(19).Text, "0000000000")
                                    
                                     
                                         Sql2 = "update cobros set iban = " & DBSet(vIban, "T")
                                         Sql2 = Sql2 & " where codmacta in (" & Cuentas & ") "
                                        
                                        '[Monica]14/11/2014: si es escalona o utxera no miramos la situacion de la remesa
                                         If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                                         
                                         Else
                                            '[Monica]09/09/2014: la remesa no tiene que estar contabilizada. Lo cambiamos
                                            'Sql2 = Sql2 & " and (codrem is null or codrem = 0) and (transfer is null or transfer = 0)"
                                            Sql2 = Sql2 & " and (siturem is null or siturem <= 'Q')"
                                         End If
                                         
                                         ConnConta.Execute Sql2
                                         
                                         Sql2 = "update pagos set iban = " & DBSet(vIban, "T")
                                         Sql2 = Sql2 & " where ctaprove in (" & Cuentas & ") "
                                         '[Monica]09/09/2014: la remesa no tiene que estar contabilizada. Lo cambiamos
                                         'Sql2 = Sql2 & " and (transfer is null or transfer = 0)"
                                        
                                         ConnConta.Execute Sql2
                                     
                                     
                                     End If
                                End If
                            End If
                        
                        End If
                        
                        ConnConta.CommitTrans
                   End If
                End If
                Set vSeccion = Nothing
            End If
            Set vSocio = Nothing
        End If
    End If
    Exit Sub
    
eModificarDatosCuentaContable:
    ConnConta.RollbackTrans
    MuestraError Err.Number, "Modificar Datos Cuenta Contable", Err.Description
End Sub

'**************************************************
'******** M I E M B R O S
'**************************************************
Private Sub TxtAux4_GotFocus(Index As Integer)
   If Not txtaux4(Index).MultiLine Then ConseguirFocoLin txtaux4(Index)
End Sub

Private Sub TxtAux4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtaux4(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux4_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux4_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim cadena As String
    
    If Not PerderFocoGnral(txtaux4(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
        Case 2 ' nif
            txtaux4(Index).Text = UCase(txtaux4(Index).Text)
            ValidarNIF txtaux4(Index).Text
        
        Case 4 ' votos
            PonerFormatoDecimal txtaux4(Index), 1
        
        Case 5 ' capital
            If PonerFormatoDecimal(txtaux4(Index), 1) Then
                CmdAceptar.SetFocus
            End If

    End Select
    
    ' ******************************************************************************
End Sub

