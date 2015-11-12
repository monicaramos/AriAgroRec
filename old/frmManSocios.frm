VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Socios"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12870
   Icon            =   "frmManSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
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
      TabIndex        =   62
      Top             =   480
      Width           =   12465
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código Socio|N|N|1|999999|rsocios|codsocio|000000|S|"
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   3510
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||rsocios|nomsocio|||"
         Top             =   315
         Width           =   4995
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre "
         Height          =   255
         Left            =   2745
         TabIndex        =   64
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   63
         Top             =   315
         Width           =   495
      End
   End
   Begin VB.TextBox text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   9600
      TabIndex        =   83
      Top             =   765
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   59
      Top             =   6360
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
         TabIndex        =   60
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11670
      TabIndex        =   34
      Top             =   6510
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10380
      TabIndex        =   33
      Top             =   6510
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6930
      Top             =   6390
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
      TabIndex        =   67
      Top             =   0
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
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
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Baja de Socio"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Fases"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         TabIndex        =   68
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11670
      TabIndex        =   66
      Top             =   6510
      Visible         =   0   'False
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4980
      Left            =   240
      TabIndex        =   61
      Top             =   1290
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   8784
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   8
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
      TabPicture(0)   =   "frmManSocios.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label29"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgZoom(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgMail(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label19"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imgFec(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1(20)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FrameDatosDtoAdministracion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1(12)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(7)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame7"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Secciones"
      TabPicture(1)   =   "frmManSocios.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Teléfonos"
      TabPicture(2)   =   "frmManSocios.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux0"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Documentos"
      TabPicture(3)   =   "frmManSocios.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(1)=   "Text3(0)"
      Tab(3).Control(2)=   "Toolbar2"
      Tab(3).Control(3)=   "lw1"
      Tab(3).Control(4)=   "Toolbar3"
      Tab(3).Control(5)=   "Frame5"
      Tab(3).Control(6)=   "Toolbar4"
      Tab(3).Control(7)=   "imgFec(3)"
      Tab(3).Control(8)=   "Label17"
      Tab(3).Control(9)=   "Label16"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Pozos"
      TabPicture(4)   =   "frmManSocios.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameAux2"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   4395
         Left            =   -66120
         TabIndex        =   147
         Top             =   450
         Visible         =   0   'False
         Width           =   3465
         Begin VB.CommandButton cmdAccCRM 
            Height          =   375
            Index           =   0
            Left            =   0
            Picture         =   "frmManSocios.frx":0098
            Style           =   1  'Graphical
            TabIndex        =   150
            ToolTipText     =   "Insertar Imágen"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdAccCRM 
            Height          =   375
            Index           =   1
            Left            =   1080
            Picture         =   "frmManSocios.frx":0A9A
            Style           =   1  'Graphical
            TabIndex        =   149
            ToolTipText     =   "Ver Documento"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdAccCRM 
            Height          =   375
            Index           =   2
            Left            =   480
            Picture         =   "frmManSocios.frx":1024
            Style           =   1  'Graphical
            TabIndex        =   148
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1335
         Left            =   150
         TabIndex        =   141
         Top             =   790
         Width           =   5625
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   1275
            MaxLength       =   35
            TabIndex        =   4
            Tag             =   "Domicilio|T|N|||rsocios|dirsocio|||"
            Top             =   270
            Width           =   4215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1275
            MaxLength       =   6
            TabIndex        =   5
            Tag             =   "C.Postal|T|N|||rsocios|codpostal|||"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   2100
            MaxLength       =   30
            TabIndex        =   6
            Tag             =   "Población|T|N|||rsocios|pobsocio|||"
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   6
            Left            =   1275
            MaxLength       =   30
            TabIndex        =   7
            Tag             =   "Provincia|T|N|||rsocios|prosocio|||"
            Top             =   945
            Width           =   4200
         End
         Begin VB.Label Label6 
            Caption         =   "Dirección"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   144
            Top             =   270
            Width           =   735
         End
         Begin VB.Label Label28 
            Caption         =   "Provincia"
            Height          =   255
            Left            =   240
            TabIndex        =   143
            Top             =   1005
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   26
            Left            =   240
            TabIndex        =   142
            Top             =   630
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Domicilio de Correo"
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
         Height          =   1245
         Left            =   150
         TabIndex        =   137
         Top             =   2160
         Width           =   5655
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   25
            Left            =   1275
            MaxLength       =   35
            TabIndex        =   8
            Tag             =   "Domicilio|T|N|||rsocios|dirsociocorreo|||"
            Top             =   240
            Width           =   4215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   1275
            MaxLength       =   6
            TabIndex        =   9
            Tag             =   "C.Postal|T|N|||rsocios|codpostalcorreo|||"
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   23
            Left            =   2100
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Población|T|N|||rsocios|pobsociocorreo|||"
            Top             =   540
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   22
            Left            =   1275
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Provincia|T|N|||rsocios|prosociocorreo|||"
            Top             =   855
            Width           =   4200
         End
         Begin VB.Label Label6 
            Caption         =   "Dirección"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   140
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label26 
            Caption         =   "Provincia"
            Height          =   255
            Left            =   240
            TabIndex        =   139
            Top             =   885
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   138
            Top             =   570
            Width           =   735
         End
      End
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         Height          =   3990
         Left            =   -74955
         TabIndex        =   130
         Top             =   405
         Width           =   12360
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   1290
            MaxLength       =   9
            TabIndex        =   132
            Tag             =   "Acciones|N|N|||rsocios_pozos|acciones|##0.00||"
            Text            =   "Acciones"
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   1950
            MaxLength       =   30
            TabIndex        =   133
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
            Height          =   290
            Index           =   0
            Left            =   300
            MaxLength       =   6
            TabIndex        =   134
            Tag             =   "Código Socio|N|N|1|999999|rsocios_pozos|codsocio|000000|S|"
            Text            =   "Socio"
            Top             =   3420
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   750
            MaxLength       =   9
            TabIndex        =   131
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
            TabIndex        =   135
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
            Bindings        =   "frmManSocios.frx":1A26
            Height          =   3220
            Index           =   2
            Left            =   60
            TabIndex        =   136
            Top             =   510
            Width           =   7950
            _ExtentX        =   14023
            _ExtentY        =   5689
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
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   -64530
         TabIndex        =   122
         Text            =   "Text4"
         Top             =   1050
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   4365
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Nacimiento|F|S|||rsocios|fechanac|dd/mm/yyyy||"
         Top             =   525
         Width           =   1200
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   4020
         Left            =   -74955
         TabIndex        =   96
         Top             =   405
         Width           =   12165
         Begin VB.Frame Frame3 
            Height          =   735
            Left            =   45
            TabIndex        =   109
            Top             =   3285
            Width           =   12075
            Begin VB.TextBox txtAux2 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   8775
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   112
               Text            =   "nomiva"
               Top             =   270
               Width           =   3015
            End
            Begin VB.TextBox txtAux2 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               Left            =   4905
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   111
               Text            =   "nomCuenta Proveedor"
               Top             =   270
               Width           =   3285
            End
            Begin VB.TextBox txtAux2 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               Left            =   945
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   110
               Text            =   "nomCuenta Cliente"
               Top             =   270
               Width           =   2925
            End
            Begin VB.Label Label13 
               Caption         =   "Cta.Cliente"
               Height          =   315
               Left            =   90
               TabIndex        =   115
               Top             =   270
               Width           =   825
            End
            Begin VB.Label Label12 
               Caption         =   "Iva"
               Height          =   315
               Left            =   8370
               TabIndex        =   114
               Top             =   270
               Width           =   375
            End
            Begin VB.Label Label9 
               Caption         =   "Cta.Prov."
               Height          =   315
               Left            =   4185
               TabIndex        =   113
               Top             =   270
               Width           =   690
            End
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   5
            Left            =   9270
            TabIndex        =   108
            ToolTipText     =   "Buscar iva"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   6
            Left            =   8550
            MaxLength       =   2
            TabIndex        =   40
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
            Height          =   315
            Index           =   1
            Left            =   1845
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   107
            Text            =   "Nombre seccion"
            Top             =   2925
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   4
            Left            =   1665
            TabIndex        =   106
            ToolTipText     =   "Buscar fecha"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   1
            Left            =   6525
            TabIndex        =   104
            ToolTipText     =   "Buscar fecha"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   0
            Left            =   5760
            TabIndex        =   103
            ToolTipText     =   "Buscar fecha"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   5
            Left            =   7605
            MaxLength       =   10
            TabIndex        =   39
            Tag             =   "Cta.Contable Proveedor|T|S|||rsocios_seccion|codmacpro|||"
            Text            =   "cta provee"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   2
            Left            =   7380
            TabIndex        =   101
            ToolTipText     =   "Buscar cuenta"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   4
            Left            =   6660
            MaxLength       =   10
            TabIndex        =   38
            Tag             =   "Cta.Contable Cliente|T|S|||rsocios_seccion|codmaccli|||"
            Text            =   "cta cliente"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   3
            Left            =   8325
            TabIndex        =   100
            ToolTipText     =   "Buscar cuenta"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   5940
            MaxLength       =   10
            TabIndex        =   37
            Tag             =   "Fecha Baja|F|S|||rsocios_seccion|fecbaja|dd/mm/yyyy||"
            Text            =   "fecha baja"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   5175
            MaxLength       =   10
            TabIndex        =   36
            Tag             =   "Fecha Alta|F|N|||rsocios_seccion|fecalta|dd/mm/yyyy||"
            Text            =   "fecha alta"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   990
            MaxLength       =   3
            TabIndex        =   35
            Tag             =   "Seccion|N|N|||rsocios_seccion|codsecci|000|S|"
            Text            =   "seccion"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   225
            MaxLength       =   6
            TabIndex        =   99
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
            TabIndex        =   97
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
            Bindings        =   "frmManSocios.frx":1A3E
            Height          =   2790
            Index           =   1
            Left            =   45
            TabIndex        =   98
            Top             =   495
            Width           =   12030
            _ExtentX        =   21220
            _ExtentY        =   4921
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
      Begin VB.Frame Frame4 
         Caption         =   "Teléfonos"
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
         Height          =   975
         Left            =   150
         TabIndex        =   91
         Top             =   3450
         Width           =   5655
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   11
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Móvil|T|S|||rsocios|movsocio|||"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Telfno 2|T|S|||rsocios|telsoci2|||"
            Text            =   "123456789012345"
            Top             =   585
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Teléfono 3|T|S|||rsocios|telsoci3|||"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   8
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Teléfono 1|T|S|||rsocios|telsoci1|||"
            Text            =   "123456789012345"
            Top             =   225
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Número 3"
            Height          =   255
            Left            =   2805
            TabIndex        =   95
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label20 
            Caption         =   "Número 2"
            Height          =   255
            Left            =   225
            TabIndex        =   94
            Top             =   630
            Width           =   870
         End
         Begin VB.Label Label25 
            Caption         =   "Móvil"
            Height          =   255
            Left            =   2790
            TabIndex        =   93
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label41 
            Caption         =   "Número 1"
            Height          =   255
            Left            =   225
            TabIndex        =   92
            Top             =   270
            Width           =   915
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   1395
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "E-mail|T|S|||rsocios|maisocio|||"
         Top             =   4545
         Width           =   4185
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   3990
         Left            =   -74955
         TabIndex        =   75
         Top             =   405
         Width           =   12360
         Begin VB.TextBox txtAux 
            Height          =   735
            Index           =   16
            Left            =   7425
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   58
            Tag             =   "Observaciones|T|S|||rsocios_telefonos|observaciones|||"
            Top             =   3195
            Width           =   4755
         End
         Begin VB.TextBox txtAux 
            Height          =   285
            Index           =   6
            Left            =   8595
            MaxLength       =   35
            TabIndex        =   47
            Tag             =   "Direccion|T|S|||rsocios_telefonos|direccion|||"
            Top             =   405
            Width           =   3540
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Inactivo"
            Height          =   315
            Index           =   1
            Left            =   11160
            TabIndex        =   51
            Tag             =   "Inactivo|N|N|||rsocios_telefonos|inactivo||N|"
            Top             =   1485
            Width           =   945
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   360
            MaxLength       =   9
            TabIndex        =   42
            Tag             =   "Código telefono|T|N|||rsocios_telefono|idtelefono||S|"
            Text            =   "idtelefon"
            Top             =   3405
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   -90
            MaxLength       =   6
            TabIndex        =   41
            Tag             =   "Código Socio|N|N|1|999999|rsocios_telefonos|codsocio|000000|S|"
            Text            =   "cods"
            Top             =   3420
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   1575
            MaxLength       =   25
            TabIndex        =   44
            Tag             =   "IMEI|T|N|||rsocios_telefonos|imei|||"
            Text            =   "imei"
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   4
            Left            =   2925
            MaxLength       =   6
            TabIndex        =   45
            Tag             =   "Cod.Pobla|T|S|||rsocios_telefonos|codpostal|||"
            Text            =   "C.P."
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Height          =   290
            Index           =   7
            Left            =   8595
            MaxLength       =   30
            TabIndex        =   48
            Tag             =   "Poblacion|T|S|||rsocios_telefonos|poblacion|||"
            Top             =   765
            Width           =   3525
         End
         Begin VB.TextBox txtAux 
            Height          =   290
            Index           =   8
            Left            =   8595
            MaxLength       =   30
            TabIndex        =   49
            Tag             =   "Provincia|T|S|||rsocios_telefonos|provincia|||"
            Text            =   "Prov"
            Top             =   1140
            Width           =   3525
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   9315
            MaxLength       =   4
            TabIndex        =   55
            Tag             =   "Sucursal|N|S|0|9999|rsocios_telefonos|codsucur|0000||"
            Top             =   2610
            Width           =   630
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   10035
            MaxLength       =   2
            TabIndex        =   56
            Tag             =   "Digito Control|T|S|||rsocios_telefonos|digcontr|00||"
            Top             =   2610
            Width           =   495
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   10575
            MaxLength       =   10
            TabIndex        =   57
            Tag             =   "Cuenta Bancaria|T|S|||rsocios_telefonos|cuentaba|0000000000||"
            Top             =   2610
            Width           =   1575
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   900
            MaxLength       =   9
            TabIndex        =   43
            Tag             =   "NIF|T|S|||rsocios_telefonos|nif|||"
            Text            =   "nif"
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Height          =   290
            Index           =   12
            Left            =   8595
            MaxLength       =   4
            TabIndex        =   54
            Tag             =   "Banco|N|S|0|9999|rsocios_telefonos|codbanco|0000||"
            Top             =   2610
            Width           =   600
         End
         Begin VB.TextBox txtAux 
            Height          =   285
            Index           =   5
            Left            =   8595
            MaxLength       =   40
            TabIndex        =   46
            Tag             =   "Nombre|T|S|||rsocios_telefonos|nombre|||"
            Top             =   45
            Width           =   3540
         End
         Begin VB.TextBox txtAux 
            Height          =   285
            Index           =   11
            Left            =   8595
            MaxLength       =   50
            TabIndex        =   53
            Tag             =   "Mail|T|S|||rsocios_telefonos|mail|||"
            Top             =   2205
            Width           =   3540
         End
         Begin VB.TextBox txtAux 
            Height          =   285
            Index           =   10
            Left            =   8595
            MaxLength       =   25
            TabIndex        =   52
            Tag             =   "SIM|T|N|||rsocios_telefonos|sim|||"
            Text            =   "1234567890123456789012345"
            Top             =   1875
            Width           =   3510
         End
         Begin VB.TextBox txtAux 
            Height          =   285
            Index           =   9
            Left            =   8595
            MaxLength       =   10
            TabIndex        =   50
            Tag             =   "Teléfono|T|S|||rsocios_telefonos|telefono1|||"
            Text            =   "1234567890"
            Top             =   1515
            Width           =   1020
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   45
            TabIndex        =   76
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
            Bindings        =   "frmManSocios.frx":1A56
            Height          =   3195
            Index           =   0
            Left            =   45
            TabIndex        =   77
            Top             =   495
            Width           =   7350
            _ExtentX        =   12965
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
         Begin VB.Label Label8 
            Caption         =   "Dirección"
            Height          =   255
            Left            =   7425
            TabIndex        =   105
            Top             =   450
            Width           =   1140
         End
         Begin VB.Image imgZoom 
            Height          =   240
            Index           =   1
            Left            =   8640
            Tag             =   "-1"
            ToolTipText     =   "Zoom descripción"
            Top             =   2925
            Width           =   240
         End
         Begin VB.Label Label39 
            Caption         =   "Observaciones"
            Height          =   255
            Left            =   7425
            TabIndex        =   87
            Top             =   2970
            Width           =   1140
         End
         Begin VB.Label Label24 
            Caption         =   "Provincia"
            Height          =   255
            Left            =   7425
            TabIndex        =   86
            Top             =   1215
            Width           =   1230
         End
         Begin VB.Label Label2 
            Caption         =   "Población"
            Height          =   255
            Left            =   7425
            TabIndex        =   85
            Top             =   855
            Width           =   1230
         End
         Begin VB.Image imgMail 
            Height          =   240
            Index           =   1
            Left            =   8280
            Top             =   2250
            Width           =   240
         End
         Begin VB.Label Label38 
            Caption         =   "Banco"
            Height          =   255
            Left            =   7425
            TabIndex        =   82
            Top             =   2655
            Width           =   825
         End
         Begin VB.Label Label33 
            Caption         =   "Nombre"
            Height          =   255
            Left            =   7425
            TabIndex        =   81
            Top             =   90
            Width           =   1140
         End
         Begin VB.Label Label32 
            Caption         =   "Mail"
            Height          =   255
            Left            =   7425
            TabIndex        =   80
            Top             =   2250
            Width           =   330
         End
         Begin VB.Label Label30 
            Caption         =   "SIM"
            Height          =   255
            Left            =   7425
            TabIndex        =   79
            Top             =   1890
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "Teléfono"
            Height          =   255
            Left            =   7425
            TabIndex        =   78
            Top             =   1560
            Width           =   735
         End
      End
      Begin VB.Frame FrameDatosDtoAdministracion 
         Caption         =   "Datos Relacionados Dto.Administración"
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
         Height          =   3315
         Left            =   5880
         TabIndex        =   70
         Top             =   450
         Width           =   6420
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Emite Fact."
            Height          =   315
            Index           =   3
            Left            =   1830
            TabIndex        =   29
            Tag             =   "Emite Factura|N|N|||rsocios|emitefact||N|"
            Top             =   2880
            Width           =   1095
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   26
            Left            =   2145
            TabIndex        =   146
            Top             =   2310
            Width           =   4080
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   27
            Tag             =   "Código Globalgap|T|S|||rsocios|codigoggap|||"
            Top             =   2310
            Width           =   615
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Factura Interna ADV"
            Height          =   315
            Index           =   2
            Left            =   3060
            TabIndex        =   19
            Tag             =   "Fact.Interna ADV|N|N|0|1|rsocios|esfactadvinterna||N|"
            Top             =   690
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   4650
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Tag             =   "Tipo Relacion|N|N|0|2|rsocios|tiporelacion||N|"
            Top             =   2880
            Width           =   1590
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   21
            Left            =   2115
            TabIndex        =   117
            Top             =   1110
            Width           =   4080
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   1485
            MaxLength       =   2
            TabIndex        =   20
            Tag             =   "Código Cooperativa|N|N|0|99|rsocios|codcoope|00||"
            Top             =   1110
            Width           =   555
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Correo"
            Height          =   315
            Index           =   0
            Left            =   5340
            TabIndex        =   26
            Tag             =   "Correo|N|N|||rsocios|correo||N|"
            Top             =   1890
            Width           =   765
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Tag             =   "Tipo Produccion|N|N|0|3|rsocios|tipoprod||N|"
            Top             =   2880
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Tag             =   "Tipo IRPF|N|N|0|2|rsocios|tipoirpf||N|"
            Top             =   2880
            Width           =   1230
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   19
            Left            =   3405
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "Cuenta Bancaria|T|S|||rsocios|cuentaba|0000000000||"
            Text            =   "Text1"
            Top             =   1905
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   18
            Left            =   2850
            MaxLength       =   2
            TabIndex        =   24
            Tag             =   "Digito Control|T|S|||rsocios|digcontr|00||"
            Text            =   "Text1"
            Top             =   1905
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   17
            Left            =   2175
            MaxLength       =   4
            TabIndex        =   23
            Tag             =   "Sucursal|N|S|0|9999|rsocios|codsucur|0000||"
            Text            =   "Text1"
            Top             =   1905
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   16
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "Banco|N|S|0|9999|rsocios|codbanco|0000||"
            Text            =   "Text1"
            Top             =   1905
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   1485
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "Código Situacion|N|N|0|99|rsocios|codsitua|00||"
            Top             =   1500
            Width           =   555
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   2115
            TabIndex        =   71
            Top             =   1500
            Width           =   4080
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   1500
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "Fecha Baja|F|S|||rsocios|fechabaja|dd/mm/yyyy||"
            Top             =   720
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   1500
            MaxLength       =   10
            TabIndex        =   17
            Tag             =   "Fecha Alta|F|N|||rsocios|fechaalta|dd/mm/yyyy||"
            Top             =   330
            Width           =   1200
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   1500
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   2910
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1200
            ToolTipText     =   "Buscar globalgap"
            Top             =   2340
            Width           =   240
         End
         Begin VB.Label Label27 
            Caption         =   "Globalgap"
            Height          =   255
            Left            =   180
            TabIndex        =   145
            Top             =   2340
            Width           =   825
         End
         Begin VB.Label Label21 
            Caption         =   "Relación Cooperativa"
            Height          =   255
            Left            =   4650
            TabIndex        =   127
            Top             =   2640
            Width           =   1650
         End
         Begin VB.Label Label15 
            Caption         =   "Cooperativa"
            Height          =   255
            Left            =   180
            TabIndex        =   118
            Top             =   1170
            Width           =   885
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1200
            ToolTipText     =   "Buscar Cooperativa"
            Top             =   1140
            Width           =   240
         End
         Begin VB.Label Label14 
            Caption         =   "Documentos Alta/Baja/Transmisión"
            Height          =   255
            Left            =   3060
            TabIndex        =   116
            Top             =   330
            Width           =   2580
         End
         Begin VB.Image imgDoc 
            Height          =   465
            Index           =   1
            Left            =   5700
            ToolTipText     =   "Impresión Documento Alta/Baja"
            Top             =   240
            Width           =   510
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1200
            Picture         =   "frmManSocios.frx":1A6E
            ToolTipText     =   "Buscar fecha"
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1200
            Picture         =   "frmManSocios.frx":1AF9
            ToolTipText     =   "Buscar fecha"
            Top             =   330
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo Productor"
            Height          =   255
            Left            =   2880
            TabIndex        =   89
            Top             =   2640
            Width           =   1260
         End
         Begin VB.Label Label43 
            Caption         =   "Tipo IRPF"
            Height          =   255
            Left            =   180
            TabIndex        =   88
            Top             =   2640
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Banco Cliente"
            Height          =   195
            Index           =   21
            Left            =   180
            TabIndex        =   84
            Top             =   1950
            Width           =   1005
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1200
            ToolTipText     =   "Buscar Situación"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Label Label23 
            Caption         =   "Situación"
            Height          =   255
            Left            =   180
            TabIndex        =   74
            Top             =   1560
            Width           =   705
         End
         Begin VB.Label Label22 
            Caption         =   "Fecha Baja"
            Height          =   255
            Left            =   210
            TabIndex        =   73
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Fecha Alta"
            Height          =   255
            Left            =   210
            TabIndex        =   72
            Top             =   360
            Width           =   1035
         End
      End
      Begin VB.TextBox Text1 
         Height          =   675
         Index           =   20
         Left            =   5895
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Tag             =   "Observaciones|T|S|||rsocios|observaciones|||"
         Top             =   4170
         Width           =   6375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1410
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "NIF / CIF|T|N|||rsocios|nifsocio|||"
         Top             =   520
         Width           =   1200
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   3690
         Left            =   -74850
         TabIndex        =   119
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
               Object.ToolTipText     =   "Impresión Documentos"
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
         TabIndex        =   120
         Top             =   420
         Width           =   7875
         _ExtentX        =   13891
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
         NumItems        =   0
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   2370
         Left            =   -74850
         TabIndex        =   126
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
         Left            =   -65775
         TabIndex        =   129
         Top             =   1440
         Width           =   2760
         Begin VB.CheckBox Check2 
            Caption         =   "Imprimir Facturas"
            Height          =   240
            Left            =   135
            TabIndex        =   124
            Top             =   720
            Width           =   2220
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Detalle Albaranes"
            Height          =   240
            Left            =   135
            TabIndex        =   123
            Top             =   270
            Width           =   2220
         End
      End
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   390
         Left            =   -65595
         TabIndex        =   128
         Top             =   1575
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
         Left            =   -65010
         Picture         =   "frmManSocios.frx":1B84
         ToolTipText     =   "Buscar fecha"
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label17 
         Caption         =   "Desde"
         Height          =   255
         Left            =   -65730
         TabIndex        =   125
         Top             =   1050
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   -65790
         TabIndex        =   121
         Top             =   510
         Width           =   2865
      End
      Begin VB.Label Label10 
         Caption         =   "F.Nacim."
         Height          =   255
         Left            =   3195
         TabIndex        =   102
         Top             =   570
         Width           =   810
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   4050
         Picture         =   "frmManSocios.frx":1C0F
         ToolTipText     =   "Buscar fecha"
         Top             =   525
         Width           =   240
      End
      Begin VB.Label Label19 
         Caption         =   "E-mail"
         Height          =   255
         Left            =   390
         TabIndex        =   90
         Top             =   4575
         Width           =   495
      End
      Begin VB.Image imgMail 
         Height          =   240
         Index           =   0
         Left            =   1125
         Top             =   4545
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   7125
         Tag             =   "-1"
         ToolTipText     =   "Zoom descripción"
         Top             =   3870
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   5925
         TabIndex        =   69
         Top             =   3870
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "NIF"
         Height          =   255
         Left            =   390
         TabIndex        =   65
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
      TabIndex        =   151
      Top             =   6450
      Visible         =   0   'False
      Width           =   2715
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
' +-+- Autor: CÈSAR                    -+-+
' +-+- Menú: General-Clientes-Clientes -+-+
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

Dim vSeccion As CSeccion
Dim b As Boolean

Private BuscaChekc As String
Private Const CarpetaIMG = "ImgFicFT"

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean

'Cambio en cuentas de la contabilidad
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
            
        Case 1 'Impresión del documento
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

    Mens = "Va a proceder a eliminar la imágen de la lista correspondiente al socio. " & vbCrLf & vbCrLf & "¿ Desea continuar ?" & vbCrLf & vbCrLf
    
    If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Sql = "delete from rfichdocs where codsocio = " & DBSet(Text1(0).Text, "N") & " and codigo = " & Me.lw1.SelectedItem.SubItems(3)
        conn.Execute Sql
        
        CargaDatosLW
        
    End If
    Exit Sub

eEliminarImagen:
    MuestraError Err.Number, "Eliminar imágen", Err.Description
End Sub



Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
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
                End If
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
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
            
           menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
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

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    
    If Index = 0 And (Modo = 3 Or Modo = 4) Then
        chkAbonos(3).Enabled = (Combo1(0).ListIndex = 2)
        If chkAbonos(3).Enabled = False Then chkAbonos(3).Value = 0
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
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
    
    If DatosADevolverBusqueda = "" Then
        Set dbAriagro = Nothing
    End If
    
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 18 'index del botó "primero"
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
        'el 10  son separadors
        .Buttons(11).Image = 20  'Baja de Socio
        
        .Buttons(13).Image = 10  'Imprimir
        .Buttons(14).Image = 25  'Imprimir
        .Buttons(15).Image = 11  'Eixir
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
    'La nevegacion para entradas, facturas....
    ImagenesNavegacion
   'Ponemos los datos del listview
    imgFec(3).Tag = vParam.FecIniCam
    Check1.Value = 1
    Check2.Value = 1
    CargaColumnas 0

    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i
    
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
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rsocios"
    Ordenacion = " ORDER BY codsocio"
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codsocio=-1"
    Data1.Refresh
       
    ' ******* si n'hi han llinies en datagrid *******
'    ReDim CadAncho(DataGridAux.Count) 'redimensione l'array a la quantitat de datagrids
'    CadAncho(0) = False
'    CadAncho(1) = False
'    CadAncho(2) = False
'    CadAncho(4) = False
    
    ModoLineas = 0
       
    ' **** si n'hi ha algun frame que no te datagrids ***
'    CargaFrame 3, False
    ' *************************************************
         
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo
    ' ************************************************
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
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
        Me.chkVistaPrevia(0).Value = 0
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
    
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    Me.chkAbonos(0).Value = 0
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
    Next i
    For i = 0 To chkAbonos.Count - 1
        Me.chkAbonos(i).Value = 0
    Next i
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************

    Me.Label31.Caption = ""

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
    For i = 0 To 3
        BloquearChk Me.chkAbonos(i), (Modo = 0 Or Modo = 2 Or Modo = 5)
    Next i
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
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
    ' *** si n'hi han llínies i imagens de buscar que no estiguen als grids ******
    'Llínies Departaments
    b = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
'    BloquearImage imgBuscar(3), Not b
'    BloquearImage imgBuscar(4), Not b
'    BloquearImage imgBuscar(7), Not b
'    imgBuscar(3).Enabled = b
'    imgBuscar(3).visible = b
    ' ****************************************************************************
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
        CargaGrid 2, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = b
    DataGridAux(2).Enabled = b
'    ' ****** si n'hi han combos a la capçalera ***********************
'    If (Modo = 0) Or (Modo = 2) Or (Modo = 4) Or (Modo = 5) Then
'        Combo1(0).Enabled = False
'        Combo1(0).BackColor = &H80000018 'groc
'    ElseIf (Modo = 1) Or (Modo = 3) Then
'        Combo1(0).Enabled = True
'        Combo1(0).BackColor = &H80000005 'blanc
'    End If
'    ' ****************************************************************
    
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
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
    
    b = (Modo = 5) And NumTabMto = 2
    For i = 1 To 3
        BloquearTxt txtAux3(i), Not b
    Next i
    b = (Modo = 5) And NumTabMto = 2 And ModoLineas = 2
    BloquearTxt txtAux3(1), b
    
     '-----------------------------
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
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b And DatosADevolverBusqueda = ""
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And DatosADevolverBusqueda = "" 'And Not DeConsulta and DatosADevolverBusqueda = ""
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    'dar de baja un socio
    Toolbar1.Buttons(11).Enabled = b
    Me.mnBajaSocio.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(13).Enabled = b
'    Toolbar1.Buttons(14).Enabled = b
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
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
        Case 0 'telefonos
            Tabla = "rsocios_telefonos"
            Sql = "SELECT rsocios_telefonos.codsocio, rsocios_telefonos.idtelefono, rsocios_telefonos.nif, "
            Sql = Sql & " rsocios_telefonos.imei, rsocios_telefonos.codpostal, rsocios_telefonos.nombre, "
            Sql = Sql & " rsocios_telefonos.direccion, rsocios_telefonos.poblacion, rsocios_telefonos.provincia, "
            Sql = Sql & " rsocios_telefonos.telefono1, rsocios_telefonos.sim, rsocios_telefonos.mail, rsocios_telefonos.codbanco, "
            Sql = Sql & " rsocios_telefonos.codsucur, rsocios_telefonos.digcontr, rsocios_telefonos.cuentaba, "
            Sql = Sql & " rsocios_telefonos.observaciones,  rsocios_telefonos.inactivo "
            Sql = Sql & " FROM " & Tabla
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE codsocio = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".idtelefono "
            
            
       Case 1 ' secciones
            Tabla = "rsocios_seccion"
             Sql = "SELECT rsocios_seccion.codsocio, rsocios_seccion.codsecci, rseccion.nomsecci, rsocios_seccion.fecalta, "
             Sql = Sql & " rsocios_seccion.fecbaja, rsocios_seccion.codmaccli, rsocios_seccion.codmacpro, rsocios_seccion.codiva "
            Sql = Sql & " FROM " & Tabla & " INNER JOIN rseccion ON rsocios_seccion.codsecci = rseccion.codsecci "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE codsocio = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".codsecci "
            
       Case 2 ' pozos
            Tabla = "rsocios_pozos"
            Sql = "SELECT rsocios_pozos.codsocio, rsocios_pozos.numfases, rsocios_pozos.acciones, rsocios_pozos.observac "
            Sql = Sql & " FROM " & Tabla
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE codsocio = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".numfases "
            
            
            
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
            vCadena = "Sólo está activo si el socio es una entidad. En este caso, cuando esté" & vbCrLf & _
                      "marcado las todas las facturas que se generen para este socio, se" & vbCrLf & _
                      "generarán como contabilizadas con el numero que le corresponda del" & vbCrLf & _
                      "tipo de movimiento." & vbCrLf & vbCrLf & _
                      "Cuando se reciba el documento se podrá cambiar el nro de factura y " & vbCrLf & _
                      "contabilizar en un proceso específico. " & vbCrLf & vbCrLf
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgDoc_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
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
        
       menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
    
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
            frmZ.pTitulo = "Observaciones del Teléfono"
            frmZ.pValor = txtAux(indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco txtAux(indice)
    End Select
            
End Sub


Private Sub lw1_Click()

    '[Monica]20/06/2013: añadida la condicion, pq fallaba
    If CByte(RecuperaValor(lw1.Tag, 1)) = 5 Then
        CargarIMG lw1.SelectedItem.SubItems(2)
    End If
    
End Sub

Private Sub lw1_DblClick()
Dim Seleccionado As Long
Dim Cadena As String

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
        If lw1.SelectedItem.Text = "RCP" Or lw1.SelectedItem.Text = "RMP" Or lw1.SelectedItem.Text = "TAL" Or lw1.SelectedItem.Text = "RVP" Then
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
'            .Titulo = "Imágen " & Me.lw1.SelectedItem.SubItems(1)
'            .NumeroParametros = 1
'            .SoloImprimir = False
'            .EnvioEMail = False
'            .NombreRPT = "rImgDocs.rpt"
'
'            .Opcion = 2015
'            .Show vbModal
'        End With
'   End If

   Call ShellExecute(Me.hwnd, "Open", Me.lw1.SelectedItem.SubItems(2), "", "", 1)
   
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
        Case 11 ' Baja de socios
            mnBajaSocio_Click
        Case 13 'Imprimir
            mnImprimir_Click
        Case 14 'Imprimir
            mnImprimirFases_Click
        Case 15    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    Timer1.Enabled = False
    Label31.visible = False
    
    
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

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
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
    Dim Cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    Cad = ""
    Cad = Cad & ParaGrid(Text1(1), 45, "Nombre")
    Cad = Cad & ParaGrid(Text1(0), 10, "Cód.")
    Cad = Cad & ParaGrid(Text1(2), 15, "NIF")
    Cad = Cad & ParaGrid(Text1(8), 15, "Teléfono")
    Cad = Cad & ParaGrid(Text1(11), 15, "Móvil")
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "1|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Socios" ' ***** repasa açò: títol de BuscaGrid *****
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
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
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
    Text1(0).Text = SugerirCodigoSiguienteStr("rsocios", "codsocio")
    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    ' ***********************************************************
    Combo1_LostFocus (0)
    
    
    ' *** si n'hi han camps de descripció a la capçalera ***
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


    CargarValoresAnteriores Me, 1

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    Combo1_LostFocus (0)
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
    ' *********************************************************
End Sub

Private Sub BotonBajaSocio()

    '[Monica]19/12/2012: damos aviso si hay entradas esta campaña
    If HayEntradasSocio(Text1(0).Text) Then
        If MsgBox("Este socio tiene entradas esta campaña. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
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
    Cad = "¿Seguro que desea eliminar el Socio?"
    Cad = Cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    ' **************************************************************************
    
    'borrem
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
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
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    For i = 0 To DataGridAux.Count - 1
        If i <> 3 Then
            CargaGrid i, True
            If Not AdoAux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
        End If
    Next i
    ' *******************************************

    ' *** si n'hi han llínies sense datagrid ***
    ' ***************************************
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
'    txtAux2(22).Text = PonerNombreDeCod(txtAux(22), "poblacio", "despobla", "codpobla", "N")

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
    PonerOpcionesMenu   'Activar opcions de menú según nivell
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
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        ' ********************************************************
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
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
                        End Select
                    ' *** els tabs que no tenen datagrid ***
                    ElseIf NumTabMto = 3 Then
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
            
            If NumTabMto = 1 Then
                If Not vSeccion Is Nothing Then
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                End If
            End If
            
            TerminaBloquear

            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
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
                cadMen = "La cuenta bancaria del socio no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(16)
                    b = False
                End If
            End If
        End If
        
        If b Then
            If Text1(26).Text <> "" Then
                Sql = "select count(*) from rsocios where codsocio <> " & DBSet(Text1(0).Text, "N") & " and codigoggap = " & DBSet(Text1(26).Text, "T")
                
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Este código de GlobalGap ya está asignado a otro socio. Revise.", vbExclamation
                    PonerFoco Text1(26)
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

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
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


Private Function eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codsocio=" & Data1.Recordset!Codsocio
        ' ***********************************************************************
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rsocios_seccion " & vWhere
        
    conn.Execute "DELETE FROM rsocios_telefonos " & vWhere
        
    ' *******************************
        
    CargarUnSocio Data1.Recordset!Codsocio, "D"
        
    'Eliminar la CAPÇALERA
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
                    cadMen = "No existe la Situación: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
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
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
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
        
        Case 7 'Fecha no comprobaremos que esté dentro de campaña
            '[Monica]24/10/2013: no tenia que dar el mensaje de dentro de campaña
            PonerFormatoFecha Text1(Index) ', True
            
        Case 13, 14 'Fechas
            If Modo = 1 Then Exit Sub
            '[Monica]24/10/2013: aqui si que debe dar el mensaje de dentro de campaña
            PonerFormatoFecha Text1(Index), True
            
        Case 25 'tipo de movimiento
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
          
        Case 16, 17 'ENTIDAD Y SUCURSAL BANCARIA
            PonerFormatoEntero Text1(Index)
          
        Case 26
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = DevuelveDesdeBDNew(cAgro, "rglobalgap", "descripcion", "codigo", Text1(Index).Text, "T")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el código de GlobalGap: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
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
                            MsgBox "Este código de GlobalGap ya está asignado a otro socio. Revise.", vbExclamation
                            PonerFoco Text1(Index)
                        End If
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 5: KEYBusqueda KeyAscii, 0 'poblacion
                Case 7: KEYBusqueda KeyAscii, 1 'actividad
                Case 26: KEYBusqueda KeyAscii, 2 'codigo de globalgap
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
        Case 0 'telefonos
            Sql = "¿Seguro que desea eliminar el telefono?"
            Sql = Sql & vbCrLf & "Teléfono: " & AdoAux(Index).Recordset!idtelefono & " - " & AdoAux(Index).Recordset!imei
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
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
            Sql = "¿Seguro que desea eliminar la sección?"
            Sql = Sql & vbCrLf & "Sección: " & AdoAux(Index).Recordset!codsecci & " - " & AdoAux(Index).Recordset!nomsecci
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
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
            Sql = "¿Seguro que desea eliminar el registro?"
            Sql = Sql & vbCrLf & "Numero Fase: " & AdoAux(Index).Recordset!numfases
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
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
'            PonerCampos
            
        End If
        ' *** si n'hi han tabs sense datagrid ***
        If Index = 3 Then CargaFrame 3, True
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
        Case 0: vTabla = "rsocios_telefonos"
        Case 1: vTabla = "rsocios_seccion"
        Case 2: vTabla = "rsocios_pozos"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
                NumF = SugerirCodigoSiguienteStr(vTabla, "idtelefono", vWhere)
            Else
                NumF = ""
            End If
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
                Case 0 'cuentas
                    For i = 0 To txtAux.Count - 1
                        txtAux(i).Text = ""
                    Next i
                    txtAux(0).Text = Text1(0).Text 'codsocio
                    txtAux(1).Text = NumF 'idtelefono
                    PonerFoco txtAux(1)
                    
            End Select
         
         Case 1   'secciones
            ' *** canviar la clau primaria de les llínies,
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
                anc = anc + 210
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
            ' *** canviar la clau primaria de les llínies,
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
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            For i = 0 To txtAux3.Count - 1
                txtAux3(i).Text = ""
            Next i
            
            txtAux3(0).Text = Text1(0).Text 'codsocio
            txtAux3(1).Text = NumF 'numero de fase
            PonerFoco txtAux3(1)
         
         
        Case 3
        
            
            
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
                txtAux3(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
        
            CargarValoresAnteriores Me, 2, "FrameAux2"
        
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
            PonerFoco txtAux3(2)
    End Select
    ' ***************************************************************************************
End Sub

' ***** Si n'hi han combos *****
' per a seleccionar la opcio del combo quan estem modificant; només per a "si" i "no"
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
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
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
                txtAux3(jj).visible = b
                txtAux3(jj).Top = alto
            Next jj
    End Select
End Sub




Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
    ImprimirDocumentos
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2 'NIF
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            ValidarNIF txtAux(Index).Text
        
        Case 5 'NOMBRE
            If txtAux(Index).Text <> "" Then txtAux(Index).Text = UCase(txtAux(Index).Text)
    
        Case 12, 13 'ENTIDAD Y SUCURSAL BANCARIA
            PonerFormatoEntero txtAux(Index)
            
        Case 16
            cmdAceptar.SetFocus
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
Dim RS As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    
    If b And NumTabMto = 2 And ModoLineas = 1 Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "rsocios_pozos", "acciones", "codsocio", txtAux3(0).Text, "N", , "numfases", txtAux3(1).Text, "N")
        If Sql <> "" Then
            MsgBox "El número de fase ya existe. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtAux3(1)
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

' *** si n'hi han formularis de buscar codi a les llínies ***
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
                            
                            txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", CStr(AdoAux(1).Recordset!CodIVA), "N")
                            
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
        If (Index = 3) Then 'datos facturacion
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
        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
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
            tots = "N||||0|;S|txtAux(1)|T|Telefono|900|;" 'codsocio,idtelefono
            tots = tots & "S|txtAux(2)|T|NIF|1200|;"
            tots = tots & "S|txtAux(3)|T|IMEI|3050|;"
            tots = tots & "S|txtAux(4)|T|C.P|700|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla tots, DataGridAux(Index), Me
        
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
            tots = "N||||0|;S|txtaux1(1)|T|Cód.|800|;S|cmdAux(4)|B|||;" 'codsocio,codsecci
            tots = tots & "S|txtAux2(1)|T|Nombre|3700|;"
            tots = tots & "S|txtaux1(2)|T|F.Alta|1200|;S|cmdAux(0)|B|||;"
            tots = tots & "S|txtaux1(3)|T|F.Baja|1200|;S|cmdAux(1)|B|||;"
            tots = tots & "S|txtaux1(4)|T|Cta.Cliente|1200|;S|cmdAux(2)|B|||;"
            tots = tots & "S|txtaux1(5)|T|Cta.Prov.|1200|;S|cmdAux(3)|B|||;"
            tots = tots & "S|txtaux1(6)|T|Iva|800|;S|cmdAux(5)|B|||;"
            arregla tots, DataGridAux(Index), Me
        
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
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

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


Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'telefonos
        Case 1: nomframe = "FrameAux1" 'secciones
        Case 2: nomframe = "FrameAux2" 'pozos
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
                Case 0, 1, 2 ' *** els index de les llinies en grid (en o sense tab) ***
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
'Modifica registre en les taules de Llínies
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
                    
            
            
            
            If NumTabMto = 1 Then
                If Not vSeccion Is Nothing Then
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                End If
            End If
            
            ' ******************************************************
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModoLineas = 0

            If NumTabMto <> 3 Then
                V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
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
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codsocio=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' *** neteja els camps dels tabs de grid que
'estan fora d'este, i els camps de descripció ***
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
            For i = 0 To txtAux3.Count - 1
                txtAux3(i).Text = ""
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
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
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

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo irpf
    Combo1(0).AddItem "Módulos"
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
Dim Cadena As String
    
    If Not PerderFocoGnral(txtaux1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 1 ' seccion
                If PonerFormatoEntero(txtaux1(Index)) Then
                    Set vSeccion = New CSeccion
                    If vSeccion.LeerDatos(txtaux1(Index)) Then
                        txtAux2(Index).Text = vSeccion.Nombre
                        If vSeccion.AbrirConta Then
                        
                            ' si estamos insertando montamos las cuentas contables con las raices
                            ' y el codigo
                            
                            Cadena = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
                            
                            '18/09/2009
                            txtaux1(4).Text = vSeccion.RaizSocio & Format(txtaux1(0).Text, Cadena)
                            txtaux1(5).Text = vSeccion.RaizProv & Format(txtaux1(0).Text, Cadena)
                        End If
                    Else
                        Set vSeccion = Nothing
                        cadMen = "No existe la Sección: " & txtaux1(Index).Text & vbCrLf
                        cadMen = cadMen & "¿Desea crearla?" & vbCrLf
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
            cmdAceptar.SetFocus

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
    KEYpress KeyAscii
End Sub



'??????????????????????????
Private Sub TxtAux3_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Cadena As String
    
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 1 ' numfases
            PonerFormatoEntero txtAux3(Index)
            
        Case 2
            PonerFormatoDecimal txtAux3(Index), 10
        
        Case 3 'observaciones
            cmdAceptar.SetFocus

    End Select
    
    ' ******************************************************************************
End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
   If Not txtAux3(Index).MultiLine Then ConseguirFocoLin txtAux3(Index)
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux3(Index).MultiLine Then KEYdown KeyCode
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
        Ancho = "1200|1000|2400|1700|1500|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|"
        'Formatos
        Formato = "|0|0|0|###,###,##0|"
        Ncol = 5
        
    Case 2
        'Detalles de Entradas
        Label16.Caption = "Detalles de Entradas"
        If vParamAplic.Cooperativa = 12 Then
            Columnas = "Albarán|Fecha|Variedad|Cajas|Kilos Netos|"
            Ancho = "1200|1100|2000|1500|2000|"
            'vwColumnRight =1  left=0   center=2
            Alinea = "0|0|0|1|1|"
            'Formatos
            Formato = "|dd/mm/yyyy|0|###,##0|###,###,##0|"
            Ncol = 5
        Else
            Columnas = "Albarán|Fecha|Campo|N.Orden|Partida|Variedad|Kilos Netos|"
            Ancho = "900|1100|900|900|1400|1400|1200|"
            'vwColumnRight =1  left=0   center=2
            Alinea = "0|0|0|0|0|0|1|"
            'Formatos
            Formato = "|dd/mm/yyyy|0|0|0|0|###,###,##0|"
            Ncol = 7
        End If
        
    Case 4
        'Facturas adv , retirada almazara y bodega , recibos de pozos
        Label16.Caption = "Facturas Cliente"
        Columnas = "Tipo|Factura|Fecha|Importe|"
        Ancho = "1000|2000|1200|3600|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|"
        'Formatos
        Formato = "|00000000|dd/mm/yyyy|" & FormatoImporte & "|"
        Ncol = 4
        
    Case 5
        ' Documentos
        Label16.Caption = "Imágenes"
        Columnas = "Código|Nombre|Documento|Id|Tipo|"
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
Dim RS As ADODB.Recordset
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
        ' advfacturas
        Cad = "select h.codtipom,h.numfactu,h.fecfactu,h.totalfac totalfac from advfacturas h WHERE 1=1"
        Cad = Cad & " and h.codsocio=" & Data1.Recordset!Codsocio
        GroupBy = ""
        BuscaChekc = "h.fecfactu"
        'La fecha
        If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
        
        Cad = Cad & " union "
        
        'rbodfacturas
        Cad = Cad & "select i.codtipom,i.numfactu,i.fecfactu,i.totalfac totalfac from rbodfacturas i WHERE 1=1"
        Cad = Cad & " and i.codsocio=" & Data1.Recordset!Codsocio
        GroupBy = ""
        BuscaChekc = "i.fecfactu"
        'La fecha
        If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
        
        Cad = Cad & " union "
    
        'rrecibpozos
        Cad = Cad & "select j.codtipom,j.numfactu,j.fecfactu,j.totalfact totalfac from rrecibpozos j WHERE 1=1"
        Cad = Cad & " and j.codsocio=" & Data1.Recordset!Codsocio
        GroupBy = ""
        BuscaChekc = "j.fecfactu"
        'La fecha
        If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
        
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
    Cad = Cad & " ORDER BY " & BuscaChekc & " DESC"
    BuscaChekc = ""
    
    lw1.ListItems.Clear
    
    
    If CByte(RecuperaValor(lw1.Tag, 1)) = 5 Then
        
        CargarArchivos
    
    Else
    
        Set RS = New ADODB.Recordset
        RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            Set It = lw1.ListItems.Add()
            If lw1.ColumnHeaders(1).Tag <> "" Then
                It.Text = Format(RS.Fields(0), lw1.ColumnHeaders(1).Tag)
            Else
                It.Text = RS.Fields(0)
            End If
            'El resto de cmpos
            For NumRegElim = 2 To CInt(RecuperaValor(lw1.Tag, 2))
                If IsNull(RS.Fields(NumRegElim - 1)) Then
                    It.SubItems(NumRegElim - 1) = " "
                Else
                    If lw1.ColumnHeaders(NumRegElim).Tag <> "" Then
                        It.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lw1.ColumnHeaders(NumRegElim).Tag)
                    Else
                        It.SubItems(NumRegElim - 1) = RS.Fields(NumRegElim - 1)
                    End If
                End If
                
            Next
            It.SmallIcon = ElIcono
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    
    
    End If
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set RS = Nothing
    
End Sub

Private Sub CargarArchivos()
Dim c As String
Dim L As Long
Dim RS As ADODB.Recordset
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
'            It.Text = "Código " & Aux
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
Dim cadTABLA As String, cOrden As String
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
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
'    cadparam = cadparam & "pSocio=" & Data1.Recordset!codsocio & "|"
'    numParam = numParam + 1
'
    cadParam = cadParam & "Fecha=Date(""" & Text3(0).Text & """)|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pDetalleAlb=" & Check1.Value & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pImpFactura=" & Check2.Value & "|"
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
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    cadTitulo = "Resumen de Documentos Socio"
        
    LlamarImprimir
End Sub

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
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
    cadParam = ""
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
Dim vSocio As CSocio
Dim Cuentas As String
Dim Sql1 As String
Dim Sql2 As String
Dim NRegs As Long
Dim RS As ADODB.Recordset


    On Error GoTo eModificarDatosCuentaContable


    If Text1(1).Text <> NombreAnt Or Text1(16).Text <> BancoAnt Or Text1(17).Text <> SucurAnt Or Text1(18).Text <> DigitoAnt Or Text1(19).Text <> CuentaAnt Or _
       DirecAnt <> Text1(3).Text Or cPostalAnt <> Text1(4).Text Or PoblaAnt <> Text1(5).Text Or ProviAnt <> Text1(6).Text Or NifAnt <> Text1(2).Text Or EMaiAnt <> Text1(12).Text Then
        
        Cad = "Se han producido cambios en los siguientes datos del Socio: " & vbCrLf
        
        If NombreAnt <> Text1(1).Text Then Cad = Cad & " Nombre,"
        If DirecAnt <> Text1(3).Text Then Cad = Cad & " Direccion,"
        If cPostalAnt <> Text1(4).Text Then Cad = Cad & " CPostal,"
        If PoblaAnt <> Text1(5).Text Then Cad = Cad & " Población,"
        If ProviAnt <> Text1(6).Text Then Cad = Cad & " Provincia,"
        If NifAnt <> Text1(2).Text Then Cad = Cad & " NIF,"
        If EMaiAnt <> Text1(12).Text Then Cad = Cad & " EMail,"
        If BancoAnt <> Text1(16).Text Then Cad = Cad & " Banco,"
        If SucurAnt <> Text1(17).Text Then Cad = Cad & " Sucursal,"
        If DigitoAnt <> Text1(18).Text Then Cad = Cad & " Dig.Control,"
        If CuentaAnt <> Text1(19).Text Then Cad = Cad & " Cuenta banco,"
        
        Cad = Mid(Cad, 1, Len(Cad) - 1)
        
        Cad = Cad & vbCrLf & vbCrLf & "¿ Desea actualizar los datos en la Contabilidad de la Sección Horto ?" & vbCrLf & vbCrLf
        
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Set vSocio = New CSocio
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
                        Sql = Sql & ", entidad = " & DBSet(Trim(Text1(16).Text), "T")
                        Sql = Sql & ", oficina = " & DBSet(Trim(Text1(17).Text), "T")
                        Sql = Sql & ", cc = " & DBSet(Trim(Text1(18).Text), "T")
                        Sql = Sql & ", cuentaba = " & DBSet(Trim(Text1(19).Text), "T")
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
                        If BancoAnt <> Text1(16).Text Or SucurAnt <> Text1(17).Text Or DigitoAnt <> Text1(18).Text Or CuentaAnt <> Text1(19).Text _
                           And Cuentas <> "" Then
                            Sql1 = "select sum(total) from ("
                            Sql1 = Sql1 & "select  count(*) total "
                            Sql1 = Sql1 & " from scobro  cc "
                            Sql1 = Sql1 & " where cc.codmacta  in (" & Cuentas & ")  and (cc.codrem is null or cc.codrem = 0) and (cc.transfer is null or cc.transfer = 0)"
                            Sql1 = Sql1 & " union "
                            Sql1 = Sql1 & " select count(*) total "
                            Sql1 = Sql1 & " from spagop pp "
                            Sql1 = Sql1 & " where pp.ctaprove in (" & Cuentas & ") and  (pp.transfer is null or pp.transfer = 0)"
                            Sql1 = Sql1 & ") aaaaaa "

                            NRegs = 0

                            Set RS = New ADODB.Recordset
                            RS.Open Sql1, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                            If Not RS.EOF Then
                                If DBLet(RS.Fields(0).Value) <> 0 Then NRegs = RS.Fields(0).Value
                            End If
                            Set RS = Nothing
                            
                            If NRegs <> 0 Then
                                Cad = "Se han producido cambios en la Cta.Bancaria del Socio."
                                Cad = Cad & vbCrLf & vbCrLf & "¿ Desea actualizar los Cobros y Pagos pendientes en Tesoreria ?" & vbCrLf & vbCrLf
                                
                                If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                                     Sql2 = "update scobro set codbanco = " & DBSet(Text1(16).Text, "N", "S") & ", codsucur = " & DBSet(Text1(17).Text, "N", "S")
                                     Sql2 = Sql2 & ", digcontr = " & DBSet(Text1(18).Text, "T", "S") & ", cuentaba = " & DBSet(Text1(19).Text, "T", "S")
                                     Sql2 = Sql2 & " where codmacta in (" & Cuentas & ") "
                                     Sql2 = Sql2 & " and (codrem is null or codrem = 0) and (transfer is null or transfer = 0)"
                                     
                                     ConnConta.Execute Sql2
                                     
                                     Sql2 = "update spagop set entidad = " & DBSet(Text1(16).Text, "T", "S") & ", oficina = " & DBSet(Text1(17).Text, "T", "S")
                                     Sql2 = Sql2 & ", cc = " & DBSet(Text1(18).Text, "T", "S") & ", cuentaba = " & DBSet(Text1(19).Text, "T", "S")
                                     Sql2 = Sql2 & " where ctaprove in (" & Cuentas & ") "
                                     Sql2 = Sql2 & " and (transfer is null or transfer = 0)"
                                    
                                     ConnConta.Execute Sql2
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

